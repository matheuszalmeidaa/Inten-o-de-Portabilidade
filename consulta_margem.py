#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consulta de Margem — Multiplus (multiplus.consignadorapido.com)

Lê os CPFs da coluna "CPF Cliente" de uma planilha (Google Sheets ou arquivo
local), consulta cada um na tela CONSULTAS do Multiplus e coleta:

    CPF, NB (matrícula/benefício), Nome, Idade, Margem 40,
    Margem disponível para Cartão, Valor Benefício, Situação, Espécie

Trata o caso "CPF com mais de uma Matrícula": quando o modal aparece, o robô
consulta TODOS os NBs listados, um por vez (escolhe o NB, clica OK, coleta os
dados, consulta o CPF de novo para abrir o modal e pegar o próximo NB).

Os resultados são gravados linha a linha em CSV (e no final em .xlsx), então
mesmo que a execução seja interrompida no meio, o que já foi consultado fica
salvo.

Uso rápido (ver README.md para detalhes):
    python consulta_margem.py --limite 2          # teste com os 2 primeiros CPFs
    python consulta_margem.py                     # roda a planilha inteira
    python consulta_margem.py --cpf 01715717970   # testa um CPF específico
"""

import argparse
import csv
import os
import re
import sys
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()  # .env da pasta atual
    load_dotenv(Path(__file__).resolve().parent / ".env")  # .env ao lado do script
except ImportError:
    pass

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ===========================================================================
#  CONFIGURAÇÃO — EDITE AQUI
#
#  Preencha os valores entre as aspas. O que ficar em branco ("") o robô
#  tenta buscar num arquivo .env na mesma pasta (opcional).
#
#  ⚠ ATENÇÃO: com a senha preenchida aqui, NUNCA envie este arquivo para o
#    GitHub nem compartilhe com quem não deva ter acesso.
# ===========================================================================
MULTIPLUS_USUARIO = ""            # ex.: "KALLYANEOPERACIONAL"
MULTIPLUS_SENHA = ""              # ex.: "minha_senha"
PLANILHA_URL = ("https://docs.google.com/spreadsheets/d/"
                "1Fln-LwllZyS3SRgBb1psLKmdLRszxoZYV0waD9AyvOM/edit?gid=0#gid=0")
COLUNA_CPF = "CPF Cliente"        # coluna da planilha que tem os CPFs
MULTIPLUS_URL = "https://multiplus.consignadorapido.com"
HEADLESS = False                  # True = não abre a janela do navegador
TIMEOUT_CONSULTA = 90             # segundos de espera por consulta
PAUSA_ENTRE_CONSULTAS = 1.0       # pausa (s) entre um CPF e outro
# ===========================================================================
#  Fim da configuração — daqui para baixo não precisa mexer.
# ===========================================================================


def _cfg(valor_no_script, nome_env, padrao=""):
    """O valor preenchido no script tem prioridade; em branco, usa o .env."""
    v = str(valor_no_script).strip() if valor_no_script is not None else ""
    return v if v else os.getenv(nome_env, padrao)


BASE_URL = _cfg(MULTIPLUS_URL, "MULTIPLUS_URL",
                "https://multiplus.consignadorapido.com").rstrip("/")
CONSULTA_URL = BASE_URL + "/consulta/consultas"

DIR_RESULTADOS = Path("resultados")
DIR_DEBUG = Path("debug")

CAMPOS_SAIDA = [
    "CPF", "NB", "Nome", "Idade", "Margem 40", "Margem Cartão",
    "Valor Benefício", "Situação", "Espécie", "Status", "Consultado em",
]

# Rótulos exatamente como aparecem no card de resultado do Multiplus
LABELS_CARD = [
    "Nome", "Idade", "Margem 40", "Margem disponível para Cartão",
    "Valor Benefício", "Situação", "Espécie",
]

# Mensagens de dialog nativo (alert/confirm) capturadas durante a consulta
_mensagens_dialog = []

# Últimos eventos de rede (para diagnóstico de falhas de login/consulta)
_rede_registros = []


# ---------------------------------------------------------------------------
# Leitura dos CPFs (Google Sheets ou arquivo local)
# ---------------------------------------------------------------------------

def so_digitos(texto):
    return re.sub(r"\D", "", str(texto or ""))


def normalizar_cpf(valor):
    """Converte '017.157.179-70' -> '01715717970'. Devolve None se inválido."""
    dig = so_digitos(valor)
    if not dig:
        return None
    if len(dig) < 11:
        dig = dig.zfill(11)  # Excel costuma comer zeros à esquerda
    return dig if len(dig) == 11 else None


def extrair_id_gid_da_url(url):
    m_id = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    m_gid = re.search(r"[#?&]gid=(\d+)", url)
    return (m_id.group(1) if m_id else None,
            int(m_gid.group(1)) if m_gid else 0)


def baixar_planilha_google(url, gid=None):
    """Baixa a aba da planilha como CSV via link de exportação.

    Requer que a planilha esteja compartilhada como "Qualquer pessoa com o
    link – Leitor" (ou publicada). Se estiver privada, baixe o CSV manualmente
    e use --arquivo.
    """
    sheet_id, gid_url = extrair_id_gid_da_url(url)
    if not sheet_id:
        raise ValueError(f"Não consegui extrair o ID da planilha da URL: {url}")
    gid_final = gid if gid is not None else gid_url
    export = (f"https://docs.google.com/spreadsheets/d/{sheet_id}/export"
              f"?format=csv&gid={gid_final}")
    resp = requests.get(export, timeout=60)
    tipo = resp.headers.get("content-type", "")
    if resp.status_code != 200 or "text/html" in tipo:
        raise RuntimeError(
            "Não consegui baixar a planilha do Google (ela deve estar privada).\n"
            "Soluções:\n"
            "  1) Na planilha: Compartilhar > Qualquer pessoa com o link > Leitor; ou\n"
            "  2) Baixe a aba como CSV (Arquivo > Fazer download > CSV) e rode com:\n"
            "     python consulta_margem.py --arquivo caminho/do/arquivo.csv"
        )
    return pd.read_csv(BytesIO(resp.content), dtype=str)


def carregar_dataframe(origem, gid=None):
    if os.path.exists(origem):
        ext = Path(origem).suffix.lower()
        if ext in (".xlsx", ".xls"):
            return pd.read_excel(origem, dtype=str)
        # CSV: detecta separador , ou ; automaticamente
        return pd.read_csv(origem, dtype=str, sep=None, engine="python")
    if "docs.google.com" in origem or re.fullmatch(r"[a-zA-Z0-9_-]{20,}", origem):
        url = origem if origem.startswith("http") else \
            f"https://docs.google.com/spreadsheets/d/{origem}/edit"
        return baixar_planilha_google(url, gid=gid)
    raise ValueError(f"Origem de CPFs não reconhecida (nem arquivo, nem URL): {origem}")


def encontrar_coluna_cpf(df, nome_coluna):
    """Acha a coluna de CPF mesmo se o nome vier duplicado ('CPF Cliente.1')."""
    alvo = nome_coluna.strip().lower()
    colunas = [str(c) for c in df.columns]
    for c in colunas:                       # nome exato
        if c.strip().lower() == alvo:
            return c
    for c in colunas:                       # duplicadas: 'CPF Cliente.1' etc.
        if c.strip().lower().startswith(alvo):
            return c
    for c in colunas:                       # qualquer coluna com 'cpf' no nome
        if "cpf" in c.strip().lower():
            return c
    raise ValueError(
        f"Coluna '{nome_coluna}' não encontrada. Colunas disponíveis: {colunas}"
    )


def carregar_cpfs(origem, nome_coluna, gid=None, manter_duplicados=False):
    df = carregar_dataframe(origem, gid=gid)
    coluna = encontrar_coluna_cpf(df, nome_coluna)
    cpfs, invalidos = [], 0
    for valor in df[coluna].tolist():
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            continue
        if not str(valor).strip():
            continue
        cpf = normalizar_cpf(valor)
        if cpf:
            cpfs.append(cpf)
        else:
            invalidos += 1
    total_bruto = len(cpfs)
    if not manter_duplicados:
        cpfs = list(dict.fromkeys(cpfs))  # remove duplicados mantendo a ordem
    print(f"[planilha] coluna usada: '{coluna}' | {total_bruto} CPFs válidos"
          f" | {len(cpfs)} únicos | {invalidos} inválidos/ignorados")
    return cpfs


# ---------------------------------------------------------------------------
# Helpers de navegador
# ---------------------------------------------------------------------------

JS_VISIVEL = """
    const vis = (el) => {
        if (!el || !el.getClientRects) return false;
        if (!el.getClientRects().length) return false;
        const st = window.getComputedStyle(el);
        return st.visibility !== 'hidden' && st.display !== 'none';
    };
"""


def clicar_por_texto(page, padrao_regex):
    """Clica em botão/link visível cujo texto casa com o regex (ex.: '^Consultar$')."""
    try:
        page.get_by_role("button", name=re.compile(padrao_regex, re.I)).first.click(timeout=4000)
        return True
    except Exception:
        pass
    achou = page.evaluate(
        "(padrao) => {" + JS_VISIVEL + """
        const re = new RegExp(padrao, 'i');
        const els = Array.from(document.querySelectorAll(
            "button, a, input[type='button'], input[type='submit']")).filter(vis);
        const alvo = els.find(e => re.test((e.innerText || e.value || '').trim()));
        if (alvo) { alvo.click(); return true; }
        return false;
        }""",
        padrao_regex,
    )
    if not achou:
        raise RuntimeError(f"Botão com texto /{padrao_regex}/ não encontrado na tela.")
    return True


def registrar_dialogos(page):
    """Aceita dialogs nativos (alert/confirm) e guarda a mensagem para o log."""
    def _handler(dialog):
        _mensagens_dialog.append(dialog.message)
        try:
            dialog.accept()
        except Exception:
            pass
    page.on("dialog", _handler)


def registrar_rede(page):
    """Guarda as últimas respostas/falhas de rede para sair no diagnóstico."""
    def _resposta(resp):
        try:
            if resp.request.resource_type in ("document", "xhr", "fetch"):
                _rede_registros.append(
                    (resp.request.method, resp.status, resp.url[:160], resp))
                del _rede_registros[:-20]
        except Exception:
            pass

    def _falha(req):
        try:
            if req.resource_type in ("document", "xhr", "fetch"):
                _rede_registros.append(
                    (req.method, f"FALHOU({req.failure})", req.url[:160], None))
                del _rede_registros[:-20]
        except Exception:
            pass

    page.on("response", _resposta)
    page.on("requestfailed", _falha)


# campo "Digite a Matrícula ou CPF" da tela de consulta.
# ":visible" é essencial: a página tem campos duplicados escondidos e, sem
# isso, o robô mira no escondido e espera para sempre.
SELETOR_CAMPO_CPF = ("input[placeholder*='cpf' i]:visible, "
                     "input[placeholder*='matr' i]:visible")


def _esperar_pagina_assentar(page, timeout_ms=15000):
    """Espera a página (SPA) terminar de baixar/renderizar, sem travar."""
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass


def diagnostico_pagina(page, nome):
    """Salva print + detalhes da tela em debug/ para investigar problemas."""
    try:
        DIR_DEBUG.mkdir(parents=True, exist_ok=True)
        ts = f"{datetime.now():%Y%m%d_%H%M%S}"
        png = DIR_DEBUG / f"{nome}_{ts}.png"
        page.screenshot(path=str(png), full_page=True)
        info = [f"URL: {page.url}", f"Titulo: {page.title()}", "", "Campos visiveis:"]
        try:
            campos = page.evaluate(
                "() => {" + JS_VISIVEL + """
                return Array.from(document.querySelectorAll('input, select, button'))
                    .filter(vis).slice(0, 30)
                    .map(e => e.tagName + ' type=' + (e.type || '') +
                              ' placeholder=' + (e.placeholder || '') +
                              ' id=' + (e.id || '') +
                              ' texto=' + ((e.innerText || e.value || '').trim().slice(0, 40)));
                }"""
            )
            info.extend("  " + c for c in campos)
        except Exception:
            info.append("  (não consegui listar)")
        try:
            iframes = page.evaluate(
                "() => Array.from(document.querySelectorAll('iframe'))"
                ".map(f => (f.src || f.title || 'sem src').slice(0, 120))")
            if iframes:
                info.append("")
                info.append("Iframes na pagina (possivel captcha):")
                info.extend("  " + s for s in iframes)
        except Exception:
            pass
        if _rede_registros:
            info.append("")
            info.append("Rede (ultimas respostas do servidor):")
            for metodo, status, url, resp in list(_rede_registros)[-15:]:
                linha = f"  {metodo} {status} {url}"
                if resp is not None:
                    try:
                        interessante = (isinstance(status, int) and status >= 400) \
                            or re.search(r"login|auth|entrar|logar", url, re.I)
                        if interessante:
                            corpo = (resp.text() or "").strip()
                            corpo = re.sub(r"\s+", " ", corpo)[:250]
                            if corpo:
                                linha += f" | corpo: {corpo}"
                    except Exception:
                        pass
                info.append(linha)
        try:
            texto = page.evaluate(
                "() => ((document.body && document.body.innerText) || '').slice(0, 1500)")
            info.append("")
            info.append("Texto da pagina:")
            info.append(texto)
        except Exception:
            pass
        txt = DIR_DEBUG / f"{nome}_{ts}.txt"
        txt.write_text("\n".join(info), encoding="utf-8")
        print(f"    [debug] print salvo:    {png}")
        print(f"    [debug] detalhes salvos: {txt}")
    except Exception as e:
        print(f"    [debug] não consegui salvar diagnóstico: {e}")


def _ha_senha_visivel(page):
    """Diz se há campo de senha visível na tela (checagem via JavaScript)."""
    try:
        return page.evaluate(
            "() => {" + JS_VISIVEL + """
            return Array.from(document.querySelectorAll(
                "#senha, input[type='password']")).some(vis);
            }"""
        )
    except Exception:
        return False


def fazer_login(page, usuario, senha):
    if "consignadorapido" not in page.url or "/consulta" in page.url:
        page.goto(BASE_URL, wait_until="domcontentloaded")
        _esperar_pagina_assentar(page)
    # campos da tela de login do Multiplus (ids reais + fallback genérico);
    # só considera campos visíveis — há duplicados escondidos na página
    senha_loc = page.locator(
        "#senha:visible, input[type='password']:visible").first
    usar_js = False
    try:
        senha_loc.wait_for(state="visible", timeout=45000)
    except PWTimeout:
        if _ha_senha_visivel(page):
            # o campo existe mas o Playwright não o considera clicável:
            # preenche direto via JavaScript
            usar_js = True
            print("[login] usando fallback JavaScript para preencher o login")
        else:
            # sem campo de senha: ou a sessão já está ativa, ou o site mostrou
            # outra coisa — registra diagnóstico e deixa a próxima etapa decidir
            print(f"[login] tela de login não apareceu (URL atual: {page.url})")
            diagnostico_pagina(page, "login_nao_apareceu")
            return
    if usar_js:
        page.evaluate(
            "(cred) => {" + JS_VISIVEL + """
            const u = Array.from(document.querySelectorAll(
                "#login, input[type='text'], input[type='email'], input:not([type])")).filter(vis)[0];
            const s = Array.from(document.querySelectorAll(
                "#senha, input[type='password']")).filter(vis)[0];
            for (const [el, valor] of [[u, cred.usuario], [s, cred.senha]]) {
                if (!el) continue;
                el.value = valor;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }
            }""",
            {"usuario": usuario, "senha": senha},
        )
    else:
        user_loc = page.locator(
            "#login:visible, input[placeholder*='login' i]:visible, "
            "input[type='text']:visible, input:not([type]):visible, "
            "input[type='email']:visible").first
        user_loc.fill(usuario)
        senha_loc.fill(senha)
    print("[login] usuário e senha preenchidos — clicando em Entrar...")
    alerta_antes = capturar_alerta(page)
    try:
        page.locator("#submit:visible").first.click(timeout=3000)
    except Exception:
        clicar_por_texto(page, r"^\s*Entrar\s*$")
    # senha sumiu = logou; apareceu aviso novo = site recusou o login
    inicio = time.time()
    fim = inicio + 60
    tentou_enter = False
    while time.time() < fim:
        if not _ha_senha_visivel(page):
            _esperar_pagina_assentar(page)
            print("[login] OK")
            return
        alerta = capturar_alerta(page)
        if alerta and alerta != alerta_antes:
            # mostra só o que APARECEU de novo (não o texto que já estava lá)
            novidade = alerta
            if alerta_antes and alerta_antes in alerta:
                novidade = alerta.replace(alerta_antes, "").strip(" |").strip()
            diagnostico_pagina(page, "login_recusado")
            raise RuntimeError(
                f"O site recusou o login: \"{novidade or alerta}\" — confira "
                "usuário e senha na seção CONFIGURAÇÃO do consulta_margem.py.")
        if not tentou_enter and time.time() - inicio > 15:
            print("[login] ainda na tela de login — tentando ENTER no campo de senha...")
            try:
                senha_loc.press("Enter", timeout=2000)
            except Exception:
                try:
                    page.keyboard.press("Enter")
                except Exception:
                    pass
            tentou_enter = True
        time.sleep(0.5)
    diagnostico_pagina(page, "login_falhou")
    raise RuntimeError(
        "Falha no login: cliquei em Entrar mas a tela de login não fechou em "
        "60s e o site não mostrou o motivo. Abra o login_falhou_*.txt em "
        "debug/ — a seção 'Rede' mostra o que o servidor respondeu. Confira "
        "também usuário e senha na seção CONFIGURAÇÃO do consulta_margem.py.")


def ir_para_consulta(page, usuario=None, senha=None):
    if "/consulta/consultas" not in page.url:
        page.goto(CONSULTA_URL, wait_until="domcontentloaded")
        _esperar_pagina_assentar(page)
    # se o site redirecionou para o login, autentica e volta
    precisa_login = _ha_senha_visivel(page)
    if precisa_login and usuario and senha:
        print("[sessão] caiu na tela de login — autenticando...")
        fazer_login(page, usuario, senha)  # se o login falhar, o erro sobe
        page.goto(CONSULTA_URL, wait_until="domcontentloaded")
        _esperar_pagina_assentar(page)
    try:
        page.locator(SELETOR_CAMPO_CPF).first.wait_for(
            state="visible", timeout=60000)
    except PWTimeout:
        diagnostico_pagina(page, "consulta_nao_carregou")
        raise RuntimeError(
            "A tela de Consultas não carregou (o campo 'Digite a Matrícula ou "
            f"CPF' não apareceu). URL atual: {page.url} — veja o print e o "
            f"arquivo de detalhes na pasta {DIR_DEBUG}/ e me envie."
        )


def garantir_logado_na_consulta(page, usuario, senha):
    """Garante sessão ativa e a tela de consulta aberta (reloga se caiu)."""
    ir_para_consulta(page, usuario, senha)


# ---------------------------------------------------------------------------
# Interação com a tela de consulta
# ---------------------------------------------------------------------------

def preencher_e_consultar(page, cpf):
    campo = page.locator(SELETOR_CAMPO_CPF).first
    campo.click()
    campo.fill(cpf)
    # Se existir o seletor de tipo (Benefício / CPF), garante que está em CPF.
    page.evaluate(
        "() => {" + JS_VISIVEL + """
        const sels = Array.from(document.querySelectorAll('select')).filter(vis);
        for (const s of sels) {
            const textos = Array.from(s.options).map(o => (o.textContent || '').trim());
            const temBeneficio = textos.some(t => /benef/i.test(t));
            const opCpf = Array.from(s.options).find(o => /^cpf$/i.test((o.textContent || '').trim()));
            if (temBeneficio && opCpf) {
                if (s.value !== opCpf.value) {
                    s.value = opCpf.value;
                    s.dispatchEvent(new Event('change', { bubbles: true }));
                }
                return;
            }
        }
        }"""
    )
    _mensagens_dialog.clear()
    clicar_por_texto(page, r"^\s*Consultar\s*$")


def modal_nb_visivel(page):
    """Detecta o modal 'CPF com mais de uma Matrícula.'"""
    return page.evaluate(
        "() => {" + JS_VISIVEL + """
        const own = (el) => Array.from(el.childNodes)
            .filter(n => n.nodeType === 3)
            .map(n => n.textContent).join(' ');
        return Array.from(document.querySelectorAll('body *'))
            .some(e => vis(e) && /mais de uma matr/i.test(own(e)));
        }"""
    )


def obter_opcoes_nb(page):
    """Lista as opções (value, texto) do select 'Selecione aqui o NB' do modal."""
    return page.evaluate(
        "() => {" + JS_VISIVEL + """
        const sels = Array.from(document.querySelectorAll('select')).filter(vis);
        const ehSelectDeTipo = (s) => Array.from(s.options)
            .some(o => /benef|offline|online/i.test(o.textContent || ''));
        let alvo = sels.find(s => s.options.length &&
            /selecione/i.test(s.options[0].textContent || ''));
        if (!alvo) {
            alvo = sels.find(s => !ehSelectDeTipo(s) &&
                s.closest("[class*='modal' i], [class*='dialog' i], [class*='bootbox' i], [class*='swal' i]"));
        }
        if (!alvo) return null;
        return Array.from(alvo.options)
            .filter(o => (o.value || '').trim() && !/selecione/i.test(o.textContent || ''))
            .map(o => ({ value: o.value, texto: (o.textContent || '').trim() }));
        }"""
    )


def escolher_nb_e_confirmar(page, valor):
    ok = page.evaluate(
        "(valor) => {" + JS_VISIVEL + """
        const sels = Array.from(document.querySelectorAll('select')).filter(vis);
        let alvo = sels.find(s => s.options.length &&
            /selecione/i.test(s.options[0].textContent || ''));
        if (!alvo) alvo = sels.find(s => Array.from(s.options).some(o => o.value === valor));
        if (!alvo) return false;
        alvo.value = valor;
        alvo.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
        }""",
        valor,
    )
    if not ok:
        raise RuntimeError("Não encontrei o select de NB no modal.")
    clicar_por_texto(page, r"^\s*OK\s*$")


def fechar_modal_se_aberto(page):
    """Fecha o modal de NB (Cancel/Esc) para não travar o próximo CPF."""
    try:
        if modal_nb_visivel(page):
            try:
                clicar_por_texto(page, r"^\s*Cancel(ar)?\s*$")
            except Exception:
                page.keyboard.press("Escape")
            time.sleep(0.5)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Extração dos dados do card de resultado
# ---------------------------------------------------------------------------

def _extrair_blocos(page):
    """Para cada rótulo do card, devolve as linhas de texto do seu container."""
    return page.evaluate(
        "(labels) => {" + JS_VISIVEL + """
        const norm = (s) => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();
        const own = (el) => Array.from(el.childNodes)
            .filter(n => n.nodeType === 3)
            .map(n => n.textContent).join(' ');
        const todos = Array.from(document.querySelectorAll('body *')).filter(vis);
        const out = {};
        for (const label of labels) {
            const alvo = norm(label);
            let el = todos.find(e => norm(own(e)) === alvo);
            if (!el) el = todos.find(e => norm(own(e)).startsWith(alvo));
            if (!el) { out[label] = null; continue; }
            // sobe até o menor container que tenha mais texto que o rótulo
            let cont = el, guard = 0;
            while (norm(cont.innerText).length <= alvo.length + 2 &&
                   cont.parentElement && guard++ < 6) {
                cont = cont.parentElement;
            }
            out[label] = (cont.innerText || '').split('\\n')
                .map(s => s.trim()).filter(Boolean);
        }
        return out;
        }""",
        LABELS_CARD,
    )


def _nb_da_aba_ativa(page):
    """Número do benefício da aba selecionada (melhor esforço)."""
    return page.evaluate(
        "() => {" + JS_VISIVEL + """
        const cands = Array.from(document.querySelectorAll(
            "[class*='active' i], [class*='selected' i], [aria-selected='true']")).filter(vis);
        for (const e of cands) {
            const t = (e.textContent || '').replace(/\\s+/g, '').trim();
            if (/^\\d{7,13}$/.test(t)) return t;
        }
        return null;
        }"""
    )


def _valor_do_bloco(label, linhas, tipo):
    """Extrai o valor das linhas do container de um rótulo."""
    if not linhas:
        return None
    baixo = [l.lower() for l in linhas]
    idx = baixo.index(label.lower()) if label.lower() in baixo else -1
    resto = linhas[idx + 1:] if idx >= 0 else list(linhas)
    outras = {l.lower() for l in LABELS_CARD if l.lower() != label.lower()}
    resto = [l for l in resto if l.lower() not in outras and l.strip() != "+"]
    if not resto:
        return None
    if tipo == "dinheiro":
        for l in resto:
            if "R$" in l:
                return l
        return None
    if tipo == "idade":
        for l in resto:
            if re.search(r"\d+\s*ano", l, re.I) or re.fullmatch(r"\d+", l):
                return l
        return resto[0]
    if tipo == "numero":
        for l in resto:
            if re.fullmatch(r"\d+", l):
                return l
        return resto[0]
    if tipo == "texto_longo":  # nome pode quebrar em várias linhas
        partes = [l for l in resto if "R$" not in l]
        return " ".join(partes) if partes else None
    return resto[0]


def extrair_dados(page):
    try:
        blocos = _extrair_blocos(page)
    except Exception:
        blocos = {}
    dados = {
        "nome": _valor_do_bloco("Nome", blocos.get("Nome"), "texto_longo"),
        "idade": _valor_do_bloco("Idade", blocos.get("Idade"), "idade"),
        "margem": _valor_do_bloco("Margem 40", blocos.get("Margem 40"), "dinheiro"),
        "margem_cartao": _valor_do_bloco(
            "Margem disponível para Cartão",
            blocos.get("Margem disponível para Cartão"), "dinheiro"),
        "valor_beneficio": _valor_do_bloco(
            "Valor Benefício", blocos.get("Valor Benefício"), "dinheiro"),
        "situacao": _valor_do_bloco("Situação", blocos.get("Situação"), "texto"),
        "especie": _valor_do_bloco("Espécie", blocos.get("Espécie"), "numero"),
    }
    try:
        dados["nb_aba"] = _nb_da_aba_ativa(page)
    except Exception:
        dados["nb_aba"] = None
    return dados


def capturar_alerta(page):
    """Texto de aviso/erro visível na tela (ex.: 'CPF não encontrado').

    Ignora indicadores de carregamento ('Aguarde...', 'Carregando...') para
    não confundir espera com erro, e junta os avisos distintos encontrados.
    """
    try:
        texto = page.evaluate(
            "() => {" + JS_VISIVEL + """
            const els = Array.from(document.querySelectorAll(
                "[class*='alert' i], [class*='toast' i], [class*='notif' i], [role='alert'], " +
                "[class*='swal' i], [class*='error' i], [class*='erro' i], " +
                "[class*='invalid' i], [class*='danger' i]"))
                .filter(vis);
            const textos = els
                .map(e => (e.innerText || '').replace(/\\s+/g, ' ').trim())
                .filter(t => t && t.length <= 250)
                .filter(t => !/aguarde|carregando|processando|loading/i.test(t));
            const unicos = Array.from(new Set(textos));
            return unicos.length ? unicos.join(' | ').slice(0, 300) : null;
            }"""
        )
    except Exception:
        texto = None
    if not texto and _mensagens_dialog:
        texto = _mensagens_dialog[-1][:200]
    return texto


# ---------------------------------------------------------------------------
# Espera pelo desfecho da consulta
# ---------------------------------------------------------------------------

def aguardar_desfecho(page, dados_antes, timeout_s):
    """Espera até: modal de NB abrir ('modal') ou o card mudar ('resultado')."""
    fim = time.time() + timeout_s
    while time.time() < fim:
        if modal_nb_visivel(page):
            return "modal"
        dados = extrair_dados(page)
        if dados.get("nome") and dados != dados_antes:
            time.sleep(1.2)  # deixa a tela terminar de renderizar
            return "resultado"
        time.sleep(0.7)
    return "timeout"


def aguardar_resultado_pos_modal(page, dados_antes, timeout_s):
    fim = time.time() + timeout_s
    while time.time() < fim:
        if not modal_nb_visivel(page):
            dados = extrair_dados(page)
            if dados.get("nome") and dados != dados_antes:
                time.sleep(1.2)
                return "resultado"
        time.sleep(0.7)
    return "timeout"


# ---------------------------------------------------------------------------
# Saída (CSV incremental + XLSX no final)
# ---------------------------------------------------------------------------

class Escritor:
    def __init__(self, caminho_csv):
        self.caminho_csv = Path(caminho_csv)
        self.caminho_csv.parent.mkdir(parents=True, exist_ok=True)
        novo = not self.caminho_csv.exists()
        self._arq = open(self.caminho_csv, "a", newline="", encoding="utf-8-sig")
        self._csv = csv.DictWriter(self._arq, fieldnames=CAMPOS_SAIDA, delimiter=";")
        if novo:
            self._csv.writeheader()
            self._arq.flush()
        self.linhas = 0

    def gravar(self, cpf, nb, dados, status):
        self._csv.writerow({
            "CPF": cpf,
            "NB": nb or dados.get("nb_aba") or "",
            "Nome": dados.get("nome") or "",
            "Idade": dados.get("idade") or "",
            "Margem 40": dados.get("margem") or "",
            "Margem Cartão": dados.get("margem_cartao") or "",
            "Valor Benefício": dados.get("valor_beneficio") or "",
            "Situação": dados.get("situacao") or "",
            "Espécie": dados.get("especie") or "",
            "Status": status,
            "Consultado em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        })
        self._arq.flush()
        self.linhas += 1

    def fechar_e_gerar_xlsx(self):
        self._arq.close()
        try:
            df = pd.read_csv(self.caminho_csv, sep=";", dtype=str, encoding="utf-8-sig")
            xlsx = self.caminho_csv.with_suffix(".xlsx")
            df.to_excel(xlsx, index=False)
            return xlsx
        except Exception:
            return None


def screenshot_debug(page, nome):
    try:
        DIR_DEBUG.mkdir(parents=True, exist_ok=True)
        caminho = DIR_DEBUG / f"{nome}_{datetime.now():%Y%m%d_%H%M%S}.png"
        page.screenshot(path=str(caminho), full_page=True)
        return caminho
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Fluxo por CPF
# ---------------------------------------------------------------------------

def processar_cpf(page, cpf, escritor, timeout_s, usuario, senha):
    garantir_logado_na_consulta(page, usuario, senha)
    fechar_modal_se_aberto(page)

    dados_antes = extrair_dados(page)
    preencher_e_consultar(page, cpf)
    desfecho = aguardar_desfecho(page, dados_antes, timeout_s)

    # ------- caso 1: modal "CPF com mais de uma Matrícula" -------
    if desfecho == "modal":
        opcoes = obter_opcoes_nb(page)
        if not opcoes:
            screenshot_debug(page, f"{cpf}_modal_sem_opcoes")
            fechar_modal_se_aberto(page)
            escritor.gravar(cpf, "", {}, "ERRO: modal abriu mas sem opções de NB")
            return
        print(f"    -> CPF com {len(opcoes)} matrículas: "
              + ", ".join(o["texto"] for o in opcoes))
        for i, opcao in enumerate(opcoes):
            if i > 0:
                # reconsulta o mesmo CPF para reabrir o modal e pegar o próximo NB
                dados_antes = extrair_dados(page)
                preencher_e_consultar(page, cpf)
                if aguardar_desfecho(page, dados_antes, timeout_s) != "modal":
                    screenshot_debug(page, f"{cpf}_modal_nao_reabriu")
                    escritor.gravar(cpf, opcao["texto"], {},
                                    "ERRO: modal não reabriu para este NB")
                    continue
            escolher_nb_e_confirmar(page, opcao["value"])
            resultado = aguardar_resultado_pos_modal(page, dados_antes, timeout_s)
            dados = extrair_dados(page)
            if resultado == "resultado":
                status = "OK"
            elif dados.get("nome") and dados != dados_antes:
                status = "OK (espera esgotou, mas os dados atualizaram)"
            elif dados.get("nome"):
                # a tela não mudou: o que está visível é do NB anterior
                dados = {}
                status = "VERIFICAR: tela não atualizou para este NB (print em debug/)"
                screenshot_debug(page, f"{cpf}_nb_{opcao['texto']}")
            else:
                status = "SEM DADOS (verificar print em debug/)"
                screenshot_debug(page, f"{cpf}_nb_{opcao['texto']}")
            escritor.gravar(cpf, opcao["texto"], dados, status)
            print(f"       NB {opcao['texto']}: {dados.get('nome') or '-'} | "
                  f"margem {dados.get('margem') or '-'} | {status}")
        return

    # ------- caso 2: resultado direto (uma matrícula só) -------
    dados = extrair_dados(page)
    mudou = dados.get("nome") and dados != dados_antes
    if desfecho == "resultado" or mudou:
        escritor.gravar(cpf, "", dados, "OK")
        print(f"    -> {dados.get('nome') or '-'} | margem {dados.get('margem') or '-'}"
              f" | situação {dados.get('situacao') or '-'}")
        return

    # ------- caso 3: nada aconteceu (CPF não encontrado / erro / lentidão) ---
    # a tela não mudou, então o card visível (se houver) é do CPF anterior:
    # grava a linha vazia para não misturar dados de outra pessoa.
    alerta = capturar_alerta(page)
    status = f"SEM RESULTADO: {alerta}" if alerta else \
        "SEM RESULTADO (tempo esgotado — ver print em debug/)"
    screenshot_debug(page, f"{cpf}_sem_resultado")
    escritor.gravar(cpf, "", {}, status)
    print(f"    -> {status}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args():
    p = argparse.ArgumentParser(
        description="Consulta margem no Multiplus a partir dos CPFs de uma planilha.")
    p.add_argument("--sheet", "--planilha", dest="sheet",
                   default=_cfg(PLANILHA_URL, "PLANILHA_URL") or None,
                   help="URL da planilha Google (padrão: PLANILHA_URL da configuração)")
    p.add_argument("--gid", type=int, default=None,
                   help="gid da aba da planilha (padrão: o da URL, ou 0 = Pagina1)")
    p.add_argument("--arquivo", help="CSV/XLSX local com os CPFs (alternativa ao --sheet)")
    p.add_argument("--coluna", default=_cfg(COLUNA_CPF, "COLUNA_CPF", "CPF Cliente"),
                   help="Nome da coluna com os CPFs (padrão: 'CPF Cliente')")
    p.add_argument("--cpf", action="append",
                   help="Consulta apenas este CPF (pode repetir a opção)")
    p.add_argument("--limite", type=int, default=None,
                   help="Consulta só os N primeiros CPFs (bom para testar)")
    p.add_argument("--manter-duplicados", action="store_true",
                   help="Não remove CPFs repetidos da planilha")
    p.add_argument("--headless", action="store_true",
                   default=bool(HEADLESS) or
                   os.getenv("HEADLESS", "false").lower() in ("1", "true", "sim"),
                   help="Roda sem abrir a janela do navegador")
    p.add_argument("--timeout", type=int,
                   default=int(float(_cfg(TIMEOUT_CONSULTA, "TIMEOUT_CONSULTA", "90"))),
                   help="Segundos de espera por consulta (padrão: 90)")
    p.add_argument("--pausa", type=float,
                   default=float(_cfg(PAUSA_ENTRE_CONSULTAS, "PAUSA_ENTRE_CONSULTAS", "1")),
                   help="Pausa em segundos entre consultas (padrão: 1)")
    p.add_argument("--saida", help="Caminho do CSV de saída")
    return p.parse_args()


def main():
    args = parse_args()

    usuario = _cfg(MULTIPLUS_USUARIO, "MULTIPLUS_USUARIO")
    senha = _cfg(MULTIPLUS_SENHA, "MULTIPLUS_SENHA")
    if not usuario or not senha:
        sys.exit("Preencha MULTIPLUS_USUARIO e MULTIPLUS_SENHA na seção "
                 "'CONFIGURAÇÃO — EDITE AQUI' no topo do consulta_margem.py "
                 "(ou num arquivo .env na mesma pasta).")

    # ----- lista de CPFs -----
    if args.cpf:
        cpfs = [c for c in (normalizar_cpf(v) for v in args.cpf) if c]
        if not cpfs:
            sys.exit("Nenhum CPF válido informado em --cpf.")
    else:
        origem = args.arquivo or args.sheet
        if not origem:
            sys.exit("Informe a origem dos CPFs: --sheet URL, --arquivo caminho.csv "
                     "ou PLANILHA_URL no .env")
        cpfs = carregar_cpfs(origem, args.coluna, gid=args.gid,
                             manter_duplicados=args.manter_duplicados)
    if args.limite:
        cpfs = cpfs[:args.limite]
    if not cpfs:
        sys.exit("Nenhum CPF para consultar.")

    # Em servidor Linux sem interface gráfica não dá para abrir a janela do
    # navegador: liga o modo invisível automaticamente.
    if (not args.headless and sys.platform.startswith("linux")
            and not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY")):
        print("[aviso] servidor sem tela detectado — rodando em modo invisível "
              "(--headless). Prints de erro continuam saindo na pasta debug/.")
        args.headless = True

    saida = args.saida or DIR_RESULTADOS / \
        f"resultados_margem_{datetime.now():%Y%m%d_%H%M%S}.csv"
    escritor = Escritor(saida)
    print(f"[saida] {escritor.caminho_csv}")
    print(f"[consulta] {len(cpfs)} CPFs | timeout {args.timeout}s | "
          f"navegador {'oculto' if args.headless else 'visível'}")

    erros = 0
    with sync_playwright() as pw:
        opcoes_chromium = {
            "headless": args.headless,
            "args": ["--disable-blink-features=AutomationControlled"],
        }
        try:
            # usa o Chromium completo (mais parecido com um navegador normal)
            browser = pw.chromium.launch(channel="chromium", **opcoes_chromium)
        except Exception:
            browser = pw.chromium.launch(**opcoes_chromium)

        # user-agent sem a marca "HeadlessChrome", que alguns sites bloqueiam
        ua = None
        try:
            sonda = browser.new_page()
            ua = sonda.evaluate("navigator.userAgent")
            sonda.close()
            ua = ua.replace("HeadlessChrome", "Chrome")
        except Exception:
            ua = None
        extras = {"user_agent": ua} if ua else {}
        contexto = browser.new_context(
            viewport={"width": 1440, "height": 900}, locale="pt-BR", **extras)
        page = contexto.new_page()
        page.set_default_timeout(30000)
        registrar_dialogos(page)
        registrar_rede(page)

        fazer_login(page, usuario, senha)
        ir_para_consulta(page, usuario, senha)

        try:
            for i, cpf in enumerate(cpfs, 1):
                print(f"[{i}/{len(cpfs)}] CPF {cpf}")
                for tentativa in (1, 2):
                    try:
                        processar_cpf(page, cpf, escritor,
                                      args.timeout, usuario, senha)
                        break
                    except KeyboardInterrupt:
                        raise
                    except Exception as e:
                        if tentativa == 1:
                            print(f"    !! erro ({e}); tentando de novo...")
                            fechar_modal_se_aberto(page)
                            try:
                                ir_para_consulta(page, usuario, senha)
                            except Exception:
                                pass
                        else:
                            erros += 1
                            screenshot_debug(page, f"{cpf}_erro")
                            escritor.gravar(cpf, "", {}, f"ERRO: {str(e)[:150]}")
                            print(f"    !! falhou de vez: {e}")
                time.sleep(args.pausa)
        except KeyboardInterrupt:
            print("\n[interrompido] resultados parciais já estão salvos no CSV.")
        finally:
            browser.close()

    xlsx = escritor.fechar_e_gerar_xlsx()
    print(f"\n[fim] {escritor.linhas} linha(s) gravada(s) | {erros} erro(s)")
    print(f"[fim] CSV : {escritor.caminho_csv}")
    if xlsx:
        print(f"[fim] XLSX: {xlsx}")


if __name__ == "__main__":
    main()
