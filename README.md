# Robô de Consulta de Margem — Multiplus

Automação que lê os CPFs da coluna **CPF Cliente** da planilha (aba *Pagina1*)
e consulta um por um na tela **CONSULTAS** do Multiplus
(`multiplus.consignadorapido.com`), coletando:

| Campo | Exemplo |
|---|---|
| CPF | 01715717970 |
| NB (matrícula/benefício) | 1082967022 |
| Nome | MARIA DE LOURDES ROCHA DE OLIVEIRA |
| Idade | 62 Anos |
| Margem 40 | R$ 36,05 |
| Margem Cartão | R$ 0,00 |
| Valor Benefício | R$ 1.621,00 |
| Situação | ATIVO |
| Espécie | 21 |

## Como o robô funciona

1. Faz login no Multiplus com o usuário/senha do arquivo `.env`;
2. Abre a tela `Consultas`, digita o CPF no campo *"Digite a Matrícula ou CPF"*
   e clica em **Consultar**;
3. **Se abrir o aviso "CPF com mais de uma Matrícula"**: o robô lê todas as
   opções do seletor *"Selecione aqui o NB"* e consulta **todas, uma por vez**
   — escolhe o NB, clica **OK**, anota os dados, consulta o mesmo CPF de novo
   para reabrir o aviso e pegar o próximo NB, até acabar;
4. Se a tela de margem abrir direto, apenas anota os dados e segue para o
   próximo CPF;
5. Grava cada resultado **na hora** em um CSV dentro da pasta `resultados/`
   (se o robô for interrompido, o que já foi consultado não se perde) e, ao
   final, gera também uma versão `.xlsx` para abrir no Excel.

CPFs repetidos na planilha são consultados **uma vez só** (use
`--manter-duplicados` se quiser repetir). CPFs sem resultado ou com erro ficam
marcados na coluna **Status**, com um print da tela salvo na pasta `debug/`.

## Instalação (uma vez só)

Requisito: [Python 3.10+](https://www.python.org/downloads/) — no Windows,
marque a opção **"Add Python to PATH"** ao instalar.

No terminal (Prompt de Comando), dentro da pasta do projeto:

```bash
pip install -r requirements.txt
playwright install chromium
```

## Configuração

1. Copie `.env.example` para `.env` (no Windows: `copy .env.example .env`);
2. Edite o `.env` e preencha `MULTIPLUS_USUARIO` e `MULTIPLUS_SENHA`;
3. A URL da planilha já vem preenchida. Para o robô conseguir ler a planilha
   direto do Google, ela precisa estar compartilhada como
   **"Qualquer pessoa com o link → Leitor"**. Se preferir manter privada,
   baixe a aba como CSV (*Arquivo → Fazer download → CSV*) e use `--arquivo`.

> ⚠️ O `.env` guarda sua senha: ele fica só na sua máquina e já está no
> `.gitignore` para nunca subir ao GitHub.

## Como usar

```bash
# 1º teste — só 2 CPFs, com o navegador visível para você acompanhar:
python consulta_margem.py --limite 2

# Rodar a planilha inteira:
python consulta_margem.py

# Rodar a partir de um CSV/Excel baixado (planilha privada):
python consulta_margem.py --arquivo "C:\Downloads\Pagina1.csv"

# Testar um CPF específico:
python consulta_margem.py --cpf 017.157.179-70

# Rodar sem mostrar o navegador (depois que validar que está tudo certo):
python consulta_margem.py --headless
```

Outras opções: `--coluna "CPF Cliente"`, `--gid 0` (aba da planilha),
`--limite N`, `--timeout 120`, `--pausa 2`, `--saida caminho.csv`.

## Saída

- `resultados/resultados_margem_AAAAMMDD_HHMMSS.csv` — separado por `;`,
  abre direto no Excel (uma linha por **CPF + NB**);
- `resultados/resultados_margem_AAAAMMDD_HHMMSS.xlsx` — mesma coisa em Excel;
- Coluna **Status**: `OK`, `SEM RESULTADO ...` ou `ERRO ...`.

## Problemas comuns

| Sintoma | O que fazer |
|---|---|
| "Falha no login" | Confira usuário/senha no `.env` (sem espaços sobrando). |
| "Não consegui baixar a planilha" | Compartilhe a planilha como *Qualquer pessoa com o link – Leitor*, ou baixe o CSV e use `--arquivo`. |
| Muitos `SEM RESULTADO` | Abra os prints em `debug/` — pode ser lentidão do sistema (aumente `--timeout`) ou CPF sem cadastro. |
| O site mudou de layout | Os prints em `debug/` mostram a tela; os textos-alvo (rótulos, botões) estão no topo do `consulta_margem.py` para ajustar. |

## Aviso — LGPD

Os arquivos gerados contêm dados pessoais de clientes (CPF, nome, benefício).
Trate-os com o mesmo cuidado da planilha original: não compartilhe fora da
operação e apague o que não for mais necessário. As pastas `resultados/` e
`debug/` já estão no `.gitignore` para não subirem ao GitHub por engano.
