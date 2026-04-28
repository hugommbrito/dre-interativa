# DRE Interativa — Export VBA

Este módulo VBA lê as abas de **Entradas** e **Saídas** diretamente no
`DRE_Export.xlsm`, filtra pelo **período configurado**, valida que não há
lançamentos em `INVESTIGAR`, e injeta os dados brutos (linha a linha) como JSON
dentro do `dre_interativa.html`. O HTML continua standalone — um arquivo só
para abrir no navegador ou enviar por e-mail.

> **Compatível com Excel para Mac e Windows.** A versão atual não usa
> `Scripting.Dictionary` nem `System.Collections.ArrayList` — apenas VBA nativo.

## Arquitetura

```
  ┌──────────────────────────────────────────┐
  │ DRE_Export.xlsm  ← você aqui             │
  │                                          │
  │  • aba Config      (período + abas)      │
  │  • aba Mapeamento  (Grupo → Rubrica DRE) │
  │  • aba Entradas    (lançamentos entrada) │
  │  • aba Saídas      (lançamentos saída)   │
  │  • macro ExportarDRE                     │
  └───────────────────┬──────────────────────┘
                      ▼
         ┌─────────────────────────┐
         │ dre_interativa.html     │  ← abre no browser
         │  (com dados reais)      │
         └─────────────────────────┘
```

Os dados ficam **no próprio arquivo `.xlsm`** — não há mais dependência de
arquivos externos de Entradas ou Saídas.

## Passo 1 — Criar o workbook

1. Abra o Excel e crie um novo arquivo em branco.
2. Salve-o como **`DRE_Export.xlsm`** (pasta de trabalho habilitada para macro)
   na mesma pasta do `dre_interativa.html`.
3. Crie as abas com os nomes exatos: **`Config`**, **`Mapeamento`**,
   **`Entradas`** e **`Saídas`**.

## Passo 2 — Importar o módulo VBA

1. No Excel, abra o Editor VBA (`Option+F11` no Mac, `Alt+F11` no Windows).
2. Clique com botão direito em `VBAProject (DRE_Export.xlsm)` → **Import File…**
3. Selecione o arquivo `vba/modDREExport.bas`.
4. Feche o Editor VBA.

> **Mac**: se aparecer um pop-up pedindo permissão de acesso a arquivos,
> aceite — o VBA precisa ler as abas e gravar o HTML.

## Passo 3 — Preencher a aba `Config`

Layout: **Coluna A = chave**, **Colunas B, C, D, … = valor(es)**.

A maioria das chaves tem um único valor (coluna B). As chaves do tipo
**lista** (`entradas_abas`, `saidas_abas`) aceitam vários valores: coloque o
primeiro em B, o segundo em C, o terceiro em D, e assim por diante.

Sem cabeçalho obrigatório — a leitura começa na linha 2. Se você preferir
manter a linha 1 para títulos (`chave | valor1 | valor2 | …`), use-a como
cabeçalho humano e deixe os dados a partir da linha 2.

| Linha | Coluna A           | Coluna B                                          |
|-------|--------------------|---------------------------------------------------|
| 2     | `html_template`    | `/Users/.../DRE Interativa/dre_interativa.html`     |
| 3     | `html_output`      | `/Users/.../DRE Interativa/dre_interativa.html`     |
| 4     | `periodo_mes_ini`  | `1`                                               |
| 5     | `periodo_ano_ini`  | `2025`                                            |
| 6     | `periodo_mes_fim`  | `12`  ← **opcional**                              |
| 7     | `periodo_ano_fim`  | `2025` ← **opcional**                             |
| 8     | `entradas_abas`    | `Entradas`  *(coluna C: `Ent 2024`, etc.)*        |
| 9     | `saidas_abas`      | `Saídas`    *(coluna C: `Sai 2024`, etc.)*        |
| 10    | `json_output`      | `/Users/.../DRE Interativa/dre_data.json` ← **opcional** |

### Detalhes

- **`periodo_mes_ini` / `periodo_ano_ini`** — início do período (obrigatórios).
- **`periodo_mes_fim` / `periodo_ano_fim`** — fim do período (**opcionais**).
  Se omitidos (ou deixados em branco), o VBA varre todos os dados e usa o
  mês/ano da data mais recente encontrada como fim do período. O nome do
  período exibido na mensagem de conclusão incluirá `(automático)`.
- **`entradas_abas` / `saidas_abas`** — nomes exatos das abas dentro do
  próprio `DRE_Export.xlsm` (case-insensitive). Se omitidos, o padrão é
  `"Entradas"` e `"Saídas"`. Use múltiplas colunas quando o histórico estiver
  separado por ano ou empresa (`Entradas 2024`, `Entradas 2025`, etc.).
- **`html_template` vs `html_output`** — apontar os dois para o mesmo arquivo
  sobrescreve o HTML a cada execução (fluxo padrão). Para preservar
  snapshots, mude o `html_output` para um nome datado (ex.:
  `dre_2025-H1.html`).
- **`json_output`** — **opcional**. Se preenchido, grava também um arquivo
  `dre_data.json` com o payload completo. Útil para o fluxo de deploy na
  Vercel: cole o conteúdo do JSON na variável de ambiente `DRE_DATA_JSON`
  no painel da Vercel (veja a seção [Deploy na Vercel](#deploy-na-vercel)).

> **Período multi-ano**: para exportar, por exemplo, de julho/2024 a
> junho/2025, use `periodo_mes_ini=7`, `periodo_ano_ini=2024`,
> `periodo_mes_fim=6`, `periodo_ano_fim=2025`.

> **Mudando o período**: altere as chaves de início e fim e rode a macro
> de novo. Para sempre pegar "tudo até hoje", deixe as chaves de fim em
> branco.

## Passo 4 — Preencher a aba `Mapeamento`

Diz ao VBA como cada **Grupo** (das Saídas) entra na estrutura da DRE.

Cabeçalho na linha 1: **Grupo | RubricaId | Rubrica | Ordem**

Cole o conteúdo abaixo (a partir da linha 2):

| Grupo                        | RubricaId | Rubrica                   | Ordem |
|------------------------------|-----------|---------------------------|-------|
| Impostos                     | ded       | Deduções                  | 1     |
| Custos Indiretos             | csp       | Custo dos Serviços        | 2     |
| Software para Produto        | csp       | Custo dos Serviços        | 2     |
| Outros Custos                | csp       | Custo dos Serviços        | 2     |
| Pessoal                      | pes       | Pessoal & Benefícios      | 3     |
| Marketing                    | com       | Comercial & Marketing     | 4     |
| Consultorias                 | com       | Comercial & Marketing     | 4     |
| Custo de Venda               | com       | Comercial & Marketing     | 4     |
| Viagem                       | com       | Comercial & Marketing     | 4     |
| Administrativo               | ga        | G&A                       | 5     |
| Estrutura Física             | ga        | G&A                       | 5     |
| Outras Despesas              | ga        | G&A                       | 5     |
| Outros Investimentos         | ga        | G&A                       | 5     |
| Equipamentos e Tecnologia    | ti        | Equipamentos & TI         | 6     |
| Despesa Financeira           | rf        | Resultado Financeiro      | 8     |
| Despesas Bancárias           | rf        | Resultado Financeiro      | 8     |
| Custo Bancário               | rf        | Resultado Financeiro      | 8     |

**Observações:**

- `RubricaId` precisa ser exatamente um destes: `ded`, `csp`, `pes`, `com`,
  `ga`, `ti`, `da`, `rf`. São as âncoras que o HTML usa para encaixar cada
  Grupo no lugar certo da DRE.
- A coluna `Rubrica` é só o rótulo humano — pode ser ajustado sem mexer no
  HTML.
- `Movimentação Financeira` **não** vai no mapeamento: o VBA exclui
  automaticamente (é transferência entre contas).
- Se surgir um Grupo novo nas Saídas **dentro do período configurado** que
  não esteja aqui, o VBA aborta a exportação e diz qual Grupo adicionar.

## Passo 5 — Estrutura esperada das abas de dados

### Aba `Entradas` (linha 1 = cabeçalho)

| Col | Campo                         | Uso no VBA                            |
|-----|-------------------------------|---------------------------------------|
| A   | Empresa                       | vai pro JSON                          |
| B   | Pacote                        | vai pro JSON                          |
| C   | NF                            | vai pro JSON                          |
| D   | Cliente                       | vai pro JSON                          |
| E   | Valor Faturado                | vai pro JSON (base de competência)    |
| F   | Valor Recebido                | vai pro JSON (usado na inadimplência) |
| G   | Objeto                        | vai pro JSON                          |
| H   | **Data Emissão**              | filtro de período + regime comp.      |
| I   | Data Vencimento               | vai pro JSON                          |
| J   | Status                        | vai pro JSON                          |
| K   | Coluna1 (auxiliar)            | **ignorada** — só reserva o índice    |
| L   | Recebimento (Data Recebimento)| vai pro JSON                          |
| M   | Vertical_1                    | vai pro JSON                          |
| N   | %_vert_1                      | vai pro JSON                          |
| O   | Vertical_2                    | vai pro JSON                          |
| P   | %_vert_2                      | vai pro JSON                          |
| Q   | Obs (auxiliar)                | **ignorada** — só reserva o índice    |
| R   | Cobrança (auxiliar)           | **ignorada** — só reserva o índice    |
| S   | Priv\|Gov (Segmento)          | vai pro JSON                          |

> **Sem UF**: o layout atual da base não tem coluna de estado. Se vocês
> adicionarem UF no futuro, é só incluir `E_UF As Long = 20` no topo de
> `modDREExport.bas` e acrescentar `"""uf"":""" & JsonEsc(...) & ""","` no
> bloco de `LerEntradasJSON`.

### Aba `Saídas` (linha 1 = cabeçalho)

| Col | Campo                    |
|-----|--------------------------|
| A   | **Data Vencimento**      |
| B   | Data Pagamento           |
| C   | Fornecedor / Favorecido  |
| D   | Doc Associado            |
| E   | Descrição                |
| F   | Centro de Custo          |
| G   | Grupo                    |
| H   | Tipo de Gasto            |
| I   | Valor                    |
| J   | Banco                    |
| K   | Observações              |

> **Múltiplas abas**: se o histórico estiver separado (ex.: `Entradas 2024`
> e `Entradas 2025`), declare ambos os nomes em `entradas_abas` na Config.
> Todas as abas listadas devem ter a mesma estrutura de colunas.

## Passo 6 — Rodar a exportação

No Excel, com o `DRE_Export.xlsm` aberto:

1. `Option+F8` (Mac) ou `Alt+F8` (Windows) → **Macros**
2. Selecione **`ExportarDRE`** → **Executar**
3. Aguarde — aparece uma mensagem com o período exportado, o caminho do HTML
   gerado e o tempo decorrido.

## O que o VBA valida antes de gerar

- ✅ Aba `Config` contém todas as chaves obrigatórias e o período é válido
  (`1 ≤ mes_ini ≤ mes_fim ≤ 12`).
- ✅ Aba `Mapeamento` tem pelo menos uma linha.
- ✅ Todas as abas listadas em `entradas_abas` e `saidas_abas` existem no
  próprio `DRE_Export.xlsm`.
- ✅ Nenhuma Saída **dentro do período** com `Grupo = INVESTIGAR` ou
  `Tipo de Gasto = INVESTIGAR` (caso contrário **aborta** e lista as
  primeiras 5 ocorrências, com aba e linha).
- ✅ Todo `Grupo` presente nas Saídas **dentro do período** existe na aba
  `Mapeamento` (senão lista os Grupos órfãos e quantos lançamentos cada
  um tem).

Lançamentos fora do período são ignorados silenciosamente.

## Troubleshooting

| Erro                                              | Causa provável                                                                                      |
|---------------------------------------------------|-----------------------------------------------------------------------------------------------------|
| `Chave 'X' não encontrada na aba Config`          | Está faltando uma das chaves obrigatórias. Confira o Passo 3.                                       |
| `período_mes_ini inválido …` / `Período inválido …` | Mês fora de 1–12, ou início posterior ao fim. Verifique as quatro chaves de período.              |
| `Aba 'X' não encontrada neste workbook`           | O nome em `entradas_abas` / `saidas_abas` não corresponde a nenhuma aba do `DRE_Export.xlsm`.       |
| `Arquivo não encontrado: …`                       | Caminho de `html_template` ou `html_output` está errado.                                            |
| `Marcadores DRE_DATA_START / END não encontrados` | O HTML foi editado manualmente e perdeu os marcadores. Use a versão atual do `dre_interativa.html`. |
| `Encontrados N lançamentos em INVESTIGAR …`       | Classifique esses lançamentos antes de exportar (são listados com aba e linha).                     |
| `Grupos sem mapeamento …`                         | Adicione o Grupo listado na aba `Mapeamento`.                                                       |
| HTML abre e mostra "Sem dados para exibir"        | A macro não rodou OU o HTML aberto é o template sem injeção. Verifique que `html_output` aponta para o arquivo que você está abrindo. |

## Como atualizar no futuro

1. Cole os novos lançamentos nas abas **Entradas** e **Saídas** do
   `DRE_Export.xlsm` (ou em novas abas por ano — declare-as em
   `entradas_abas` / `saidas_abas` na Config).
2. Se surgir aba nova, inclua o nome em `entradas_abas` / `saidas_abas`.
3. Ajuste `periodo_ano` / `periodo_mes_ini` / `periodo_mes_fim` para o
   recorte desejado.
4. Abra o `DRE_Export.xlsm` e rode a macro `ExportarDRE` (`Option+F8`).
5. Abra o `dre_interativa.html` no browser — os dados refletem o último
   snapshot.

Nenhum servidor, nenhum banco. Tudo local.

---

## Deploy na Vercel

O repositório GitHub pode ser **público** — o template HTML não contém dados financeiros. Os dados reais ficam armazenados no Redis (Vercel KV / Upstash) e são protegidos por senha.

### Estrutura do repositório

```
/
├── dre_interativa.html   ← template (DRE_DATA = null)
├── api/
│   ├── data.js           ← GET — serve dados após autenticação (senha de leitura)
│   ├── upload.js         ← POST — recebe e persiste novo dre_data.json (senha de upload)
│   └── analyze.js        ← POST — proxy para API Anthropic (análise IA)
└── vercel.json           ← roteamento
```

### Setup inicial (uma vez)

1. Faça push do repositório para o GitHub e conecte-o à Vercel.
2. No painel da Vercel → **Storage**, crie um banco **KV (Upstash Redis)** e vincule ao projeto. As variáveis `KV_REST_API_URL`, `KV_REST_API_TOKEN` e `KV_REST_API_READ_ONLY_TOKEN` são preenchidas automaticamente.
3. No painel da Vercel → **Settings → Environment Variables**, crie:

| Variável             | Descrição                                                    |
|----------------------|--------------------------------------------------------------|
| `DRE_PASSWORD_READ`  | Senha para acessar os dados reais no browser                 |
| `DRE_PASSWORD_UPLOAD`| Senha para enviar novos dados ao servidor (pode ser diferente)|
| `ANTHROPIC_API_KEY`  | Chave da API Anthropic para habilitar a análise IA           |

4. Faça o primeiro deploy. A URL estará disponível para acesso.

### Atualizar dados na Vercel

1. Rode a macro `ExportarDRE` no Excel (configure `json_output` na aba Config para gerar o `dre_data.json`).
2. No browser, acesse a URL da Vercel e entre com a **senha de leitura**.
3. Clique no botão **"Atualizar dados"** (disponível no banner da área logada).
4. Selecione o `dre_data.json` gerado e informe a **senha de upload**.
5. Confirme — os dados são enviados para o Redis e ficam disponíveis imediatamente, sem redeploy.

### Tela de acesso

Ao abrir a URL da Vercel, o usuário vê três opções:

| Opção | Descrição |
|-------|-----------|
| **Demonstração** | Dados fictícios, sem senha, disponível para qualquer visitante |
| **Carregar arquivo** | Seleciona o `dre_data.json` local — dados ficam apenas nesta sessão do browser |
| **Acessar dados reais** | Insere a `DRE_PASSWORD_READ` para buscar os dados do servidor |
