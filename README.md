# DRE Interativa

DRE interativa gerada a partir de planilhas Excel via macro VBA. Arquivo HTML standalone que roda direto no browser — sem dependência de servidor para uso local.

## Modos de uso

| Modo | Como acessar | Dados |
|------|-------------|-------|
| **Local (VBA)** | Rodar macro `ExportarDRE` → abrir HTML no Finder | Reais, embutidos no HTML |
| **Demonstração** | Tela inicial → "Ver demonstração" | Fictícios, gerados no browser |
| **Arquivo JSON** | Tela inicial → "Carregar arquivo" → `dre_data.json` | Reais, apenas nesta sessão |
| **Vercel** | URL do deploy → "Acessar dados reais" → senha de leitura | Reais, no servidor (Redis) |

## Estrutura do repositório

```
├── dre_interativa.html   # Template HTML (DRE_DATA = null — sem dados reais)
├── api/
│   ├── data.js           # GET — serve dados após autenticação (senha de leitura)
│   ├── upload.js         # POST — recebe e persiste novo dre_data.json (senha de upload)
│   └── analyze.js        # POST — proxy para API da Anthropic (análise IA)
├── vercel.json           # Roteamento Vercel
└── vba/
    ├── modDREExport.bas  # Módulo VBA — importar no DRE_Export.xlsm
    └── README_INSTALACAO.md  # Instruções completas de setup
```

> **Arquivos fora do repositório (não commitados):**
> `DRE_Export.xlsm` e `dre_data.json` contêm dados financeiros reais e estão no `.gitignore`.

## Setup rápido

### Uso local (VBA)

Veja as instruções detalhadas em [vba/README_INSTALACAO.md](vba/README_INSTALACAO.md).

### Deploy na Vercel

1. Conecte este repositório à Vercel.
2. No painel da Vercel → **Storage**, crie um banco **KV (Upstash Redis)** e vincule ao projeto. As variáveis `KV_REST_API_URL`, `KV_REST_API_TOKEN` e `KV_REST_API_READ_ONLY_TOKEN` são preenchidas automaticamente.
3. Configure as variáveis de ambiente restantes:

| Variável | Descrição |
|---|---|
| `DRE_PASSWORD_READ` | Senha para acessar os dados reais no browser |
| `DRE_PASSWORD_UPLOAD` | Senha para enviar novos dados ao servidor (pode ser diferente) |
| `ANTHROPIC_API_KEY` | Chave da API Anthropic para habilitar a análise IA |

4. Deploy automático ao fazer push.

### Atualizar dados na Vercel

1. Rode a macro `ExportarDRE` no Excel (configure `json_output` na aba Config para gerar o `dre_data.json`).
2. No browser, acesse a URL da Vercel e entre com a senha de leitura.
3. Clique no botão **"Atualizar dados"** (disponível no banner da área logada).
4. Selecione o `dre_data.json` gerado e informe a **senha de upload**.
5. Confirme — os dados são enviados para o Redis e ficam disponíveis imediatamente.
