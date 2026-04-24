# DRE Interativa

DRE interativa gerada a partir de planilhas Excel via macro VBA. Arquivo HTML standalone que roda direto no browser — sem dependência de servidor para uso local.

## Modos de uso

| Modo | Como acessar | Dados |
|------|-------------|-------|
| **Local (VBA)** | Rodar macro `ExportarDRE` → abrir HTML no Finder | Reais, embutidos no HTML |
| **Demonstração** | Tela inicial → "Ver demonstração" | Fictícios, gerados no browser |
| **Arquivo JSON** | Tela inicial → "Carregar arquivo" → `dre_data.json` | Reais, apenas nesta sessão |
| **Vercel** | URL do deploy → "Acessar dados reais" → senha | Reais, no servidor |

## Estrutura do repositório

```
├── dre_interativa.html   # Template HTML (DRE_DATA = null — sem dados reais)
├── api/
│   └── data.js           # Serverless function: serve dados após autenticação
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

1. Conecte este repositório à Vercel
2. Configure as variáveis de ambiente:
   - `DRE_PASSWORD` — senha para acesso aos dados reais
   - `DRE_DATA_JSON` — conteúdo do `dre_data.json` gerado pelo VBA
3. Deploy automático ao fazer push

Para atualizar os dados: rode a macro, copie o `dre_data.json` gerado e cole em `DRE_DATA_JSON` no painel da Vercel.
