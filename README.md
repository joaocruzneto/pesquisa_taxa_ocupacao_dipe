# Automação de Pesquisa de Mancha e Ocupação (DIPE)

## 📋 Descrição
Este projeto automatiza o fluxo de tratamento de dados da **Divisão de Pesquisas (DIPE)** do Grande Recife Consórcio de Transporte. O script foi desenvolvido para substituir processos manuais repetitivos em planilhas Google, garantindo maior rapidez e precisão nos cálculos de frota e índices de ocupação.

## ✨ Principais Funcionalidades
- **Menu Personalizado:** Adiciona uma interface intuitiva diretamente no Google Sheets.
- **Tratamento de Frota:** Padronização automática dos prefixos das empresas (MOB/CNO).
- **Cálculos Inteligentes:** Conversão automática de siglas de campo (ex: `7CV`, `BC`, `LT`) em dados numéricos de ocupação.
- **Relatórios Automáticos:** Geração de tabelas de análise por faixa horária e local em segundos.

## 🛠️ Tecnologias
- **Linguagem:** Google Apps Script (JavaScript V8)
- **Plataforma:** Google Workspace (Sheets)

## 📖 Como Instalar (Uso Interno)
1. No Google Sheets, acesse `Extensões` > `Apps Script`.
2. Crie um novo arquivo e cole o conteúdo de `src/main.gs`.
3. Salve e atualize a planilha para visualizar o menu **"Pesquisa de Mancha"**.
