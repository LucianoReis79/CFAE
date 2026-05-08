# Gerador Automático de Excel Filtrado

Este aplicativo Streamlit permite realizar o upload de uma planilha mestre e gerar automaticamente múltiplos arquivos Excel individuais, filtrados por valores únicos de uma coluna selecionada, preservando toda a formatação original (cores, larguras de colunas, fontes e mesclagens).

## 🛠️ Tecnologias Utilizadas
- **Python** como linguagem base [31].
- **Streamlit** para a interface web interativa [1].
- **Pandas** para manipulação rápida de dados [8].
- **Openpyxl** para manipulação de estilos e estruturas de workbooks Excel [2].

## 🚀 Como Executar

### 1. Criar Ambiente Virtual (Recomendado)
No terminal, dentro da pasta do projeto:
```bash
# Windows
python -m venv venv
.\venv\Scripts\activate

# Linux/Mac
python3 -m venv venv
source venv/bin/activate