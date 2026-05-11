# Gerador Automático de Arquivos Excel

Aplicação desenvolvida em Python + Streamlit para geração automática de múltiplos arquivos Excel filtrados.

---

# Funcionalidades

- Upload de arquivos:
  - .xlsx
  - .xlsm

- Escolha de:
  - aba
  - coluna de filtro

- Geração:
  - todos os arquivos
  - apenas itens selecionados

- Preserva:
  - formatação
  - cores
  - estilos
  - bordas
  - largura de colunas
  - mesclagens
  - filtros

- Remove:
  - fórmulas
  - abas extras
  - tabelas Excel (ListObject)

- Mantém apenas:
  - cabeçalhos
  - linhas filtradas

---

# Estrutura do Projeto

```text
excel-filter-app/
│
├── app.py
├── requirements.txt
├── README.md
│
├── src/
│   ├── excel_utils.py
│   ├── file_utils.py
│   └── ui.py
│
├── temp/
│
└── output/
```

---

# Instalação

## 1. Clonar projeto

```bash
git clone <repositorio>
```

---

## 2. Criar ambiente virtual

### Windows

```bash
python -m venv venv
```

Ativar:

```bash
venv\Scripts\activate
```

---

## 3. Instalar dependências

```bash
pip install -r requirements.txt
```

---

# Execução

```bash
streamlit run app.py
```

---

# Tecnologias

- Python
- Streamlit
- Pandas
- OpenPyXL

---

# Compatibilidade

Compatível com:

- Excel .xlsx
- Excel .xlsm

---

# Observações Técnicas

## Conversão de fórmulas

Os arquivos finais não mantêm fórmulas.

Todas as células são convertidas para valores finais.

---

## Tabelas do Excel

As estruturas de tabela (ListObjects) são removidas automaticamente antes do salvamento.

Isso evita:

- corrupção
- alertas de reparo
- referências quebradas

---

## Segurança

O projeto utiliza:

- arquivos temporários
- try/except
- limpeza de nomes inválidos
- remoção segura de abas
- manipulação via openpyxl