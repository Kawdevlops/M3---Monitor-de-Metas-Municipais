HEAD
# 📊 Meta 49 - Automação de Recapeamento

Aplicação desenvolvida em **Python + Streamlit** para automatizar o processamento e consolidação dos dados da **Meta 49 - Recapeamento**, integrando informações de arquivos **Excel (CONSEMAVI)** e **Word (CONVIAS)**.

---

## 🚀 Objetivo

Automatizar a leitura, tratamento e consolidação dos dados de recapeamento, eliminando processos manuais e garantindo:

- ✔️ Padronização dos dados  
- ✔️ Integração entre fontes diferentes  
- ✔️ Geração automática de relatório final  
- ✔️ Redução de erros humanos  

---

## 🧠 Como funciona

O sistema realiza:

1. Leitura do arquivo **Word (CONVIAS)**  
2. Leitura do arquivo **Excel (CONSEMAVI)**  
3. Tratamento e padronização dos dados  
4. Consolidação por:
   - Subprefeitura  
   - Mês  
5. Cálculo automático:
   - Convias  
   - Consemavi  
   - Total requalificado  
6. Preenchimento automático de um modelo Excel formatado  

---

## 🖥️ Interface

A aplicação foi construída com **Streamlit**, permitindo:

- Upload de arquivos (Word e Excel)
- Seleção de:
  - Ano  
  - Mês (ou todos)
- Geração automática do relatório
- Visualização do resultado
- Download do Excel final

---

## 🛠️ Tecnologias utilizadas

- Python  
- Pandas  
- OpenPyXL  
- Python-Docx  
- Streamlit  

---

---

## ▶️ Como executar localmente

```bash
# Criar ambiente virtual
python -m venv venv

# Ativar (Windows)
venv\Scripts\activate

# Instalar dependências
pip install -r requirements.txt

# Rodar aplicação
streamlit run app_ui.py

## 📂 Estrutura do projeto
=======
# M3 Monitor de Metas Municipais

Sistema desenvolvido para automatizar a coleta, tratamento e consolidação de dados de metas municipais, integrando informações de diferentes fontes (Excel e Word) e gerando relatórios padronizados para acompanhamento e tomada de decisão.


