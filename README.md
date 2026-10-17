# Projeto Integrador – Dashboard Levantamento de BD (Streamlit)

## Estrutura da pasta
```
projeto_bd/
├── app.py
├── requirements.txt
├── README.md
└── dados/
    ├── levantamento_BD_CDVs_via_1_e_2.xlsx
    └── levantamento_BD_Sinais_via_1_e_2.xlsx
```

## Passo a passo (Windows – PowerShell)
```powershell
# 1) entrar na pasta do projeto
cd C:\Users\SEU_USUARIO\workspace\projeto_bd

# 2) criar o ambiente virtual
python -m venv venv

# 3) ativar o ambiente virtual
venv\Scripts\Activate.ps1
#   (se der erro de permissão:  Set-ExecutionPolicy -Scope Process RemoteSigned   e tente de novo)
#   (no CMD:  venv\Scripts\activate.bat)

# 4) instalar as bibliotecas
pip install -r requirements.txt

# 5) rodar o app
streamlit run app.py
```
O navegador abre em http://localhost:8501. Para fechar: `Ctrl + C`. Para sair do venv: `deactivate`.

## Linux / Mac
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
