# CPP - Controle de Passagem de Paciente

Sistema desktop em Python para controle administrativo de passagens de pacientes, consulta de procedimentos hospitalares e geracao de relatorios em Excel e PDF.

## Tecnologias

- Python
- PySide6
- SQLAlchemy
- PostgreSQL ou SQLite para desenvolvimento
- Pandas e OpenPyXL
- ReportLab
- PyInstaller

## Como rodar

1. Crie um ambiente virtual:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Instale as dependencias:

```powershell
pip install -r requirements.txt
```

3. Coloque a tabela de precos no caminho padrao ou configure no `.env`:

```powershell
PRICE_TABLE_PATH=C:\Users\barbi\Downloads\TABELA PARTICULARES-(Atualização 24-03-2026).xlsx
REPORT_TEMPLATE_PATH=C:\Users\barbi\Downloads\PLANILHA VALORES NAO COBRADO PA.xlsx
```

4. Execute o sistema:

```powershell
python -m cpp_app
```

O primeiro acesso padrao e:

- Usuario: `admin`
- Senha: `admin123`

## Observacao

Por padrao, o projeto usa SQLite para facilitar testes locais. Para PostgreSQL, configure a variavel `DATABASE_URL` no arquivo `.env`.

Na primeira execucao, o sistema importa automaticamente os procedimentos com codigo encontrados na planilha configurada em `PRICE_TABLE_PATH`.

Os relatorios em Excel usam a planilha configurada em `REPORT_TEMPLATE_PATH` como modelo institucional.
