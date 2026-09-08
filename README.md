# CPP - Controle de Passagem de Paciente

Sistema desktop em Python para controle administrativo de passagens de pacientes, consulta de procedimentos hospitalares e geração de relatórios em Excel e PDF.

## Tecnologias

- Python
- PySide6 (interface gráfica)
- SQLAlchemy
- PostgreSQL
- OpenPyXL (leitura da tabela de preços e geração de relatórios em Excel)
- ReportLab (geração de relatórios em PDF)
- python-dotenv (leitura do arquivo `.env`)
- PyInstaller (empacotamento do `.exe`, via `cpp.spec`)

## Funcionalidades

- **Atendimentos**: cadastro de pacientes e exames/procedimentos, com busca por código ou nome na tabela de preços e cálculo automático do valor conforme o tipo (Particular, Parceiros, Pro Premium).
- **Pesquisa**: consulta e edição dos atendimentos já registrados.
- **Relatórios**: exportação de atendimentos em Excel e PDF, por mês/ano.
- **Orçamento**: geração de orçamento em PDF para procedimentos avulsos.
- **Usuários** (somente admin): cadastro, edição e remoção de usuários do sistema.
- **Atualizar tabela de preços** (somente admin): botão na aba de Atendimentos que recarrega os procedimentos diretamente do arquivo configurado em `PRICE_TABLE_PATH`, sem precisar reiniciar o sistema.

## Como rodar

1. Crie um ambiente virtual:
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Instale as dependências:
```powershell
pip install -r requirements.txt
```

3. Crie um banco PostgreSQL vazio e configure o arquivo `.env` na raiz do projeto (crie um novo, não versione o seu):
```powershell
DATABASE_URL=postgresql+psycopg://usuario:senha@localhost:5432/cpp
PRICE_TABLE_PATH=C:\caminho\para\TABELA PARTICULARES-(Atualização 24-03-2026).xlsx
REPORT_TEMPLATE_PATH=C:\caminho\para\PLANILHA VALORES NAO COBRADO PA.xlsx
```
`DATABASE_URL` é obrigatória — sem ela o sistema não inicia. As tabelas são criadas automaticamente no banco na primeira execução.

4. Execute o sistema:
```powershell
python -m cpp_app
```

O primeiro acesso padrão é:
- Usuário: `admin`
- Senha: `admin123`

> Troque a senha do usuário `admin` assim que possível — ela vem fixa no primeiro acesso.

## Observações

- O projeto usa exclusivamente PostgreSQL; a variável `DATABASE_URL` deve estar sempre configurada no `.env` apontando para o banco.
- Na primeira execução, o sistema importa automaticamente os procedimentos com código encontrados na planilha configurada em `PRICE_TABLE_PATH`. Atualizações posteriores da planilha podem ser aplicadas a qualquer momento pelo botão **Atualizar tabela de preços** (aba Atendimentos, visível apenas para usuários admin).
- Os relatórios em Excel usam a planilha configurada em `REPORT_TEMPLATE_PATH` como modelo institucional.
- O arquivo `.env` contém credenciais e caminhos locais — não deve ser commitado nem compartilhado; mantenha-o fora do controle de versão (`.gitignore`).
- Para gerar o executável (`.exe`), use o PyInstaller com a configuração já pronta em `cpp.spec`:
```powershell
pyinstaller cpp.spec
```
