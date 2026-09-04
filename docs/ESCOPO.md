# Escopo de Tecnologia

## 1. Nome do Projeto

CPP - Controle de Passagem de Paciente

## 2. Objetivo

Desenvolver um sistema desktop para Windows, destinado ao uso administrativo do hospital, permitindo o registro de procedimentos realizados, consulta automatica de informacoes da tabela de precos hospitalar, armazenamento dos dados dos pacientes e geracao de relatorios em Excel e PDF.

O sistema sera disponibilizado atraves de um arquivo executavel `.exe`, para utilizacao na rede interna do hospital.

## 3. Escopo Funcional

### Autenticacao

- Login e senha.
- Controle de permissao por perfil.

### Consulta de Procedimentos

- Consulta de procedimentos por codigo.
- Exibicao automatica de codigo, nome e valor de referencia.

### Cadastro de Atendimento

- Nome do paciente.
- Data do atendimento.
- Valor pago.
- SAVE.
- Procedimento vinculado.

### Armazenamento

- Salvamento em banco de dados.
- Historico completo dos atendimentos.

### Consultas

- Pesquisa por paciente.
- Pesquisa por data.
- Pesquisa por procedimento.
- Visualizacao do historico cadastrado.

### Relatorios

- Exportacao para Excel.
- Exportacao para PDF.

## 4. Tecnologias Utilizadas

- Python
- PySide6
- SQLAlchemy
- PostgreSQL
- SQLite para desenvolvimento local
- Pandas
- OpenPyXL
- ReportLab
- PyInstaller
- Git e GitHub

## 5. Arquitetura da Solucao

### Camada de Interface

Telas de login, cadastro, consulta e emissao de relatorios.

### Camada de Negocio

Validacoes, consulta de procedimentos, regras de cadastro, persistencia e relatorios.

### Camada de Dados

Armazenamento de usuarios, procedimentos e atendimentos.

## 6. Requisitos Nao Funcionais

- Compatibilidade com Windows 10 e Windows 11.
- Consulta rapida de procedimentos.
- Interface simples para usuarios administrativos.
- Protecao basica por autenticacao.
- Distribuicao por arquivo executavel.

## 7. Fluxo Operacional

1. Usuario realiza login.
2. Usuario informa o codigo do procedimento.
3. Sistema consulta a tabela de precos.
4. Sistema preenche automaticamente as informacoes do procedimento.
5. Usuario informa os dados complementares.
6. Registro e salvo no banco de dados.
7. Usuario consulta registros cadastrados.
8. Usuario gera relatorios em Excel ou PDF.

## 8. Resultado Esperado

Centralizar o registro dos procedimentos realizados, reduzir erros de preenchimento, agilizar o processo administrativo e facilitar a emissao de relatorios para controle interno do hospital.
