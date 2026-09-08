import os
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    load_dotenv = None


if getattr(sys, "frozen", False):
    # Rodando como .exe compilado (PyInstaller): usa a pasta onde o executavel esta,
    # nao o caminho interno do pacote (que fica escondido dentro de _internal).
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"
EXPORT_DIR = Path.home() / "Downloads"

DATA_DIR.mkdir(exist_ok=True)
EXPORT_DIR.mkdir(exist_ok=True)

if load_dotenv:
    load_dotenv(BASE_DIR / ".env")

DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError(
        "DATABASE_URL nao foi encontrada. Crie um arquivo .env na mesma pasta "
        "do CPP.exe (ou do projeto, em modo desenvolvimento) com a linha:\n"
        "DATABASE_URL=postgresql+psycopg://usuario:senha@host:5432/cpp\n"
        "O sistema usa exclusivamente PostgreSQL; nao ha suporte a SQLite."
    )
PRICE_TABLE_PATH = os.getenv(
    "PRICE_TABLE_PATH",
    str(Path.home() / "Downloads" / "TABELA PARTICULARES-(Atualização 24-03-2026).xlsx"),
)
REPORT_TEMPLATE_PATH = os.getenv(
    "REPORT_TEMPLATE_PATH",
    str(Path.home() / "Downloads" / "PLANILHA VALORES NAO COBRADO PA.xlsx"),
)