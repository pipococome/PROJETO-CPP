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

DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{DATA_DIR / 'cpp.db'}")
PRICE_TABLE_PATH = os.getenv(
    "PRICE_TABLE_PATH",
    str(Path.home() / "Downloads" / "TABELA PARTICULARES-(Atualização 24-03-2026).xlsx"),
)
REPORT_TEMPLATE_PATH = os.getenv(
    "REPORT_TEMPLATE_PATH",
    str(Path.home() / "Downloads" / "PLANILHA VALORES NAO COBRADO PA.xlsx"),
)

