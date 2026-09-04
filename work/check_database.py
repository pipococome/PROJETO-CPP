from pathlib import Path
from urllib.parse import urlparse

from sqlalchemy import create_engine, inspect, text


def read_env(path: Path) -> dict[str, str]:
    values = {}
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip()
    return values


env = read_env(Path(".env"))
database_url = env.get("DATABASE_URL")

print("ENV existe:", Path(".env").exists())

if not database_url:
    print("DATABASE_URL: nao configurada")
    raise SystemExit(0)

parsed = urlparse(database_url)
masked = (
    f"{parsed.scheme}://***:***@{parsed.hostname or ''}"
    f"{(':' + str(parsed.port)) if parsed.port else ''}{parsed.path or ''}"
)
print("DATABASE_URL:", masked)

engine = create_engine(database_url, future=True)
with engine.connect() as connection:
    dialect = connection.dialect.name
    print("Dialeto:", dialect)
    print("Servidor:", connection.execute(text("select version()")).scalar())
    inspector = inspect(connection)
    tables = inspector.get_table_names()
    print("Tabelas:", ", ".join(tables) if tables else "(nenhuma)")
    for table in ["users", "procedures", "attendances"]:
        if table in tables:
            count = connection.execute(text(f"select count(*) from {table}")).scalar()
            print(f"{table}: {count} registro(s)")
