from pathlib import Path
from urllib.parse import unquote, urlparse
import os
import subprocess


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


def run_psql(psql_path: Path, parsed, sql: str) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["PGPASSWORD"] = unquote(parsed.password or "")
    args = [
        str(psql_path),
        "-h",
        parsed.hostname or "localhost",
        "-p",
        str(parsed.port or 5432),
        "-U",
        unquote(parsed.username or ""),
        "-d",
        (parsed.path or "/").lstrip("/"),
        "-v",
        "ON_ERROR_STOP=1",
        "-c",
        sql,
    ]
    return subprocess.run(args, env=env, text=True, capture_output=True, timeout=20)


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

if parsed.scheme not in {"postgresql", "postgresql+psycopg", "postgresql+psycopg2"}:
    print("Banco configurado nao parece PostgreSQL.")
    raise SystemExit(0)

psql_candidates = [
    Path(r"C:\Program Files\PostgreSQL\18\bin\psql.exe"),
    Path(r"C:\Program Files\PostgreSQL\17\bin\psql.exe"),
    Path(r"C:\Program Files\PostgreSQL\16\bin\psql.exe"),
    Path(r"C:\Program Files\PostgreSQL\15\bin\psql.exe"),
]
psql_path = next((path for path in psql_candidates if path.exists()), None)
if not psql_path:
    print("psql.exe nao encontrado nos caminhos padrao.")
    raise SystemExit(1)

print("psql:", psql_path)

checks = {
    "versao": "select version();",
    "tabelas": (
        "select table_name from information_schema.tables "
        "where table_schema='public' order by table_name;"
    ),
    "contagens": (
        "select 'users' tabela, count(*) total from users "
        "union all select 'procedures', count(*) from procedures "
        "union all select 'attendances', count(*) from attendances;"
    ),
    "colunas": (
        "select table_name, column_name, data_type "
        "from information_schema.columns "
        "where table_schema='public' "
        "and table_name in ('users','procedures','attendances') "
        "order by table_name, ordinal_position;"
    ),
    "qualidade_atendimentos": (
        "select "
        "count(*) as total, "
        "count(*) filter (where charged_value is null) as sem_valor_cobrado, "
        "count(*) filter (where report_month is null or report_year is null) as sem_mes_relatorio, "
        "count(*) filter (where sps_code is null or trim(sps_code) = '') as sem_save "
        "from attendances;"
    ),
    "tipos_preco": (
        "select selected_price_type, count(*) total "
        "from attendances group by selected_price_type order by selected_price_type;"
    ),
    "meses_relatorio": (
        "select report_year, report_month, count(*) total "
        "from attendances group by report_year, report_month order by report_year, report_month;"
    ),
}

for label, sql in checks.items():
    print(f"\n[{label}]")
    result = run_psql(psql_path, parsed, sql)
    if result.returncode != 0:
        print(result.stderr.strip() or result.stdout.strip())
        raise SystemExit(result.returncode)
    print(result.stdout.strip())
