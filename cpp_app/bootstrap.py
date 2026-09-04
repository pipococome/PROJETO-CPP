from sqlalchemy import inspect, select, text

from cpp_app.config import PRICE_TABLE_PATH
from cpp_app.database import SessionLocal, create_tables, engine
from cpp_app.models import Procedure, User
from cpp_app.price_table import extract_procedures
from cpp_app.security import hash_password


DEFAULT_PROCEDURES = [
    {
        "code": "10101012",
        "name": "Consulta medica ambulatorial",
        "reference_value": "120.00",
        "partner_value": "100.00",
        "premium_value": "80.00",
        "source_sheet": "Exemplo",
    },
    {
        "code": "20202024",
        "name": "Exame laboratorial simples",
        "reference_value": "38.50",
        "partner_value": "32.00",
        "premium_value": "25.00",
        "source_sheet": "Exemplo",
    },
]


def ensure_schema() -> None:
    columns = {column["name"] for column in inspect(engine).get_columns("procedures")}
    statements = []
    if "partner_value" not in columns:
        statements.append("ALTER TABLE procedures ADD COLUMN partner_value NUMERIC(12, 2)")
    if "premium_value" not in columns:
        statements.append("ALTER TABLE procedures ADD COLUMN premium_value NUMERIC(12, 2)")
    attendance_columns = {column["name"] for column in inspect(engine).get_columns("attendances")}
    if "charged_value" not in attendance_columns:
        statements.append("ALTER TABLE attendances ADD COLUMN charged_value NUMERIC(12, 2)")
    if "selected_price_type" not in attendance_columns:
        statements.append("ALTER TABLE attendances ADD COLUMN selected_price_type VARCHAR(30) DEFAULT 'PARTICULAR' NOT NULL")
    if "report_month" not in attendance_columns:
        statements.append("ALTER TABLE attendances ADD COLUMN report_month INTEGER")
    if "report_year" not in attendance_columns:
        statements.append("ALTER TABLE attendances ADD COLUMN report_year INTEGER")
    if statements:
        with engine.begin() as connection:
            for statement in statements:
                connection.execute(text(statement))


def refresh_procedure_prices() -> None:
    imported = extract_procedures(PRICE_TABLE_PATH)
    if not imported:
        return

    with SessionLocal() as session:
        for item in imported:
            existing = session.scalar(
                select(Procedure).where(
                    Procedure.code == item["code"],
                    Procedure.name == item["name"],
                    Procedure.source_sheet == item["source_sheet"],
                )
            )
            if existing:
                existing.reference_value = item["reference_value"]
                existing.partner_value = item.get("partner_value")
                existing.premium_value = item.get("premium_value")
                existing.price_type = item.get("price_type", "PARTICULAR")
            else:
                session.add(Procedure(**item))
        session.commit()


def bootstrap_database() -> None:
    create_tables()
    ensure_schema()
    with SessionLocal() as session:
        admin = session.scalar(select(User).where(User.username == "admin"))
        if not admin:
            session.add(User(username="admin", password_hash=hash_password("admin123"), role="admin"))

        existing_procedure = session.scalar(select(Procedure).limit(1))
        if not existing_procedure:
            imported = extract_procedures(PRICE_TABLE_PATH)
            for item in imported or DEFAULT_PROCEDURES:
                session.add(Procedure(**item))
        else:
            missing_new_values = session.scalar(select(Procedure).where(Procedure.partner_value.is_not(None)).limit(1))
            if not missing_new_values:
                session.commit()
                refresh_procedure_prices()
                return

        session.commit()