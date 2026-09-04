from datetime import date
from decimal import Decimal

from sqlalchemy import delete, or_, select
from sqlalchemy.orm import joinedload

from cpp_app.database import SessionLocal
from cpp_app.models import Attendance, Procedure, User
from cpp_app.security import hash_password, verify_password


PRICE_TYPES = {
    "PARTICULAR": "reference_value",
    "PARCEIROS": "partner_value",
    "PRO PREMIUM": "premium_value",
}


def get_price_by_type(procedure: Procedure, price_type: str) -> Decimal:
    attribute = PRICE_TYPES.get(price_type, "reference_value")
    value = getattr(procedure, attribute, None)
    if value is None:
        value = procedure.reference_value
    return Decimal(str(value))


def authenticate(username: str, password: str) -> User | None:
    with SessionLocal() as session:
        user = session.scalar(select(User).where(User.username == username.strip()))
        if user and verify_password(password, user.password_hash):
            session.expunge(user)
            return user
    return None


def create_user(username: str, password: str, role: str) -> User:
    username = username.strip()

    if not username:
        raise ValueError("Informe o nome do usuario.")
    if not password:
        raise ValueError("Informe a senha do usuario.")
    if role not in ("user", "admin"):
        raise ValueError("Classe de usuario invalida.")

    with SessionLocal() as session:
        existing = session.scalar(select(User).where(User.username == username))
        if existing:
            raise ValueError(f"Ja existe um usuario com o nome '{username}'.")

        user = User(
            username=username,
            password_hash=hash_password(password),
            role=role,
        )
        session.add(user)
        session.commit()
        session.refresh(user)
        session.expunge(user)
        return user


def list_users() -> list[User]:
    with SessionLocal() as session:
        users = session.scalars(select(User).order_by(User.username)).all()
        for user in users:
            session.expunge(user)
        return list(users)


def delete_user(user_id: int) -> None:
    with SessionLocal() as session:
        user = session.get(User, user_id)
        if not user:
            raise ValueError("Usuario nao encontrado.")

        if user.role == "admin":
            other_admins = session.scalar(
                select(User).where(User.role == "admin", User.id != user_id)
            )
            if not other_admins:
                raise ValueError("Nao e possivel excluir o unico administrador do sistema.")

        session.delete(user)
        session.commit()


def get_procedure_by_code(code: str) -> Procedure | None:
    with SessionLocal() as session:
        procedure = session.scalar(
            select(Procedure).where(Procedure.code == code.strip()).order_by(Procedure.source_sheet, Procedure.name)
        )
        if procedure:
            session.expunge(procedure)
        return procedure


def create_attendance(
    patient_name: str,
    attendance_date: date,
    paid_value: Decimal,
    sps_code: str,
    procedure_id: int,
    selected_price_type: str,
    report_month: int,
    report_year: int,
) -> Attendance:
    with SessionLocal() as session:
        procedure = session.get(Procedure, procedure_id)
        if not procedure:
            raise ValueError("Procedimento nao encontrado.")
        charged_value = get_price_by_type(procedure, selected_price_type)

        attendance = Attendance(
            patient_name=patient_name.strip(),
            attendance_date=attendance_date,
            paid_value=paid_value,
            charged_value=charged_value,
            selected_price_type=selected_price_type,
            report_month=report_month,
            report_year=report_year,
            sps_code=sps_code.strip(),
            procedure_id=procedure.id,
        )
        session.add(attendance)
        session.commit()
        session.refresh(attendance)
        return attendance


def update_attendance(
    attendance_id: int,
    patient_name: str,
    attendance_date: date,
    paid_value: Decimal,
    sps_code: str,
    procedure_id: int,
    selected_price_type: str,
    report_month: int,
    report_year: int,
) -> Attendance:
    with SessionLocal() as session:
        attendance = session.get(Attendance, attendance_id)
        if not attendance:
            raise ValueError("Atendimento nao encontrado.")

        procedure = session.get(Procedure, procedure_id)
        if not procedure:
            raise ValueError("Procedimento nao encontrado.")

        attendance.patient_name = patient_name.strip()
        attendance.attendance_date = attendance_date
        attendance.paid_value = paid_value
        attendance.charged_value = get_price_by_type(procedure, selected_price_type)
        attendance.selected_price_type = selected_price_type
        attendance.report_month = report_month
        attendance.report_year = report_year
        attendance.sps_code = sps_code.strip()
        attendance.procedure_id = procedure.id

        session.commit()
        session.refresh(attendance)
        return attendance


def delete_attendance(attendance_id: int) -> None:
    with SessionLocal() as session:
        attendance = session.get(Attendance, attendance_id)
        if not attendance:
            raise ValueError("Atendimento nao encontrado.")
        session.delete(attendance)
        session.commit()


def search_attendances(
    patient: str = "",
    procedure: str = "",
    sps_code: str = "",
    start: date | None = None,
    end: date | None = None,
    report_month: int | None = None,
    report_year: int | None = None,
):
    with SessionLocal() as session:

        query = (
            select(Attendance)
            .options(joinedload(Attendance.procedure))
            .order_by(Attendance.id.desc())
        )

        if patient.strip():
            query = query.where(
                Attendance.patient_name.ilike(
                    f"%{patient.strip()}%"
                )
            )

        if procedure.strip():
            termo = procedure.strip()
            query = query.join(Procedure).where(
                or_(
                    Procedure.code.ilike(f"%{termo}%"),
                    Procedure.name.ilike(f"%{termo}%"),
                )
            )

        if sps_code.strip():
            query = query.where(
                Attendance.sps_code.ilike(
                    f"%{sps_code.strip()}%"
                )
            )

        if start:
            query = query.where(
                Attendance.attendance_date >= start
            )

        if end:
            query = query.where(
                Attendance.attendance_date <= end
            )

        # NOVO FILTRO POR RELATÓRIO

        if report_month is not None:
            query = query.where(
                Attendance.report_month == report_month
            )

        if report_year is not None:
            query = query.where(
                Attendance.report_year == report_year
            )

        rows = list(session.scalars(query))

        return rows


def clear_all_attendances() -> int:
    """Apaga TODOS os atendimentos do banco. Retorna quantos registros foram excluidos."""
    with SessionLocal() as session:
        result = session.execute(delete(Attendance))
        session.commit()
        return result.rowcount