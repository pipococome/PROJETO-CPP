from datetime import date, datetime

from sqlalchemy import Date, DateTime, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from cpp_app.database import Base


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    username: Mapped[str] = mapped_column(String(80), unique=True, nullable=False)
    password_hash: Mapped[str] = mapped_column(String(128), nullable=False)
    role: Mapped[str] = mapped_column(String(30), default="admin", nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class Procedure(Base):
    __tablename__ = "procedures"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)

    code: Mapped[str] = mapped_column(
        String(30),
        nullable=False,
        index=True,
    )

    # Aumentado para suportar descrições muito grandes
    name: Mapped[str] = mapped_column(
        Text,
        nullable=False,
    )

    reference_value: Mapped[float] = mapped_column(
        Numeric(12, 2),
        nullable=False,
    )

    partner_value: Mapped[float] = mapped_column(
        Numeric(12, 2),
        nullable=True,
    )

    premium_value: Mapped[float] = mapped_column(
        Numeric(12, 2),
        nullable=True,
    )

    # Aumentado para suportar nomes de abas extensos
    source_sheet: Mapped[str] = mapped_column(
        Text,
        default="",
        nullable=False,
    )

    # Aumentado para evitar erros futuros
    price_type: Mapped[str] = mapped_column(
        Text,
        default="PARTICULAR",
        nullable=False,
    )


class Attendance(Base):
    __tablename__ = "attendances"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)

    patient_name: Mapped[str] = mapped_column(
        String(255),
        nullable=False,
        index=True,
    )

    attendance_date: Mapped[date] = mapped_column(
        Date,
        nullable=False,
        index=True,
    )

    paid_value: Mapped[float] = mapped_column(
        Numeric(12, 2),
        nullable=False,
    )

    charged_value: Mapped[float] = mapped_column(
        Numeric(12, 2),
        nullable=True,
    )

    selected_price_type: Mapped[str] = mapped_column(
        String(30),
        default="PARTICULAR",
        nullable=False,
    )

    report_month: Mapped[int] = mapped_column(
        Integer,
        nullable=True,
    )

    report_year: Mapped[int] = mapped_column(
        Integer,
        nullable=True,
    )

    sps_code: Mapped[str] = mapped_column(
        String(60),
        nullable=False,
    )

    procedure_id: Mapped[int] = mapped_column(
        ForeignKey("procedures.id"),
        nullable=False,
    )

    created_at: Mapped[datetime] = mapped_column(
        DateTime,
        default=datetime.utcnow,
    )

    procedure: Mapped[Procedure] = relationship("Procedure")