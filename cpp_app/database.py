from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

from cpp_app.config import DATABASE_URL

engine = create_engine(DATABASE_URL, future=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
Base = declarative_base()


def create_tables() -> None:
    from cpp_app import models  # noqa: F401

    Base.metadata.create_all(bind=engine)