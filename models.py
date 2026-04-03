"""
SQLAlchemy models for the latex-api database.
"""
import uuid
from datetime import datetime

from sqlalchemy import (
    Column, String, Integer, BigInteger, Boolean,
    DateTime, Text, Enum, Float, create_engine
)
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import declarative_base, sessionmaker
import enum
import os

Base = declarative_base()


class ConversionStatus(str, enum.Enum):
    success = "success"
    failed = "failed"
    processing = "processing"


class DocumentTemplate(str, enum.Enum):
    auto = "auto"
    article = "article"
    report = "report"
    ieee = "ieee"
    beamer = "beamer"


class Transaction(Base):
    """
    Enregistre chaque conversion DOCX → LaTeX effectuée via l'API.
    """
    __tablename__ = "transactions"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    document_name = Column(String(255), nullable=False)
    document_size_bytes = Column(BigInteger, nullable=False)
    status = Column(
        Enum(ConversionStatus, name="conversion_status"),
        nullable=False,
        default=ConversionStatus.processing,
    )
    template_used = Column(
        Enum(DocumentTemplate, name="document_template"),
        nullable=False,
        default=DocumentTemplate.auto,
    )
    language = Column(String(10), nullable=True)          # "fr", "en", "fr+en"
    has_bibliography = Column(Boolean, nullable=False, default=False)
    image_count = Column(Integer, nullable=False, default=0)
    input_tokens = Column(Integer, nullable=True)
    output_tokens = Column(Integer, nullable=True)
    processing_duration_ms = Column(Integer, nullable=True)
    error_message = Column(Text, nullable=True)
    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)

    def __repr__(self) -> str:
        return (
            f"<Transaction id={self.id} doc={self.document_name!r} "
            f"status={self.status} created_at={self.created_at}>"
        )


# ---------------------------------------------------------------------------
# Database engine helpers
# ---------------------------------------------------------------------------

def get_engine(database_url: str | None = None):
    url = database_url or os.environ.get(
        "DATABASE_URL",
        "postgresql://latex_user:latex_pass@localhost:5432/latex_api",
    )
    return create_engine(url, echo=False)


def get_session_factory(engine):
    return sessionmaker(bind=engine, autocommit=False, autoflush=False)


def init_db(engine):
    """Crée toutes les tables si elles n'existent pas encore."""
    Base.metadata.create_all(engine)
