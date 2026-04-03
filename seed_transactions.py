"""
Seeder de transactions pour latex-api.

Usage:
    python seed_transactions.py [--count N] [--url postgresql://...]

Options:
    --count N    Nombre de transactions à insérer (défaut: 100)
    --url URL    URL de connexion PostgreSQL
                 (défaut: env DATABASE_URL ou localhost:5432)
    --reset      Vide la table avant d'insérer
    --help       Affiche cette aide
"""
import argparse
import random
import sys
import uuid
from datetime import datetime, timedelta

from models import (
    ConversionStatus,
    DocumentTemplate,
    Transaction,
    get_engine,
    get_session_factory,
    init_db,
)

# ---------------------------------------------------------------------------
# Données réalistes pour la simulation
# ---------------------------------------------------------------------------

DOCUMENT_NAMES = [
    "these_doctorat_{}.docx",
    "rapport_stage_{}.docx",
    "article_recherche_{}.docx",
    "memoire_master_{}.docx",
    "cours_latex_{}.docx",
    "publication_ieee_{}.docx",
    "conference_paper_{}.docx",
    "revue_litterature_{}.docx",
    "rapport_projet_{}.docx",
    "manuel_utilisateur_{}.docx",
    "analyse_donnees_{}.docx",
    "etude_cas_{}.docx",
]

LANGUAGES = ["fr", "en", "fr+en"]
LANGUAGE_WEIGHTS = [0.45, 0.45, 0.10]

TEMPLATES = list(DocumentTemplate)
TEMPLATE_WEIGHTS = [0.50, 0.20, 0.10, 0.15, 0.05]  # auto, article, report, ieee, beamer

# Distribution des statuts : 90% succès, 7% échec, 3% en cours
STATUS_CHOICES = [
    ConversionStatus.success,
    ConversionStatus.failed,
    ConversionStatus.processing,
]
STATUS_WEIGHTS = [0.90, 0.07, 0.03]

ERROR_MESSAGES = [
    "Claude API timeout after 60s",
    "Empty document: no extractable content found",
    "Invalid .docx structure: corrupted ZIP",
    "Rate limit exceeded — retry after 30s",
    "LaTeX marker not found in Claude response",
    "Document too large: exceeds 200KB JSON limit",
    "Unsupported image format: .emf conversion failed",
]


def random_date(start: datetime, end: datetime) -> datetime:
    delta = end - start
    return start + timedelta(seconds=random.randint(0, int(delta.total_seconds())))


def build_transaction(index: int) -> Transaction:
    status = random.choices(STATUS_CHOICES, STATUS_WEIGHTS)[0]
    template = random.choices(TEMPLATES, TEMPLATE_WEIGHTS)[0]
    language = random.choices(LANGUAGES, LANGUAGE_WEIGHTS)[0]

    # Taille du document : 10 Ko à 5 Mo
    doc_size = random.randint(10_000, 5_000_000)

    # Durée de traitement selon la taille (plus grand = plus long)
    base_ms = int(doc_size / 1000)              # ~1ms par Ko
    processing_ms = random.randint(
        max(500, base_ms // 2),
        min(120_000, base_ms * 3),
    ) if status != ConversionStatus.processing else None

    # Tokens Claude (corrélés à la taille du document)
    if status == ConversionStatus.success and processing_ms is not None:
        input_tokens = random.randint(500, 8000)
        output_tokens = random.randint(800, 12000)
    elif status == ConversionStatus.failed:
        input_tokens = random.randint(200, 4000)
        output_tokens = random.randint(0, 1000)
    else:
        input_tokens = None
        output_tokens = None

    # Bibliographie & images (uniquement pour les succès)
    has_bibliography = random.random() < 0.65 if status == ConversionStatus.success else False
    image_count = (
        random.choices([0, 1, 2, 3, 4, 5, 6, 8, 10], weights=[30, 20, 15, 12, 8, 6, 4, 3, 2])[0]
        if status == ConversionStatus.success
        else 0
    )

    error_message = (
        random.choice(ERROR_MESSAGES) if status == ConversionStatus.failed else None
    )

    # Date de création : sur les 6 derniers mois
    now = datetime.utcnow()
    created_at = random_date(now - timedelta(days=180), now)

    name_template = random.choice(DOCUMENT_NAMES)
    document_name = name_template.format(str(index).zfill(4))

    return Transaction(
        id=uuid.uuid4(),
        document_name=document_name,
        document_size_bytes=doc_size,
        status=status,
        template_used=template,
        language=language,
        has_bibliography=has_bibliography,
        image_count=image_count,
        input_tokens=input_tokens,
        output_tokens=output_tokens,
        processing_duration_ms=processing_ms,
        error_message=error_message,
        created_at=created_at,
    )


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Insère des données de transactions simulées dans PostgreSQL.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--count", type=int, default=100, help="Nombre de transactions (défaut: 100)")
    parser.add_argument("--url", type=str, default=None, help="URL de connexion PostgreSQL")
    parser.add_argument("--reset", action="store_true", help="Vide la table avant d'insérer")
    return parser.parse_args()


def main():
    args = parse_args()

    print("=" * 60)
    print("  Seeder de transactions — latex-api")
    print("=" * 60)

    # Connexion
    try:
        engine = get_engine(args.url)
        Session = get_session_factory(engine)
        init_db(engine)
        print(f"[OK] Connecté à la base de données")
    except Exception as exc:
        print(f"[ERREUR] Connexion impossible : {exc}")
        sys.exit(1)

    with Session() as session:
        # Reset optionnel
        if args.reset:
            deleted = session.query(Transaction).delete()
            session.commit()
            print(f"[OK] Table vidée ({deleted} enregistrements supprimés)")

        # Génération & insertion
        print(f"[...] Génération de {args.count} transactions...")
        transactions = [build_transaction(i) for i in range(1, args.count + 1)]

        batch_size = 50
        inserted = 0
        for i in range(0, len(transactions), batch_size):
            batch = transactions[i : i + batch_size]
            session.add_all(batch)
            session.commit()
            inserted += len(batch)
            print(f"  {inserted}/{args.count} transactions insérées", end="\r")

        print()  # nouvelle ligne après le \r

    # Résumé
    with Session() as session:
        total = session.query(Transaction).count()
        by_status = {
            s.value: session.query(Transaction).filter_by(status=s).count()
            for s in ConversionStatus
        }

    print()
    print("Résumé")
    print("-" * 40)
    print(f"  Total en base        : {total}")
    for status, count in by_status.items():
        print(f"  {status:<20} : {count}")
    print("=" * 60)
    print("[DONE] Seeder terminé avec succès.")


if __name__ == "__main__":
    main()
