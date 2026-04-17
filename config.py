import os
from dotenv import load_dotenv

load_dotenv()


def _normalize_db_url(url: str) -> str:
    """Render/Heroku expõem a URL com 'postgres://'. SQLAlchemy 2.x só aceita 'postgresql://'."""
    if url.startswith("postgres://"):
        return "postgresql://" + url[len("postgres://"):]
    return url


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "troque-esta-chave-por-uma-chave-segura")
    SQLALCHEMY_DATABASE_URI = _normalize_db_url(os.getenv("DATABASE_URL", "sqlite:///rda.db"))
    SQLALCHEMY_TRACK_MODIFICATIONS = False
