import os
from dotenv import load_dotenv

load_dotenv()


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "troque-esta-chave-por-uma-chave-segura")
    SQLALCHEMY_DATABASE_URI = os.getenv("DATABASE_URL", "sqlite:///rda.db")
    SQLALCHEMY_TRACK_MODIFICATIONS = False
