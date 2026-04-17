from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from .extensions import db


STATUS_CHOICES = ["Em Andamento", "Concluído", "Atrasado"]
ROLE_CHOICES = ["colaborador", "gestor", "admin"]


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False)
    nome_completo = db.Column(db.String(150), nullable=False, default="")
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="colaborador")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    records = db.relationship("RDARecord", backref="owner", lazy=True, cascade="all, delete-orphan")

    @property
    def display_name(self) -> str:
        return self.nome_completo.strip() if self.nome_completo and self.nome_completo.strip() else self.username

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    def is_admin(self) -> bool:
        return self.role == "admin"

    def is_gestor(self) -> bool:
        return self.role == "gestor"

    def is_colaborador(self) -> bool:
        return self.role == "colaborador"


class RDARecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)

    colaborador = db.Column(db.String(150), nullable=False)
    cliente = db.Column(db.String(150), nullable=False)
    data = db.Column(db.String(10), nullable=False)

    hora_inicio_manha = db.Column(db.String(5), nullable=False, default="00:00")
    hora_final_manha = db.Column(db.String(5), nullable=False, default="00:00")
    hora_inicio_tarde = db.Column(db.String(5), nullable=False, default="00:00")
    hora_final_tarde = db.Column(db.String(5), nullable=False, default="00:00")

    duracao = db.Column(db.String(5), nullable=False, default="00:00")
    realizado = db.Column(db.Text, nullable=False)

    status_rda = db.Column(db.String(30), default="Em Andamento", nullable=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
