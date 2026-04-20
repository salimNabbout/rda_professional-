from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from .extensions import db


STATUS_CHOICES = ["Em Andamento", "Concluído", "Atrasado"]
ROLE_CHOICES = ["colaborador", "gestor", "admin"]

# Lista fixa e imutável de entregáveis (escopo padrão CETEM).
# Mesma lista usada no campo "O que foi realizado" do RDA.
TAP_ENTREGAVEIS = [
    "Visita Técnica ou Comercial",
    "Reunião Técnica ou Comercial",
    "Elaboração de Proposta",
    "Projetos",
    "Desenvolvimento Software PLC",
    "Desenvolvimento Software IHM",
    "Desenvolvimento Software SCADA",
    "Desenvolvimento Software Agêntica",
    "Design (Telas)",
    "Montagem de Painéis",
    "Instalações de Painéis",
    "Comissionamento",
    "Partida Assistida",
    "Treinamento",
    "Elaboração de DataBook",
]

TAP_STATUS_CHOICES = ["Aguardando", "Fechado", "Perdido", "Concluído"]


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False)
    nome_completo = db.Column(db.String(150), nullable=False, default="")
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="colaborador")
    acesso_tap = db.Column(db.Boolean, nullable=False, default=False)
    acesso_produtividade = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    records = db.relationship("RDARecord", backref="owner", lazy=True, cascade="all, delete-orphan")

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

    @property
    def display_name(self) -> str:
        return self.nome_completo.strip() if self.nome_completo and self.nome_completo.strip() else self.username


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
    foi_atrasado = db.Column(db.Boolean, default=False, nullable=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)


class TAP(db.Model):
    """Termo de Abertura de Projeto."""
    __tablename__ = "tap"

    id = db.Column(db.Integer, primary_key=True)
    ctrl_numero = db.Column(db.String(50), unique=True, nullable=False)
    cliente = db.Column(db.String(200), nullable=False)
    status_proposta = db.Column(db.String(30), nullable=False, default="Aguardando")
    hh_valor = db.Column(db.Float, nullable=False, default=300.0)

    created_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_by = db.relationship("User", foreign_keys=[created_by_id])

    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    itens = db.relationship(
        "TAPItem", backref="tap", lazy="joined",
        cascade="all, delete-orphan", order_by="TAPItem.ordem",
    )

    @property
    def rotulo_cliente(self) -> str:
        """Formato usado na caixa Cliente do RDA: 'Cliente / CTRL Nº'."""
        return f"{self.cliente} / {self.ctrl_numero}"

    @property
    def valor_total(self) -> float:
        return sum(i.valor_total for i in self.itens)

    @property
    def valor_total_corrigido(self) -> float:
        return sum(i.valor_total_corrigido for i in self.itens)

    @property
    def disponivel_no_rda(self) -> bool:
        """True quando o projeto deve aparecer na caixa Cliente do RDA."""
        return self.status_proposta == "Fechado"


class RDAAuditLog(db.Model):
    """Pilha de salvamentos do RDA — registra cada operação de gravação
    (criação ou edição) feita por um colaborador, com data e hora do servidor."""
    __tablename__ = "rda_audit_log"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    user_display = db.Column(db.String(150), nullable=False, default="")
    action = db.Column(db.String(20), nullable=False)  # 'criar' | 'editar' | 'excluir'
    record_id = db.Column(db.Integer, nullable=True)  # pode ser None em exclusão
    resumo = db.Column(db.String(250), nullable=False, default="")  # ex: cliente/atividade/data
    ts = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)

    user = db.relationship("User", foreign_keys=[user_id])


class AppSetting(db.Model):
    """Tabela key/value simples para configurações globais da app."""
    __tablename__ = "app_setting"

    id = db.Column(db.Integer, primary_key=True)
    chave = db.Column(db.String(80), unique=True, nullable=False)
    valor = db.Column(db.String(200), nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)


class TAPItem(db.Model):
    __tablename__ = "tap_item"

    id = db.Column(db.Integer, primary_key=True)
    tap_id = db.Column(db.Integer, db.ForeignKey("tap.id"), nullable=False)
    ordem = db.Column(db.Integer, nullable=False)
    entregavel = db.Column(db.String(150), nullable=False)
    qtd_recursos = db.Column(db.Float, nullable=False, default=0.0)
    tempo = db.Column(db.Float, nullable=False, default=0.0)
    percentual_correcao = db.Column(db.Float, nullable=False, default=0.0)
    inicio_atividade = db.Column(db.String(10), nullable=True)  # YYYY-MM-DD
    fim_atividade = db.Column(db.String(10), nullable=True)  # YYYY-MM-DD — usado para auto-atraso

    @property
    def valor_total(self) -> float:
        return (self.qtd_recursos or 0) * (self.tempo or 0) * (self.tap.hh_valor or 0)

    @property
    def valor_total_corrigido(self) -> float:
        pct = self.percentual_correcao or 0
        if pct <= 0:
            return self.valor_total
        return self.valor_total * (1 - pct)
