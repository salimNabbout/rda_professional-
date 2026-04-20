from datetime import timedelta

from flask import Flask
from config import Config
from .extensions import db, login_manager, migrate


BRT_OFFSET = timedelta(hours=-3)


def to_brt(dt):
    """Converte um datetime UTC (naive) para BRT (UTC-3)."""
    if dt is None:
        return None
    return dt + BRT_OFFSET


def create_app():
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_object(Config)

    db.init_app(app)
    login_manager.init_app(app)
    migrate.init_app(app, db)

    from .models import User

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    from .auth.routes import auth_bp
    from .main.routes import main_bp
    from .admin.routes import admin_bp
    from .tap.routes import tap_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(tap_bp)

    @app.template_filter("brt")
    def brt_filter(dt):
        """Converte datetime UTC para BRT para exibição."""
        return to_brt(dt)

    @app.template_filter("brl")
    def brl_filter(valor):
        """Formata número como moeda brasileira: 1234.5 -> 'R$ 1.234,50'."""
        try:
            v = float(valor or 0)
        except (TypeError, ValueError):
            v = 0.0
        s = f"{v:,.2f}"
        s = s.replace(",", "_").replace(".", ",").replace("_", ".")
        return f"R$ {s}"

    with app.app_context():
        db.create_all()
        _ensure_schema_upgrades()
        _normalize_concluido()

    return app


def _ensure_schema_upgrades():
    """Aplica alterações de esquema leves (SQLite) sem depender do Flask-Migrate."""
    from sqlalchemy import text, inspect

    inspector = inspect(db.engine)

    if "user" in inspector.get_table_names():
        cols_user = {c["name"] for c in inspector.get_columns("user")}
        if "acesso_produtividade" not in cols_user:
            with db.engine.begin() as conn:
                conn.execute(text("ALTER TABLE user ADD COLUMN acesso_produtividade BOOLEAN NOT NULL DEFAULT 0"))

    if "rda_record" not in inspector.get_table_names():
        return
    cols = {c["name"] for c in inspector.get_columns("rda_record")}
    if "foi_atrasado" not in cols:
        with db.engine.begin() as conn:
            conn.execute(text("ALTER TABLE rda_record ADD COLUMN foi_atrasado BOOLEAN NOT NULL DEFAULT 0"))
            conn.execute(text("UPDATE rda_record SET foi_atrasado = 1 WHERE status_rda = 'Atrasado'"))

    if "tap_item" in inspector.get_table_names():
        cols_item = {c["name"] for c in inspector.get_columns("tap_item")}
        if "data_prevista" not in cols_item:
            with db.engine.begin() as conn:
                conn.execute(text("ALTER TABLE tap_item ADD COLUMN data_prevista VARCHAR(10)"))
                cols_item.add("data_prevista")
        if "inicio_atividade" not in cols_item:
            with db.engine.begin() as conn:
                conn.execute(text("ALTER TABLE tap_item ADD COLUMN inicio_atividade VARCHAR(10)"))
        if "fim_atividade" not in cols_item:
            with db.engine.begin() as conn:
                conn.execute(text("ALTER TABLE tap_item ADD COLUMN fim_atividade VARCHAR(10)"))
                # Migra valores antigos de data_prevista → fim_atividade
                conn.execute(text(
                    "UPDATE tap_item SET fim_atividade = data_prevista "
                    "WHERE data_prevista IS NOT NULL AND fim_atividade IS NULL"
                ))


def _normalize_concluido():
    """Propaga 'Concluído' para todo registro do mesmo (cliente + atividade).
    Se qualquer registro de um grupo está Concluído, todos os outros 'Em Andamento'
    passam a Concluído. PDI é ignorado. Idempotente."""
    from .models import RDARecord
    grupos = (
        db.session.query(RDARecord.cliente, RDARecord.realizado)
        .filter(RDARecord.status_rda == "Concluído", RDARecord.realizado != "PDI")
        .distinct()
        .all()
    )
    total = 0
    for cliente, realizado in grupos:
        if not cliente or not realizado:
            continue
        total += (
            RDARecord.query.filter(
                RDARecord.cliente == cliente,
                RDARecord.realizado == realizado,
                RDARecord.status_rda == "Em Andamento",
            )
            .update({"status_rda": "Concluído"}, synchronize_session=False)
        )
    if total:
        db.session.commit()
