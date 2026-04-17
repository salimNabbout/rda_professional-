"""
Migra dados do schema antigo (pre-melhorias) para o novo.

Uso:
    python migrate_data.py

O que faz:
1. Faz backup do DB atual em instance/rda_backup_AAAAMMDD_HHMMSS.db
2. Le os registros antigos
3. Apaga o DB, deixa o Flask recriar com novo schema
4. Reimporta os registros convertendo tipos (data string -> Date) e status antigos

Mapeamento de status:
    Em aberto     -> Concluído
    Em analise    -> Em Andamento
    Aprovado      -> Concluído
    Reprovado     -> Concluído

Campos removidos (aprovador, observacoes_aprovacao) sao descartados.
"""

import os
import sqlite3
import shutil
from datetime import date, datetime

DB_PATH = os.path.join(os.path.dirname(__file__), "instance", "rda.db")

STATUS_MAP = {
    "Em aberto": "Concluído",
    "Em análise": "Em Andamento",
    "Em analise": "Em Andamento",
    "Aprovado": "Concluído",
    "Reprovado": "Concluído",
    # Formato intermediario (Iniciado) mapeado
    "Iniciado": "Concluído",
    # Ja no formato novo - passa direto
    "Concluído": "Concluído",
    "Em Andamento": "Em Andamento",
    "Atrasado": "Atrasado",
}


def backup():
    if not os.path.exists(DB_PATH):
        print("Nao ha DB para migrar (instance/rda.db nao existe).")
        return False
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(os.path.dirname(DB_PATH), f"rda_backup_{ts}.db")
    shutil.copy2(DB_PATH, backup_path)
    print(f"Backup criado em: {backup_path}")
    return True


def read_old_data():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    users = []
    try:
        cur.execute("SELECT * FROM user")
        users = [dict(r) for r in cur.fetchall()]
    except sqlite3.OperationalError as e:
        print(f"Aviso: tabela user nao lida: {e}")

    records = []
    try:
        cur.execute("SELECT * FROM rda_record")
        records = [dict(r) for r in cur.fetchall()]
    except sqlite3.OperationalError as e:
        print(f"Aviso: tabela rda_record nao lida: {e}")

    conn.close()
    return users, records


def reset_db():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print("DB antigo removido. Sera recriado com novo schema.")


def import_data(users, records):
    # Importa com Flask para que o schema novo seja criado primeiro
    from app import create_app
    from app.extensions import db
    from app.models import User, RDARecord

    app = create_app()
    with app.app_context():
        # Users
        for u in users:
            if User.query.filter_by(username=u["username"]).first():
                continue
            user = User(
                id=u["id"],
                username=u["username"],
                password_hash=u["password_hash"],
                role=u.get("role", "colaborador"),
                is_active_flag=True,
                must_change_password=False,
                created_at=datetime.fromisoformat(u["created_at"]) if isinstance(u.get("created_at"), str) else (u.get("created_at") or datetime.utcnow()),
            )
            db.session.add(user)
        db.session.commit()

        # Records
        for r in records:
            try:
                data_val = r.get("data")
                if isinstance(data_val, str):
                    data_val = date.fromisoformat(data_val)
                status_old = r.get("status_rda", "")
                status_new = STATUS_MAP.get(status_old, "Iniciado")

                record = RDARecord(
                    id=r["id"],
                    user_id=r["user_id"],
                    colaborador=r["colaborador"],
                    cliente=r["cliente"],
                    data=data_val,
                    hora_inicio=r["hora_inicio"],
                    hora_final=r["hora_final"],
                    duracao=r["duracao"],
                    realizado=r["realizado"],
                    status_rda=status_new,
                    responsavel_rda=r.get("responsavel_rda"),
                    periodo_referencia=r.get("periodo_referencia"),
                    is_active=True,
                )
                db.session.add(record)
            except Exception as e:
                print(f"Erro ao migrar registro id={r.get('id')}: {e}")
        db.session.commit()

        print(f"Importados: {len(users)} usuarios, {len(records)} registros RDA.")


if __name__ == "__main__":
    print("== Migracao de dados RDA ==")
    if not backup():
        print("Sem backup, saindo.")
        exit(0)

    print("Lendo dados antigos...")
    users, records = read_old_data()
    print(f"  {len(users)} usuarios, {len(records)} registros encontrados.")

    print("Recriando DB com novo schema...")
    reset_db()

    print("Importando dados...")
    import_data(users, records)

    print("Migracao concluida!")
