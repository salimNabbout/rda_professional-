from datetime import datetime, timedelta
from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from sqlalchemy import func
from app.permissions import roles_required
from app.extensions import db
from app.models import User, RDAAuditLog, RDARecord

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


@admin_bp.route("/acessos")
@login_required
@roles_required("admin")
def acessos():
    """Controle de acesso: pilha de salvamentos do RDA por colaborador."""
    f_user_id = request.args.get("f_user", "").strip()
    f_data_ini = request.args.get("f_ini", "").strip()
    f_data_fim = request.args.get("f_fim", "").strip()

    query = RDAAuditLog.query
    if f_user_id:
        try:
            query = query.filter(RDAAuditLog.user_id == int(f_user_id))
        except ValueError:
            pass
    # Janela de filtro vem em data local (BRT); ts é gravado em UTC.
    # Convertemos BRT -> UTC somando 3h.
    BRT_TO_UTC = timedelta(hours=3)
    if f_data_ini:
        try:
            dt = datetime.strptime(f_data_ini, "%Y-%m-%d") + BRT_TO_UTC
            query = query.filter(RDAAuditLog.ts >= dt)
        except ValueError:
            pass
    if f_data_fim:
        try:
            dt = datetime.strptime(f_data_fim, "%Y-%m-%d") + timedelta(days=1) + BRT_TO_UTC
            query = query.filter(RDAAuditLog.ts < dt)
        except ValueError:
            pass

    logs = query.order_by(RDAAuditLog.ts.desc()).limit(500).all()

    # Agrupa por usuário + dia (em BRT) para visão calendário
    BRT_OFFSET = timedelta(hours=-3)
    calendario = {}
    for log in logs:
        dia = (log.ts + BRT_OFFSET).strftime("%Y-%m-%d")
        chave = (log.user_id, log.user_display, dia)
        calendario[chave] = calendario.get(chave, 0) + 1

    calendario_lista = [
        {"user_id": k[0], "user_display": k[1], "dia": k[2], "qtd": v}
        for k, v in calendario.items()
    ]
    calendario_lista.sort(key=lambda x: (x["dia"], x["user_display"]), reverse=True)

    usuarios = User.query.order_by(User.nome_completo.asc()).all()

    return render_template(
        "admin/acessos.html",
        logs=logs,
        calendario=calendario_lista,
        usuarios=usuarios,
        filtros={"f_user": f_user_id, "f_ini": f_data_ini, "f_fim": f_data_fim},
    )


@admin_bp.route("/users")
@login_required
@roles_required("admin")
def users():
    users = User.query.order_by(User.username.asc()).all()
    return render_template("admin/users.html", users=users)


@admin_bp.route("/users/new", methods=["POST"])
@login_required
@roles_required("admin")
def create_user():
    username = request.form.get("username", "").strip()
    nome_completo = request.form.get("nome_completo", "").strip()
    password = request.form.get("password", "")
    role = request.form.get("role", "colaborador").strip()
    acesso_tap = request.form.get("acesso_tap") == "1"
    acesso_produtividade = request.form.get("acesso_produtividade") == "1"

    if not username or not nome_completo or not password:
        flash("Preencha usuário, nome completo e senha.", "error")
        return redirect(url_for("admin.users"))
    if role not in ["colaborador", "gestor", "admin"]:
        flash("Perfil inválido.", "error")
        return redirect(url_for("admin.users"))
    if len(password) < 4:
        flash("A senha deve ter pelo menos 4 caracteres.", "error")
        return redirect(url_for("admin.users"))

    exists = User.query.filter(func.lower(User.username) == username.lower()).first()
    if exists:
        flash(f"Usuário '{username}' já existe.", "error")
        return redirect(url_for("admin.users"))

    user = User(
        username=username,
        nome_completo=nome_completo,
        role=role,
        acesso_tap=acesso_tap,
        acesso_produtividade=acesso_produtividade,
        must_change_password=True,
    )
    user.set_password(password)
    db.session.add(user)
    db.session.commit()
    flash(f"Usuário '{username}' criado. Ele deverá trocar a senha no primeiro acesso.", "success")
    return redirect(url_for("admin.users"))


@admin_bp.route("/users/bulk-update", methods=["POST"])
@login_required
@roles_required("admin")
def bulk_update_users():
    """Atualiza em lote todos os usuários com alterações no formulário."""
    alterados = 0
    erros = []
    for user in User.query.all():
        prefixo = f"user_{user.id}_"
        nome = request.form.get(f"{prefixo}nome_completo", "").strip()
        role = request.form.get(f"{prefixo}role", user.role).strip()
        acesso_tap = request.form.get(f"{prefixo}acesso_tap") == "1"
        acesso_produtividade = request.form.get(f"{prefixo}acesso_produtividade") == "1"

        if f"{prefixo}nome_completo" not in request.form:
            continue
        if not nome:
            erros.append(f"{user.username}: nome em branco")
            continue
        if role not in ["colaborador", "gestor", "admin"]:
            erros.append(f"{user.username}: perfil inválido")
            continue

        mudou = (
            user.nome_completo != nome
            or user.role != role
            or user.acesso_tap != acesso_tap
            or user.acesso_produtividade != acesso_produtividade
        )
        if mudou:
            user.nome_completo = nome
            user.role = role
            user.acesso_tap = acesso_tap
            user.acesso_produtividade = acesso_produtividade
            alterados += 1

    if alterados:
        db.session.commit()

    if erros:
        flash("Não salvo: " + "; ".join(erros), "error")
    elif alterados:
        flash(f"{alterados} usuário(s) atualizado(s) com sucesso.", "success")
    else:
        flash("Nenhuma alteração detectada.", "success")
    return redirect(url_for("admin.users"))


@admin_bp.route("/users/<int:user_id>/delete", methods=["POST"])
@login_required
@roles_required("admin")
def delete_user(user_id: int):
    if user_id == current_user.id:
        flash("Você não pode excluir sua própria conta.", "error")
        return redirect(url_for("admin.users"))

    user = User.query.get_or_404(user_id)
    nome = user.display_name
    qtd_registros = len(user.records)
    db.session.delete(user)
    db.session.commit()
    if qtd_registros:
        flash(
            f"Usuário '{nome}' excluído. {qtd_registros} registro(s) RDA associado(s) também foram removidos.",
            "success",
        )
    else:
        flash(f"Usuário '{nome}' excluído.", "success")
    return redirect(url_for("admin.users"))


@admin_bp.route("/limpar-dados", methods=["POST"])
@login_required
@roles_required("admin")
def limpar_dados():
    """Limpeza seletiva de dados — apenas admin."""
    confirmacao = request.form.get("confirmacao", "").strip()
    if confirmacao != "CONFIRMAR":
        flash("Confirmação inválida. Digite CONFIRMAR exatamente.", "error")
        return redirect(url_for("admin.users"))

    apagar_rda     = request.form.get("apagar_rda")     == "1"
    apagar_logs    = request.form.get("apagar_logs")    == "1"

    total_rda  = 0
    total_logs = 0

    if apagar_rda:
        total_rda = RDARecord.query.delete(synchronize_session=False)

    if apagar_logs:
        total_logs = RDAAuditLog.query.delete(synchronize_session=False)

    db.session.commit()

    partes = []
    if apagar_rda:   partes.append(f"{total_rda} registro(s) RDA")
    if apagar_logs:  partes.append(f"{total_logs} log(s) de auditoria")

    if partes:
        flash(f"Limpeza concluída: {', '.join(partes)} removido(s).", "success")
    else:
        flash("Nenhuma opção selecionada — nada foi apagado.", "error")

    return redirect(url_for("admin.users"))
