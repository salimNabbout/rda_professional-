from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required
from app.permissions import roles_required
from app.extensions import db
from app.models import User

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


@admin_bp.route("/users")
@login_required
@roles_required("admin")
def users():
    users = User.query.order_by(User.username.asc()).all()
    return render_template("admin/users.html", users=users)


@admin_bp.route("/users/<int:user_id>/role", methods=["POST"])
@login_required
@roles_required("admin")
def update_user_role(user_id: int):
    role = request.form.get("role", "colaborador").strip()
    if role not in ["colaborador", "gestor", "admin"]:
        flash("Perfil inválido.", "error")
        return redirect(url_for("admin.users"))

    user = User.query.get_or_404(user_id)
    user.role = role
    db.session.commit()
    flash("Perfil do usuário atualizado com sucesso.", "success")
    return redirect(url_for("admin.users"))


@admin_bp.route("/users/<int:user_id>/nome", methods=["POST"])
@login_required
@roles_required("admin")
def update_user_nome(user_id: int):
    nome = request.form.get("nome_completo", "").strip()
    if not nome:
        flash("Nome completo não pode ficar em branco.", "error")
        return redirect(url_for("admin.users"))

    user = User.query.get_or_404(user_id)
    user.nome_completo = nome
    db.session.commit()
    flash("Nome do usuário atualizado.", "success")
    return redirect(url_for("admin.users"))


@admin_bp.route("/users/<int:user_id>/tap", methods=["POST"])
@login_required
@roles_required("admin")
def update_user_tap(user_id: int):
    user = User.query.get_or_404(user_id)
    user.acesso_tap = request.form.get("acesso_tap") == "1"
    db.session.commit()
    flash(
        f"Acesso ao TAP {'liberado' if user.acesso_tap else 'revogado'} para {user.display_name}.",
        "success",
    )
    return redirect(url_for("admin.users"))
