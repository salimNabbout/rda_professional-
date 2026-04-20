from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_user, logout_user, login_required, current_user
from sqlalchemy import func
from app.extensions import db
from app.models import User
from app.forms import ChangePasswordForm, ForcedChangePasswordForm

auth_bp = Blueprint("auth", __name__, url_prefix="/auth")


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("main.index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user = User.query.filter(func.lower(User.username) == username.lower()).first()
        if not user or not user.check_password(password):
            flash("Usuário ou senha inválidos.", "error")
            return render_template("auth/login.html", register_mode=False)

        login_user(user)

        if user.must_change_password:
            flash("Por segurança, defina uma nova senha antes de continuar.", "error")
            return redirect(url_for("auth.change_password"))

        flash("Login realizado com sucesso.", "success")
        return redirect(url_for("main.index"))

    return render_template("auth/login.html", register_mode=False)


@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("main.index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        nome_completo = request.form.get("nome_completo", "").strip()
        password = request.form.get("password", "")

        if not username or not password or not nome_completo:
            flash("Preencha usuário, nome completo e senha.", "error")
            return render_template("auth/login.html", register_mode=True)

        exists = User.query.filter(func.lower(User.username) == username.lower()).first()
        if exists:
            flash("Esse usuário já existe.", "error")
            return render_template("auth/login.html", register_mode=True)

        user = User(username=username, nome_completo=nome_completo, role="colaborador")
        user.set_password(password)
        db.session.add(user)
        db.session.commit()

        flash("Conta criada com sucesso. Agora é só entrar.", "success")
        return redirect(url_for("auth.login"))

    return render_template("auth/login.html", register_mode=True)


@auth_bp.route("/trocar-senha", methods=["GET", "POST"])
@login_required
def change_password():
    forced = current_user.must_change_password

    if forced:
        form = ForcedChangePasswordForm()
    else:
        form = ChangePasswordForm()

    if form.validate_on_submit():
        if not forced:
            if not current_user.check_password(form.current_password.data):
                flash("Senha atual incorreta.", "error")
                return render_template("auth/change_password.html", form=form, forced=False)

        current_user.set_password(form.password.data)
        current_user.must_change_password = False
        db.session.commit()
        flash("Senha alterada com sucesso.", "success")
        return redirect(url_for("main.index"))

    return render_template("auth/change_password.html", form=form, forced=forced)


@auth_bp.route("/perfil")
@login_required
def perfil():
    form = ChangePasswordForm()
    return render_template("auth/perfil.html", form=form)


@auth_bp.route("/logout")
def logout():
    logout_user()
    flash("Sessão encerrada.", "success")
    return redirect(url_for("auth.login"))
