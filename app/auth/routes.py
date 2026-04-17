from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_user, logout_user, current_user
from app.extensions import db
from app.models import User

auth_bp = Blueprint("auth", __name__, url_prefix="/auth")


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("main.index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user = User.query.filter_by(username=username).first()
        if not user or not user.check_password(password):
            flash("Usuário ou senha inválidos.", "error")
            return render_template("auth/login.html", register_mode=False)

        login_user(user)
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

        exists = User.query.filter_by(username=username).first()
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


@auth_bp.route("/logout")
def logout():
    logout_user()
    flash("Sessão encerrada.", "success")
    return redirect(url_for("auth.login"))
