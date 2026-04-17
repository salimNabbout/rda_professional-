from flask import render_template, request, jsonify, current_app
from .extensions import db


def register_error_handlers(app):
    @app.errorhandler(401)
    def unauthorized(e):
        if request.path.startswith("/api/"):
            return jsonify(error="Nao autenticado"), 401
        return render_template("errors/401.html"), 401

    @app.errorhandler(403)
    def forbidden(e):
        if request.path.startswith("/api/"):
            return jsonify(error="Acesso negado"), 403
        return render_template("errors/403.html"), 403

    @app.errorhandler(404)
    def not_found(e):
        if request.path.startswith("/api/"):
            return jsonify(error="Nao encontrado"), 404
        return render_template("errors/404.html"), 404

    @app.errorhandler(429)
    def ratelimited(e):
        msg = "Muitas tentativas. Aguarde e tente novamente."
        if request.path.startswith("/api/"):
            return jsonify(error=msg), 429
        return render_template("errors/429.html", mensagem=msg), 429

    @app.errorhandler(500)
    def server_error(e):
        db.session.rollback()
        current_app.logger.exception("Erro 500")
        if request.path.startswith("/api/"):
            return jsonify(error="Erro interno"), 500
        return render_template("errors/500.html"), 500
