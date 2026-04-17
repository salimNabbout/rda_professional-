import os
from app import create_app

app = create_app()

if __name__ == "__main__":
    env = os.getenv("FLASK_ENV", "development").lower()
    debug = env != "production"
    port = int(os.getenv("PORT", "5000"))

    # Evita loop infinito de restart quando o projeto esta dentro do OneDrive:
    # O sync mexe em arquivos da venv e o reloader do Werkzeug reinicia sem parar.
    # Mantemos o debugger (erros detalhados) mas desligamos o auto-reload.
    use_reloader = debug and os.getenv("FLASK_RELOAD", "0") == "1"

    app.run(
        host="0.0.0.0",
        port=port,
        debug=debug,
        use_reloader=use_reloader,
    )
