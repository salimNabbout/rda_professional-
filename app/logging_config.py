import logging
import os
from logging.handlers import RotatingFileHandler


def setup_logging(app):
    """Configura logging rotativo em instance/logs/app.log."""
    logs_dir = os.path.join(app.instance_path, "logs")
    os.makedirs(logs_dir, exist_ok=True)

    log_path = os.path.join(logs_dir, "app.log")
    handler = RotatingFileHandler(log_path, maxBytes=1_000_000, backupCount=5, encoding="utf-8")

    level = logging.DEBUG if app.config.get("DEBUG") else logging.INFO
    handler.setLevel(level)

    formatter = logging.Formatter(
        "[%(asctime)s] %(levelname)s in %(module)s: %(message)s"
    )
    handler.setFormatter(formatter)

    app.logger.addHandler(handler)
    app.logger.setLevel(level)
    app.logger.info("Logging inicializado (nivel=%s)", logging.getLevelName(level))
