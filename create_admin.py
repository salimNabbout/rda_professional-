from app import create_app
from app.extensions import db
from app.models import User

app = create_app()

with app.app_context():
    username = "admin"
    password = "123456"
    user = User.query.filter_by(username=username).first()
    if user:
        print(f"Usuário '{username}' já existe.")
    else:
        user = User(username=username, role="admin")
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        print(f"Admin criado com sucesso. Usuário: {username} | Senha: {password}")
        print("Troque essa senha imediatamente em produção.")
