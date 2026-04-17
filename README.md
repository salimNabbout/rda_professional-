# RDA App Profissional

Sistema web para RDA - Relatório Diário de Atividade, com:
- login
- cadastro de usuários
- papéis de acesso (`colaborador`, `gestor`, `admin`)
- CRUD de registros
- filtros
- exportação CSV
- API JSON autenticada
- painel administrativo de usuários
- base pronta para Flask-Migrate

## Requisitos
- Python 3.10+
- pip

## Instalação

### 1. Criar e ativar ambiente virtual

#### Windows
```bash
python -m venv .venv
.venv\Scripts\activate
```

#### Linux / Mac
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Instalar dependências
```bash
pip install -r requirements.txt
```

### 3. Criar arquivo `.env`
Copie o exemplo:
```bash
cp .env.example .env
```

No Windows, crie manualmente ou use:
```bash
copy .env.example .env
```

### 4. Rodar o sistema
```bash
python app.py
```

Acesse:
```text
http://127.0.0.1:5000/auth/login
```

## Criar admin inicial
```bash
python create_admin.py
```

## Migrações com Flask-Migrate

### Linux / Mac
```bash
export FLASK_APP=app.py
flask db init
flask db migrate -m "initial migration"
flask db upgrade
```

### Windows PowerShell
```powershell
$env:FLASK_APP = "app.py"
flask db init
flask db migrate -m "initial migration"
flask db upgrade
```

## Regras de perfil
- `colaborador`: vê e edita apenas os próprios registros
- `gestor`: vê todos os registros
- `admin`: vê tudo e acessa a lista de usuários

## Rotas principais
- `/auth/login`
- `/auth/register`
- `/`
- `/api/records`
- `/export/csv`
- `/admin/users`

## Próximos upgrades recomendados
- troca de senha
- reset de senha
- auditoria de alterações
- PDF real no backend
- deploy em Render ou Railway
