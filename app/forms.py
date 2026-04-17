import re
from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, DateField, TimeField, TextAreaField, SelectField, HiddenField, BooleanField
from wtforms.validators import DataRequired, Length, ValidationError, EqualTo, Optional

from .models import STATUS_CHOICES, ROLE_CHOICES


def strong_password(form, field):
    """Valida senha: min 8, contem letra e numero."""
    pw = field.data or ""
    if len(pw) < 8:
        raise ValidationError("Senha deve ter no minimo 8 caracteres.")
    if not re.search(r"[A-Za-z]", pw):
        raise ValidationError("Senha deve conter ao menos uma letra.")
    if not re.search(r"\d", pw):
        raise ValidationError("Senha deve conter ao menos um numero.")


class LoginForm(FlaskForm):
    username = StringField("Usuario", validators=[DataRequired(), Length(max=80)])
    password = PasswordField("Senha", validators=[DataRequired()])


class RegisterForm(FlaskForm):
    username = StringField("Usuario", validators=[DataRequired(), Length(min=3, max=80)])
    password = PasswordField("Senha", validators=[DataRequired(), strong_password])
    password2 = PasswordField("Confirmar senha", validators=[DataRequired(), EqualTo("password", "Senhas nao conferem.")])


class ChangePasswordForm(FlaskForm):
    current_password = PasswordField("Senha atual", validators=[DataRequired()])
    password = PasswordField("Nova senha", validators=[DataRequired(), strong_password])
    password2 = PasswordField("Confirmar nova senha", validators=[DataRequired(), EqualTo("password", "Senhas nao conferem.")])


class ForcedChangePasswordForm(FlaskForm):
    """Troca obrigatoria apos reset pelo admin - nao pede senha atual."""
    password = PasswordField("Nova senha", validators=[DataRequired(), strong_password])
    password2 = PasswordField("Confirmar nova senha", validators=[DataRequired(), EqualTo("password", "Senhas nao conferem.")])


class RDAForm(FlaskForm):
    id = HiddenField()
    colaborador = StringField("Nome do Colaborador", validators=[DataRequired(), Length(max=150)])
    cliente = StringField("Cliente", validators=[DataRequired(), Length(max=150)])
    data = DateField("Data", validators=[DataRequired()])
    hora_inicio = TimeField("Hora Inicio", validators=[DataRequired()])
    hora_final = TimeField("Hora Final", validators=[DataRequired()])
    realizado = TextAreaField("O que foi realizado", validators=[DataRequired(), Length(max=5000)])
    status_rda = SelectField("Status do RDA", choices=[(s, s) for s in STATUS_CHOICES], default="Iniciado")
    responsavel_rda = StringField("Responsavel pelo RDA", validators=[Optional(), Length(max=150)])
    periodo_referencia = StringField("Periodo de Referencia", validators=[Optional(), Length(max=100)])

    def validate_hora_final(self, field):
        if self.hora_inicio.data and field.data and field.data <= self.hora_inicio.data:
            raise ValidationError("Hora Final deve ser maior que Hora Inicio.")


class ChangeRoleForm(FlaskForm):
    role = SelectField("Perfil", choices=[(r, r) for r in ROLE_CHOICES])


class EmptyForm(FlaskForm):
    """Formulario vazio apenas com CSRF - para acoes POST simples."""
    pass
