from functools import wraps
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import login_required, current_user

from app.extensions import db
from app.models import TAP, TAPItem, TAP_ENTREGAVEIS, TAP_STATUS_CHOICES


tap_bp = Blueprint("tap", __name__, url_prefix="/tap")


def tap_access_required(view_func):
    """Somente usuários com acesso_tap=True (ou admin) podem acessar."""
    @wraps(view_func)
    @login_required
    def wrapper(*args, **kwargs):
        if not (current_user.acesso_tap or current_user.is_admin()):
            abort(403)
        return view_func(*args, **kwargs)
    return wrapper


def _parse_float(v, default=0.0):
    try:
        return float(str(v).replace(",", ".").strip()) if v not in (None, "") else default
    except (ValueError, TypeError):
        return default


def _inicializar_itens(tap: TAP):
    """Cria os 15 itens fixos vazios para uma TAP recém-criada."""
    for ordem, entregavel in enumerate(TAP_ENTREGAVEIS, start=1):
        tap.itens.append(TAPItem(
            ordem=ordem, entregavel=entregavel,
            qtd_recursos=0.0, tempo=0.0, percentual_correcao=0.0,
        ))


@tap_bp.route("/")
@tap_access_required
def list_taps():
    taps = TAP.query.order_by(TAP.created_at.desc()).all()
    return render_template("tap/list.html", taps=taps)


@tap_bp.route("/novo", methods=["GET", "POST"])
@tap_access_required
def novo():
    if request.method == "POST":
        return _salvar(None)
    # prepara uma TAP em memória só para renderizar o form com 15 linhas vazias
    tap_stub = TAP(
        ctrl_numero="", cliente="", status_proposta="Aguardando",
        hh_valor=300.0, created_by_id=current_user.id,
    )
    _inicializar_itens(tap_stub)
    return render_template("tap/form.html", tap=tap_stub, editing=False)


@tap_bp.route("/<int:tap_id>", methods=["GET", "POST"])
@tap_access_required
def editar(tap_id: int):
    tap = TAP.query.get_or_404(tap_id)
    if request.method == "POST":
        return _salvar(tap)
    return render_template("tap/form.html", tap=tap, editing=True)


def _salvar(tap: TAP | None):
    ctrl = request.form.get("ctrl_numero", "").strip()
    cliente = request.form.get("cliente", "").strip()
    status = request.form.get("status_proposta", "Aguardando").strip()
    hh_valor = _parse_float(request.form.get("hh_valor"), 300.0)

    if status not in TAP_STATUS_CHOICES:
        status = "Aguardando"

    if not ctrl or not cliente:
        flash("CTRL Nº e Cliente são obrigatórios.", "error")
        return redirect(url_for("tap.novo") if tap is None else url_for("tap.editar", tap_id=tap.id))

    conflito = TAP.query.filter_by(ctrl_numero=ctrl).first()
    if conflito and (tap is None or conflito.id != tap.id):
        flash(f"Já existe uma TAP com o CTRL Nº '{ctrl}'.", "error")
        return redirect(url_for("tap.novo") if tap is None else url_for("tap.editar", tap_id=tap.id))

    if tap is None:
        tap = TAP(ctrl_numero=ctrl, cliente=cliente, status_proposta=status,
                  hh_valor=hh_valor, created_by_id=current_user.id)
        _inicializar_itens(tap)
        db.session.add(tap)
    else:
        tap.ctrl_numero = ctrl
        tap.cliente = cliente
        tap.status_proposta = status
        tap.hh_valor = hh_valor

    db.session.flush()  # garante que os itens estejam acessíveis por ordem

    # Atualiza os 15 itens a partir dos campos do form.
    # % vem do UI em 0..100 (ex: 10 = 10%) e é armazenada como decimal (0.1).
    itens_por_ordem = {i.ordem: i for i in tap.itens}
    for ordem in range(1, len(TAP_ENTREGAVEIS) + 1):
        qtd = _parse_float(request.form.get(f"item_{ordem}_qtd"))
        tempo = _parse_float(request.form.get(f"item_{ordem}_tempo"))
        pct_ui = _parse_float(request.form.get(f"item_{ordem}_pct"))
        item = itens_por_ordem.get(ordem)
        if not item:
            continue
        item.qtd_recursos = qtd
        item.tempo = tempo
        item.percentual_correcao = pct_ui / 100.0

    db.session.commit()
    flash("TAP salva com sucesso.", "success")
    return redirect(url_for("tap.editar", tap_id=tap.id))


@tap_bp.route("/<int:tap_id>/excluir", methods=["POST"])
@tap_access_required
def excluir(tap_id: int):
    tap = TAP.query.get_or_404(tap_id)
    db.session.delete(tap)
    db.session.commit()
    flash("TAP excluída.", "success")
    return redirect(url_for("tap.list_taps"))
