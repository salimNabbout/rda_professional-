from datetime import datetime
from functools import wraps
from io import StringIO, BytesIO
import csv

from flask import Blueprint, render_template, request, redirect, url_for, flash, abort, make_response, send_file
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


# ---------------------------------------------------------------------------
# Exportações
# ---------------------------------------------------------------------------

def _fmt_brl(valor: float) -> str:
    s = f"{valor:,.2f}"
    return "R$ " + s.replace(",", "_").replace(".", ",").replace("_", ".")


@tap_bp.route("/export/csv")
@tap_access_required
def export_csv_lista():
    taps = TAP.query.order_by(TAP.created_at.desc()).all()

    output = StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow([
        "CTRL Nº", "Cliente", "Status da Proposta", "HH (R$)",
        "Valor Total (R$)", "Valor Corrigido (R$)", "Disponível no RDA",
        "Criado em",
    ])
    for t in taps:
        writer.writerow([
            t.ctrl_numero, t.cliente, t.status_proposta,
            f"{t.hh_valor:.2f}".replace(".", ","),
            f"{t.valor_total:.2f}".replace(".", ","),
            f"{t.valor_total_corrigido:.2f}".replace(".", ","),
            "Sim" if t.disponivel_no_rda else "Não",
            t.created_at.strftime("%d/%m/%Y %H:%M"),
        ])

    response = make_response("\ufeff" + output.getvalue())
    response.headers["Content-Type"] = "text/csv; charset=utf-8"
    response.headers["Content-Disposition"] = "attachment; filename=taps_lista.csv"
    return response


@tap_bp.route("/export/pdf")
@tap_access_required
def export_pdf_lista():
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    taps = TAP.query.order_by(TAP.created_at.desc()).all()

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    titulo = Paragraph("<b>TAPs — Termos de Abertura de Projeto</b>", styles["Title"])
    gerado = Paragraph(
        f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} — Total: {len(taps)} TAP(s).",
        styles["Normal"],
    )

    head_style = ParagraphStyle("head", parent=styles["BodyText"], fontSize=8, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=8, leading=10)

    headers = ["CTRL Nº", "Cliente", "Status", "HH", "Valor Total", "Valor Corrigido", "No RDA?"]
    data = [[Paragraph(h, head_style) for h in headers]]
    for t in taps:
        data.append([
            Paragraph(t.ctrl_numero, cell_style),
            Paragraph(t.cliente, cell_style),
            Paragraph(t.status_proposta, cell_style),
            Paragraph(_fmt_brl(t.hh_valor), cell_style),
            Paragraph(_fmt_brl(t.valor_total), cell_style),
            Paragraph(_fmt_brl(t.valor_total_corrigido), cell_style),
            Paragraph("Sim" if t.disponivel_no_rda else "Não", cell_style),
        ])

    col_widths = [35*mm, 65*mm, 28*mm, 28*mm, 36*mm, 36*mm, 22*mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2d5a")),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
    ]))
    doc.build([titulo, Spacer(1, 4*mm), gerado, Spacer(1, 4*mm), table])
    buffer.seek(0)

    return send_file(
        buffer, as_attachment=True,
        download_name=f"taps_lista_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mimetype="application/pdf",
    )


@tap_bp.route("/<int:tap_id>/export/csv")
@tap_access_required
def export_csv_detalhe(tap_id: int):
    tap = TAP.query.get_or_404(tap_id)

    output = StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(["CTRL Nº", tap.ctrl_numero])
    writer.writerow(["Cliente", tap.cliente])
    writer.writerow(["Status da Proposta", tap.status_proposta])
    writer.writerow(["HH (R$)", f"{tap.hh_valor:.2f}".replace(".", ",")])
    writer.writerow([])
    writer.writerow(["ID", "Entregável", "Qtd. Rec.", "Tempo (h)", "Valor Total (R$)", "% Correção", "Valor Corrigido (R$)"])
    for it in tap.itens:
        writer.writerow([
            it.ordem, it.entregavel,
            f"{it.qtd_recursos:g}".replace(".", ","),
            f"{it.tempo:g}".replace(".", ","),
            f"{it.valor_total:.2f}".replace(".", ","),
            f"{(it.percentual_correcao or 0) * 100:g}".replace(".", ",") + "%",
            f"{it.valor_total_corrigido:.2f}".replace(".", ","),
        ])
    writer.writerow([])
    writer.writerow(["", "", "", "TOTAL SEM CORREÇÃO", f"{tap.valor_total:.2f}".replace(".", ","), "", ""])
    writer.writerow(["", "", "", "TOTAL COM CORREÇÃO", "", "", f"{tap.valor_total_corrigido:.2f}".replace(".", ",")])

    response = make_response("\ufeff" + output.getvalue())
    response.headers["Content-Type"] = "text/csv; charset=utf-8"
    response.headers["Content-Disposition"] = f"attachment; filename=tap_{tap.ctrl_numero}.csv"
    return response


@tap_bp.route("/<int:tap_id>/export/pdf")
@tap_access_required
def export_pdf_detalhe(tap_id: int):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    tap = TAP.query.get_or_404(tap_id)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=12 * mm, rightMargin=12 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
    )
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=8, leading=10)
    head_style = ParagraphStyle("head", parent=styles["BodyText"], fontSize=8, textColor=colors.white, alignment=1)

    titulo = Paragraph("<b>Termo de Abertura de Projeto (TAP)</b>", styles["Title"])
    cab = Paragraph(
        f"<b>CTRL Nº:</b> {tap.ctrl_numero} &nbsp;&nbsp;&nbsp; "
        f"<b>Cliente:</b> {tap.cliente}<br/>"
        f"<b>Status da Proposta:</b> {tap.status_proposta} &nbsp;&nbsp;&nbsp; "
        f"<b>HH:</b> {_fmt_brl(tap.hh_valor)}<br/>"
        f"<b>Gerado em:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        styles["Normal"],
    )

    headers = ["ID", "Entregável", "Qtd", "Tempo", "Valor Total", "% Corr.", "Valor Corrigido"]
    data = [[Paragraph(h, head_style) for h in headers]]
    for it in tap.itens:
        data.append([
            Paragraph(str(it.ordem), cell_style),
            Paragraph(it.entregavel, cell_style),
            Paragraph(f"{it.qtd_recursos:g}", cell_style),
            Paragraph(f"{it.tempo:g}", cell_style),
            Paragraph(_fmt_brl(it.valor_total), cell_style),
            Paragraph(f"{(it.percentual_correcao or 0) * 100:g}%", cell_style),
            Paragraph(_fmt_brl(it.valor_total_corrigido), cell_style),
        ])
    data.append([
        Paragraph("", cell_style), Paragraph("<b>Totais</b>", cell_style),
        Paragraph("", cell_style), Paragraph("", cell_style),
        Paragraph(f"<b>{_fmt_brl(tap.valor_total)}</b>", cell_style),
        Paragraph("", cell_style),
        Paragraph(f"<b>{_fmt_brl(tap.valor_total_corrigido)}</b>", cell_style),
    ])

    col_widths = [10*mm, 70*mm, 14*mm, 16*mm, 30*mm, 16*mm, 30*mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2d5a")),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.whitesmoke, colors.white]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#eef3f8")),
    ]))

    doc.build([titulo, Spacer(1, 4*mm), cab, Spacer(1, 6*mm), table])
    buffer.seek(0)

    return send_file(
        buffer, as_attachment=True,
        download_name=f"tap_{tap.ctrl_numero}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mimetype="application/pdf",
    )
