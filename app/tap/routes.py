from datetime import datetime
from functools import wraps
from io import StringIO, BytesIO
import csv

from flask import Blueprint, render_template, request, redirect, url_for, flash, abort, make_response, send_file
from flask_login import login_required, current_user

from app.extensions import db
from app.models import TAP, TAPItem, TAP_ENTREGAVEIS, TAP_STATUS_CHOICES, RDARecord, AppSetting


def _ano_suffix_atual() -> str:
    return datetime.utcnow().strftime("%y")


def _get_setting(chave: str) -> str:
    s = AppSetting.query.filter_by(chave=chave).first()
    return s.valor if s else ""


def _set_setting(chave: str, valor: str) -> None:
    s = AppSetting.query.filter_by(chave=chave).first()
    if s:
        s.valor = valor
    else:
        db.session.add(AppSetting(chave=chave, valor=valor))
    db.session.commit()


def _parse_ctrl(ctrl: str):
    """Separa 'NNNN/AA' em (numero_int, ano_str). Retorna (None, None) se inválido."""
    if not ctrl or "/" not in ctrl:
        return None, None
    partes = ctrl.strip().split("/")
    if len(partes) != 2:
        return None, None
    try:
        return int(partes[0]), partes[1].strip()
    except ValueError:
        return None, None


def proximo_ctrl() -> str:
    """Calcula o próximo CTRL no formato NNNN/AA.
    Regra:
    - Ano = ano corrente (2 dígitos).
    - Número = MAX(números dos TAPs existentes com o mesmo ano) + 1.
    - Se não houver TAP no ano corrente, usa a base configurada para o ano
      (setting 'ctrl_base_{AA}'); se também não houver, cai para 1.
    """
    ano = _ano_suffix_atual()
    max_num = 0
    for tap in TAP.query.all():
        n, a = _parse_ctrl(tap.ctrl_numero)
        if n is not None and a == ano and n > max_num:
            max_num = n
    if max_num == 0:
        base = _get_setting(f"ctrl_base_{ano}")
        if base:
            try:
                max_num = max(0, int(base) - 1)
            except ValueError:
                pass
    return f"{max_num + 1}/{ano}"


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
    ano = _ano_suffix_atual()
    ctrl_base = _get_setting(f"ctrl_base_{ano}")
    return render_template(
        "tap/list.html",
        taps=taps,
        ano_atual=ano,
        ctrl_base=ctrl_base,
        proximo=proximo_ctrl(),
    )


@tap_bp.route("/ctrl-base", methods=["POST"])
@tap_access_required
def set_ctrl_base():
    """Somente admin pode definir o primeiro CTRL da numeração do ano."""
    if not current_user.is_admin():
        flash("Apenas administradores podem alterar a numeração base.", "error")
        return redirect(url_for("tap.list_taps"))
    base_ctrl = request.form.get("base_ctrl", "").strip()
    n, a = _parse_ctrl(base_ctrl)
    if n is None or not a:
        flash("Informe um CTRL válido no formato NNNN/AA (ex: 2577/26).", "error")
        return redirect(url_for("tap.list_taps"))
    _set_setting(f"ctrl_base_{a}", str(n))
    flash(f"Numeração base definida: {n}/{a}. Próximo sugerido: {proximo_ctrl()}.", "success")
    return redirect(url_for("tap.list_taps"))


@tap_bp.route("/novo", methods=["GET", "POST"])
@tap_access_required
def novo():
    if request.method == "POST":
        return _salvar(None)
    # prepara uma TAP em memória só para renderizar o form com 15 linhas vazias
    tap_stub = TAP(
        ctrl_numero=proximo_ctrl(), cliente="", status_proposta="Aguardando",
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
    def _parse_data(valor: str):
        valor = (valor or "").strip()
        if not valor:
            return None
        try:
            datetime.strptime(valor, "%Y-%m-%d")
            return valor
        except ValueError:
            return None

    itens_por_ordem = {i.ordem: i for i in tap.itens}
    for ordem in range(1, len(TAP_ENTREGAVEIS) + 1):
        qtd = _parse_float(request.form.get(f"item_{ordem}_qtd"))
        tempo = _parse_float(request.form.get(f"item_{ordem}_tempo"))
        pct_ui = _parse_float(request.form.get(f"item_{ordem}_pct"))
        inicio = _parse_data(request.form.get(f"item_{ordem}_inicio"))
        fim = _parse_data(request.form.get(f"item_{ordem}_fim"))
        item = itens_por_ordem.get(ordem)
        if not item:
            continue
        item.qtd_recursos = qtd
        item.tempo = tempo
        item.percentual_correcao = pct_ui / 100.0
        item.inicio_atividade = inicio
        item.fim_atividade = fim

    db.session.commit()
    flash("TAP salva com sucesso.", "success")
    return redirect(url_for("tap.list_taps"))


@tap_bp.route("/<int:tap_id>/excluir", methods=["POST"])
@tap_access_required
def excluir(tap_id: int):
    tap = TAP.query.get_or_404(tap_id)
    apagar_rda = request.form.get("apagar_rda") == "1"

    # Vínculo por string: RDARecord.cliente == tap.rotulo_cliente ("Cliente / CTRL Nº").
    qtd_rdas = RDARecord.query.filter_by(cliente=tap.rotulo_cliente).count()
    if qtd_rdas > 0 and not apagar_rda:
        flash(
            f"Não é possível excluir a TAP '{tap.rotulo_cliente}': "
            f"existem {qtd_rdas} registro(s) do RDA referenciando este projeto. "
            f"Confirme a exclusão dos registros do RDA para prosseguir.",
            "error",
        )
        return redirect(url_for("tap.list_taps"))

    rotulo = tap.rotulo_cliente
    if apagar_rda and qtd_rdas > 0:
        RDARecord.query.filter_by(cliente=rotulo).delete(synchronize_session=False)

    db.session.delete(tap)
    db.session.commit()
    if apagar_rda and qtd_rdas > 0:
        flash(f"TAP excluída e {qtd_rdas} registro(s) do RDA apagado(s).", "success")
    else:
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
    def _dd_mm_aa(iso):
        if not iso:
            return ""
        try:
            return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return iso

    writer.writerow([
        "ID", "Entregável", "Qtd. Rec.", "Tempo (h)", "Valor Total (R$)",
        "% Correção", "Valor Corrigido (R$)", "Início Atividade", "Fim Atividade",
    ])
    for it in tap.itens:
        writer.writerow([
            it.ordem, it.entregavel,
            f"{it.qtd_recursos:g}".replace(".", ","),
            f"{it.tempo:g}".replace(".", ","),
            f"{it.valor_total:.2f}".replace(".", ","),
            f"{(it.percentual_correcao or 0) * 100:g}".replace(".", ",") + "%",
            f"{it.valor_total_corrigido:.2f}".replace(".", ","),
            _dd_mm_aa(it.inicio_atividade),
            _dd_mm_aa(it.fim_atividade),
        ])
    writer.writerow([])
    writer.writerow(["", "", "", "TOTAL SEM CORREÇÃO", f"{tap.valor_total:.2f}".replace(".", ","), "", "", "", ""])
    writer.writerow(["", "", "", "TOTAL COM CORREÇÃO", "", "", f"{tap.valor_total_corrigido:.2f}".replace(".", ","), "", ""])

    response = make_response("\ufeff" + output.getvalue())
    response.headers["Content-Type"] = "text/csv; charset=utf-8"
    response.headers["Content-Disposition"] = f"attachment; filename=tap_{tap.ctrl_numero}.csv"
    return response


@tap_bp.route("/<int:tap_id>/export/pdf")
@tap_access_required
def export_pdf_detalhe(tap_id: int):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    tap = TAP.query.get_or_404(tap_id)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=8, leading=10)
    head_style = ParagraphStyle("head", parent=styles["BodyText"], fontSize=8, textColor=colors.white, alignment=1)

    def _dd_mm_aa(iso):
        if not iso:
            return ""
        try:
            return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return iso

    titulo = Paragraph("<b>Termo de Abertura de Projeto (TAP)</b>", styles["Title"])
    cab = Paragraph(
        f"<b>CTRL Nº:</b> {tap.ctrl_numero} &nbsp;&nbsp;&nbsp; "
        f"<b>Cliente:</b> {tap.cliente}<br/>"
        f"<b>Status da Proposta:</b> {tap.status_proposta} &nbsp;&nbsp;&nbsp; "
        f"<b>HH:</b> {_fmt_brl(tap.hh_valor)}<br/>"
        f"<b>Gerado em:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        styles["Normal"],
    )

    headers = ["ID", "Entregável", "Qtd", "Tempo", "Valor Total", "% Corr.", "Valor Corrigido", "Início Ativ.", "Fim Ativ."]
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
            Paragraph(_dd_mm_aa(it.inicio_atividade), cell_style),
            Paragraph(_dd_mm_aa(it.fim_atividade), cell_style),
        ])
    data.append([
        Paragraph("", cell_style), Paragraph("<b>Totais</b>", cell_style),
        Paragraph("", cell_style), Paragraph("", cell_style),
        Paragraph(f"<b>{_fmt_brl(tap.valor_total)}</b>", cell_style),
        Paragraph("", cell_style),
        Paragraph(f"<b>{_fmt_brl(tap.valor_total_corrigido)}</b>", cell_style),
        Paragraph("", cell_style),
        Paragraph("", cell_style),
    ])

    col_widths = [10*mm, 80*mm, 15*mm, 15*mm, 32*mm, 15*mm, 32*mm, 28*mm, 28*mm]
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
