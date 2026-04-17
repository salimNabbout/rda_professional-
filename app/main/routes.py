from datetime import datetime
from collections import defaultdict
import csv
from io import StringIO, BytesIO

from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify, make_response, send_file
from flask_login import login_required, current_user

from app.extensions import db
from app.models import RDARecord


main_bp = Blueprint("main", __name__)


def _minutos_intervalo(inicio: str, fim: str) -> int:
    """Retorna duração em minutos. Strings vazias/nulas = 0."""
    if not inicio or not fim:
        return 0
    try:
        i = datetime.strptime(inicio, "%H:%M")
        f = datetime.strptime(fim, "%H:%M")
    except ValueError:
        return 0
    if f < i:
        raise ValueError("A Hora Final não pode ser menor que a Hora Início.")
    return int((f - i).total_seconds() // 60)


def calcular_duracao_total(hi_m: str, hf_m: str, hi_t: str, hf_t: str) -> str:
    total = _minutos_intervalo(hi_m, hf_m) + _minutos_intervalo(hi_t, hf_t)
    return f"{total // 60:02d}:{total % 60:02d}"


def _minutos_registro(r: RDARecord) -> int:
    try:
        h, m = r.duracao.split(":")
        return int(h) * 60 + int(m)
    except (ValueError, AttributeError):
        return 0


def horas_totais(records) -> str:
    total = sum(_minutos_registro(r) for r in records)
    return f"{total // 60:02d}:{total % 60:02d}"


def formatar_data_br(iso: str) -> str:
    """YYYY-MM-DD -> dd/mm/aaaa. Retorna o valor original se não for ISO."""
    if not iso:
        return ""
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return iso


def base_query_for_user():
    if current_user.role in ["gestor", "admin"]:
        return RDARecord.query
    return RDARecord.query.filter_by(user_id=current_user.id)


def aplicar_filtros(query):
    f_data = request.args.get("f_data", "").strip()
    f_cliente = request.args.get("f_cliente", "").strip()
    f_colaborador = request.args.get("f_colaborador", "").strip()

    if f_data:
        query = query.filter(RDARecord.data == f_data)
    if f_cliente:
        query = query.filter(RDARecord.cliente.ilike(f"%{f_cliente}%"))
    if f_colaborador:
        query = query.filter(RDARecord.colaborador.ilike(f"%{f_colaborador}%"))

    return query, {
        "f_data": f_data,
        "f_cliente": f_cliente,
        "f_colaborador": f_colaborador,
    }


def stats_mes_atual():
    """Retorna dados agregados do mês corrente para a dashboard."""
    hoje = datetime.utcnow()
    prefixo = hoje.strftime("%Y-%m")

    query = base_query_for_user().filter(RDARecord.data.like(f"{prefixo}-%"))
    records = query.all()

    por_projeto = defaultdict(int)
    for r in records:
        por_projeto[r.cliente or "(sem cliente)"] += _minutos_registro(r)

    total_min = sum(por_projeto.values())

    labels = list(por_projeto.keys())
    horas_decimais = [round(m / 60, 2) for m in por_projeto.values()]

    return {
        "mes_label": hoje.strftime("%m/%Y"),
        "total_horas_mes": f"{total_min // 60:02d}:{total_min % 60:02d}",
        "labels_projeto": labels,
        "horas_por_projeto": horas_decimais,
        "total_registros_mes": len(records),
    }


def _form_inicial_vazio():
    """Retorna dicionário com nome pré-preenchido do usuário."""
    return {
        "id": "",
        "colaborador": current_user.display_name,
        "cliente": "",
        "data": "",
        "hora_inicio_manha": "",
        "hora_final_manha": "",
        "hora_inicio_tarde": "",
        "hora_final_tarde": "",
        "realizado": "",
        "status_rda": "Andamento",
    }


@main_bp.route("/")
@login_required
def index():
    query = base_query_for_user()
    query, filters = aplicar_filtros(query)
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio_manha.desc(), RDARecord.id.desc()).all()

    return render_template(
        "main/dashboard.html",
        records=records,
        total_registros=len(records),
        total_horas=horas_totais(records),
        filters=filters,
        form=_form_inicial_vazio(),
        editing=False,
        stats=stats_mes_atual(),
        formatar_data_br=formatar_data_br,
    )


@main_bp.route("/save", methods=["POST"])
@login_required
def save_record():
    record_id = request.form.get("id", "").strip()

    payload = {
        "colaborador": request.form.get("colaborador", "").strip(),
        "cliente": request.form.get("cliente", "").strip(),
        "data": request.form.get("data", "").strip(),
        "hora_inicio_manha": request.form.get("hora_inicio_manha", "").strip(),
        "hora_final_manha": request.form.get("hora_final_manha", "").strip(),
        "hora_inicio_tarde": request.form.get("hora_inicio_tarde", "").strip(),
        "hora_final_tarde": request.form.get("hora_final_tarde", "").strip(),
        "realizado": request.form.get("realizado", "").strip(),
        "status_rda": request.form.get("status_rda", "Andamento").strip(),
    }

    try:
        obrigatorios = [payload["colaborador"], payload["cliente"], payload["data"], payload["realizado"]]
        if not all(obrigatorios):
            raise ValueError("Preencha Colaborador, Cliente, Data e o que foi realizado.")

        # Campos de hora em branco viram "00:00"
        for campo in ["hora_inicio_manha", "hora_final_manha", "hora_inicio_tarde", "hora_final_tarde"]:
            if not payload[campo]:
                payload[campo] = "00:00"

        payload["duracao"] = calcular_duracao_total(
            payload["hora_inicio_manha"], payload["hora_final_manha"],
            payload["hora_inicio_tarde"], payload["hora_final_tarde"],
        )

        if payload["duracao"] == "00:00":
            raise ValueError("Informe pelo menos um período (manhã ou tarde).")

        if record_id:
            record = RDARecord.query.filter_by(id=int(record_id)).first_or_404()
            if current_user.role == "colaborador" and record.user_id != current_user.id:
                raise ValueError("Você não tem permissão para editar este registro.")
            for key, value in payload.items():
                setattr(record, key, value)
        else:
            record = RDARecord(user_id=current_user.id, **payload)
            db.session.add(record)

        db.session.commit()
        flash("Registro salvo com sucesso.", "success")
        return redirect(url_for("main.index"))

    except ValueError as exc:
        flash(str(exc), "error")
        query = base_query_for_user()
        query, filters = aplicar_filtros(query)
        records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio_manha.desc(), RDARecord.id.desc()).all()
        form_data = {"id": record_id, **payload}
        return render_template(
            "main/dashboard.html",
            records=records,
            total_registros=len(records),
            total_horas=horas_totais(records),
            filters=filters,
            form=form_data,
            editing=bool(record_id),
            stats=stats_mes_atual(),
            formatar_data_br=formatar_data_br,
        )


@main_bp.route("/edit/<int:record_id>")
@login_required
def edit_record(record_id: int):
    record = RDARecord.query.filter_by(id=record_id).first_or_404()
    if current_user.role == "colaborador" and record.user_id != current_user.id:
        flash("Você não tem permissão para editar este registro.", "error")
        return redirect(url_for("main.index"))

    query = base_query_for_user()
    query, filters = aplicar_filtros(query)
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio_manha.desc(), RDARecord.id.desc()).all()

    return render_template(
        "main/dashboard.html",
        records=records,
        total_registros=len(records),
        total_horas=horas_totais(records),
        filters=filters,
        form=record,
        editing=True,
        stats=stats_mes_atual(),
        formatar_data_br=formatar_data_br,
    )


@main_bp.route("/delete/<int:record_id>", methods=["POST"])
@login_required
def delete_record(record_id: int):
    record = RDARecord.query.filter_by(id=record_id).first_or_404()
    if current_user.role == "colaborador" and record.user_id != current_user.id:
        flash("Você não tem permissão para excluir este registro.", "error")
        return redirect(url_for("main.index"))

    db.session.delete(record)
    db.session.commit()
    flash("Registro excluído.", "success")
    return redirect(url_for("main.index"))


@main_bp.route("/api/records")
@login_required
def api_records():
    query = base_query_for_user()
    query, filters = aplicar_filtros(query)
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio_manha.desc(), RDARecord.id.desc()).all()

    return jsonify({
        "user": current_user.username,
        "role": current_user.role,
        "filters": filters,
        "total_registros": len(records),
        "total_horas": horas_totais(records),
        "records": [
            {
                "id": r.id,
                "colaborador": r.colaborador,
                "cliente": r.cliente,
                "data": formatar_data_br(r.data),
                "hora_inicio_manha": r.hora_inicio_manha,
                "hora_final_manha": r.hora_final_manha,
                "hora_inicio_tarde": r.hora_inicio_tarde,
                "hora_final_tarde": r.hora_final_tarde,
                "duracao": r.duracao,
                "realizado": r.realizado,
                "status_rda": r.status_rda,
                "owner_username": r.owner.username if r.owner else None,
            }
            for r in records
        ],
    })


@main_bp.route("/export/csv")
@login_required
def export_csv():
    query = base_query_for_user()
    query, _ = aplicar_filtros(query)
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio_manha.desc(), RDARecord.id.desc()).all()

    output = StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow([
        "Colaborador", "Cliente", "Data",
        "Manhã Início", "Manhã Final", "Tarde Início", "Tarde Final",
        "Duração", "O que foi Realizado", "Status", "Dono do Registro",
    ])

    for item in records:
        writer.writerow([
            item.colaborador,
            item.cliente,
            formatar_data_br(item.data),
            item.hora_inicio_manha,
            item.hora_final_manha,
            item.hora_inicio_tarde,
            item.hora_final_tarde,
            item.duracao,
            item.realizado,
            item.status_rda,
            item.owner.username if item.owner else "",
        ])

    response = make_response("\ufeff" + output.getvalue())
    response.headers["Content-Type"] = "text/csv; charset=utf-8"
    response.headers["Content-Disposition"] = "attachment; filename=rda_relatorio.csv"
    return response


@main_bp.route("/export/pdf")
@login_required
def export_pdf():
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    query = base_query_for_user()
    query, _ = aplicar_filtros(query)
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio_manha.desc(), RDARecord.id.desc()).all()

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    titulo = Paragraph("<b>Relatório Diário de Atividade (RDA)</b>", styles["Title"])
    gerado = Paragraph(
        f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} — Total: {len(records)} registros — "
        f"Horas: {horas_totais(records)}",
        styles["Normal"],
    )

    cell_style = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=7, leading=9)
    head_style = ParagraphStyle("head", parent=styles["BodyText"], fontSize=7, leading=9, textColor=colors.white, alignment=1)

    headers = ["Colaborador", "Cliente", "Data", "Manhã Ini", "Manhã Fim", "Tarde Ini", "Tarde Fim", "Duração", "Realizado", "Status"]
    data = [[Paragraph(h, head_style) for h in headers]]

    for r in records:
        data.append([
            Paragraph(r.colaborador or "", cell_style),
            Paragraph(r.cliente or "", cell_style),
            Paragraph(formatar_data_br(r.data), cell_style),
            Paragraph(r.hora_inicio_manha or "", cell_style),
            Paragraph(r.hora_final_manha or "", cell_style),
            Paragraph(r.hora_inicio_tarde or "", cell_style),
            Paragraph(r.hora_final_tarde or "", cell_style),
            Paragraph(r.duracao or "", cell_style),
            Paragraph(r.realizado or "", cell_style),
            Paragraph(r.status_rda or "", cell_style),
        ])

    col_widths = [30*mm, 28*mm, 18*mm, 16*mm, 16*mm, 16*mm, 16*mm, 16*mm, 70*mm, 22*mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2d5a")),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
    ]))

    doc.build([titulo, Spacer(1, 4 * mm), gerado, Spacer(1, 4 * mm), table])
    buffer.seek(0)

    return send_file(
        buffer, as_attachment=True,
        download_name=f"rda_relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mimetype="application/pdf",
    )
