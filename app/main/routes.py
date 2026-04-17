from datetime import datetime
import csv
from io import StringIO

from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify, make_response
from flask_login import login_required, current_user

from app.extensions import db
from app.models import RDARecord


main_bp = Blueprint("main", __name__)


def calcular_duracao(hora_inicio: str, hora_final: str) -> str:
    inicio = datetime.strptime(hora_inicio, "%H:%M")
    fim = datetime.strptime(hora_final, "%H:%M")
    if fim < inicio:
        raise ValueError("A Hora Final não pode ser menor que a Hora Início.")
    total_minutos = int((fim - inicio).total_seconds() // 60)
    return f"{total_minutos // 60:02d}:{total_minutos % 60:02d}"


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


def horas_totais(records):
    total = 0
    for item in records:
        h, m = item.duracao.split(":")
        total += int(h) * 60 + int(m)
    return f"{total // 60:02d}:{total % 60:02d}"


@main_bp.route("/")
@login_required
def index():
    query = base_query_for_user()
    query, filters = aplicar_filtros(query)
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio.desc(), RDARecord.id.desc()).all()

    return render_template(
        "main/dashboard.html",
        records=records,
        total_registros=len(records),
        total_horas=horas_totais(records),
        filters=filters,
        form=None,
        editing=False,
    )


@main_bp.route("/save", methods=["POST"])
@login_required
def save_record():
    record_id = request.form.get("id", "").strip()

    payload = {
        "colaborador": request.form.get("colaborador", "").strip(),
        "cliente": request.form.get("cliente", "").strip(),
        "data": request.form.get("data", "").strip(),
        "hora_inicio": request.form.get("hora_inicio", "").strip(),
        "hora_final": request.form.get("hora_final", "").strip(),
        "realizado": request.form.get("realizado", "").strip(),
        "status_rda": request.form.get("status_rda", "Em aberto").strip(),
        "aprovador": request.form.get("aprovador", "").strip(),
        "responsavel_rda": request.form.get("responsavel_rda", "").strip(),
        "periodo_referencia": request.form.get("periodo_referencia", "").strip(),
        "observacoes_aprovacao": request.form.get("observacoes_aprovacao", "").strip(),
    }

    try:
        obrigatorios = [
            payload["colaborador"], payload["cliente"], payload["data"],
            payload["hora_inicio"], payload["hora_final"], payload["realizado"]
        ]
        if not all(obrigatorios):
            raise ValueError("Preencha todos os campos obrigatórios.")

        payload["duracao"] = calcular_duracao(payload["hora_inicio"], payload["hora_final"])

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
        records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio.desc(), RDARecord.id.desc()).all()
        form_data = {"id": record_id, **payload}
        return render_template(
            "main/dashboard.html",
            records=records,
            total_registros=len(records),
            total_horas=horas_totais(records),
            filters=filters,
            form=form_data,
            editing=bool(record_id),
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
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio.desc(), RDARecord.id.desc()).all()

    return render_template(
        "main/dashboard.html",
        records=records,
        total_registros=len(records),
        total_horas=horas_totais(records),
        filters=filters,
        form=record,
        editing=True,
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
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio.desc(), RDARecord.id.desc()).all()

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
                "data": r.data,
                "hora_inicio": r.hora_inicio,
                "hora_final": r.hora_final,
                "duracao": r.duracao,
                "realizado": r.realizado,
                "status_rda": r.status_rda,
                "aprovador": r.aprovador,
                "responsavel_rda": r.responsavel_rda,
                "periodo_referencia": r.periodo_referencia,
                "observacoes_aprovacao": r.observacoes_aprovacao,
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
    records = query.order_by(RDARecord.data.desc(), RDARecord.hora_inicio.desc(), RDARecord.id.desc()).all()

    output = StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow([
        "Nome do Colaborador", "Cliente", "Data", "Hora_Inicio", "Hora_Final", "Duracao",
        "O que foi Realizado", "Status RDA", "Aprovador", "Responsavel RDA", "Periodo Referencia", "Dono do Registro"
    ])

    for item in records:
        writer.writerow([
            item.colaborador,
            item.cliente,
            item.data,
            item.hora_inicio,
            item.hora_final,
            item.duracao,
            item.realizado,
            item.status_rda,
            item.aprovador,
            item.responsavel_rda,
            item.periodo_referencia,
            item.owner.username if item.owner else "",
        ])

    response = make_response("\ufeff" + output.getvalue())
    response.headers["Content-Type"] = "text/csv; charset=utf-8"
    response.headers["Content-Disposition"] = "attachment; filename=rda_relatorio.csv"
    return response
