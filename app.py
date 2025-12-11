import os
import json
from datetime import date, datetime, timedelta
import pandas as pd
import holidays
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session, jsonify

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.secret_key = "ClaveSuperSecreta123"   # CAMBIAR

# Carpetas
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["PDF_FOLDER"] = "pdf"

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["PDF_FOLDER"], exist_ok=True)

# Archivos
EMPLOYEE_FILE = os.path.join(app.config["UPLOAD_FOLDER"], "LibroVacaciones.xlsx")
REQUEST_FILE = os.path.join(app.config["UPLOAD_FOLDER"], "requests.json")

festivos = holidays.Colombia()

ADMIN_PASSWORD = "admin1234"   # CAMBIAR


# ------------------------------------------------------------------
# UTILIDADES
# ------------------------------------------------------------------

def iso_to_ddmmyyyy(iso_str):
    if not iso_str:
        return ""
    try:
        d = date.fromisoformat(iso_str)
        return d.strftime("%d/%m/%Y")
    except:
        return iso_str

def ddmmyyyy_to_date(s):
    # acepta dd/mm/yyyy o yyyy-mm-dd
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except:
        try:
            return date.fromisoformat(s)
        except:
            return None

def obtener_pagadas_hasta(df, index):
    current_name = df.loc[index, "empleado"]

    pagadas = pd.to_datetime(df.loc[index, "pagadas_hasta"], errors="coerce")
    if pd.notna(pagadas):
        return pagadas.date()

    for i in range(index - 1, -1, -1):
        if df.loc[i, "empleado"] == current_name:
            prev = pd.to_datetime(df.loc[i, "pagadas_hasta"], errors="coerce")
            if pd.notna(prev):
                return prev.date()

    ingreso = pd.to_datetime(df.loc[index, "fecha_ingreso"], errors="coerce").date()
    return ingreso


def calcular_fecha_proyectada_15_dias(pagadas_hasta, dias_sln_previos):
    dias_requeridos = 360
    fecha_proyectada = pagadas_hasta + timedelta(days=dias_requeridos + dias_sln_previos)
    return fecha_proyectada


def load_employees():
    # Si no existe el excel, devolvemos un ejemplo vacío para evitar crash.
    if not os.path.exists(EMPLOYEE_FILE):
        # retornar diccionario vacío
        return {}

    df = pd.read_excel(EMPLOYEE_FILE, header=0, skiprows=[0])

    df.columns = [
        "codigo", "empleado", "fecha_ingreso", "dias_trabajados",
        "dias_sln", "dias_neto", "dias_derecho", "dias_otorgados",
        "dias_pendientes", "pagadas_hasta", "dias_vacaciones",
        "tipo", "periodo", "valor_pagado", "fecha_inicial", "fecha_final"
    ]

    today = date.today()
    empleados = {}

    for idx, row in df.iterrows():
        nombre = row["empleado"]
        codigo = row["codigo"]

        pagadas = obtener_pagadas_hasta(df, idx)
        ingreso = pd.to_datetime(row["fecha_ingreso"], errors="coerce").date()

        dias_sln = int(row["dias_sln"]) if not pd.isna(row["dias_sln"]) else 0

        dias_laborados = (today - pagadas).days - dias_sln
        dias_laborados = max(dias_laborados, 0)

        dias_generados = dias_laborados * (15 / 360)

        usados = float(row["dias_otorgados"]) if not pd.isna(row["dias_otorgados"]) else 0
        disponibles = round(dias_generados - usados, 2)

        fecha_15 = calcular_fecha_proyectada_15_dias(pagadas, dias_sln)

        # Si tu archivo de empleados trae 'area', podrías mapearlo aquí.
        area = row["tipo"] if ("tipo" in df.columns and not pd.isna(row["tipo"])) else "General"

        empleados[nombre] = {
            "nombre": nombre,
            "codigo": codigo,
            "area": area,
            "dias_generados": round(dias_generados, 2),
            "dias_tomados": usados,
            "dias_disponibles": disponibles,
            "fecha_15_dias": fecha_15,
            "inicio_vacaciones": "",
            "fin_vacaciones": "",
            "dias_solicitados_req": 0,
            "estado": "Sin solicitud"
        }

    return empleados


def load_requests():
    if not os.path.exists(REQUEST_FILE):
        return []
    with open(REQUEST_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_requests(reqs):
    with open(REQUEST_FILE, "w", encoding="utf-8") as f:
        json.dump(reqs, f, indent=4, ensure_ascii=False)


def calcular_fin(inicio, dias_habiles=15):
    # Calcula la fecha final sumando dias_habiles días hábiles
    actuales = inicio
    cont = 0

    while cont < dias_habiles:
        actuales += timedelta(days=1)

        # Excluir domingos
        if actuales.weekday() == 6:
            continue

        # Excluir festivos
        if actuales in festivos:
            continue

        cont += 1

    return actuales


def rango_solapa(inicio1, fin1, inicio2, fin2):
    # todos date objects
    return not (fin1 < inicio2 or fin2 < inicio1)


# ------------------------------------------------------------------
# LOGIN ADMIN
# ------------------------------------------------------------------

@app.route("/admin_login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        password = request.form.get("password")
        if password == ADMIN_PASSWORD:
            session["admin_authenticated"] = True
            return redirect(url_for("admin_view"))
        else:
            return render_template("admin_login.html", error="Contraseña incorrecta")

    return render_template("admin_login.html")


@app.route("/logout_admin")
def logout_admin():
    session.pop("admin_authenticated", None)
    return redirect(url_for("employee_request_form"))


# ------------------------------------------------------------------
# RUTAS PRINCIPALES
# ------------------------------------------------------------------

@app.route("/")
def employee_request_form():
    employees = load_employees()
    requests = load_requests()

    # Vincular solicitudes a empleados (para mostrar en la tabla)
    for r in requests:
        if r["nombre"] in employees:
            e = employees[r["nombre"]]
            e["inicio_vacaciones"] = iso_to_ddmmyyyy(r.get("inicio"))
            e["fin_vacaciones"] = iso_to_ddmmyyyy(r.get("fin"))
            e["dias_solicitados_req"] = r.get("dias", 0)
            e["estado"] = r.get("estado", "Sin solicitud")

    msg = request.args.get("message")

    return render_template(
        "resultados.html",
        results_type="request",
        data=list(employees.values()),
        eventos=json.dumps(requests),
        is_admin=False,
        message=msg
    )


@app.route("/admin")
def admin_view():
    if not session.get("admin_authenticated"):
        return redirect(url_for("admin_login"))

    requests = load_requests()
    return render_template(
        "resultados.html",
        results_type="admin",
        data=requests,
        eventos=json.dumps(requests),
        is_admin=True
    )


# ------------------------------------------------------------------
# ENDPOINT EVENTS (FullCalendar)
# ------------------------------------------------------------------

@app.route("/events")
def events():
    # Devuelve eventos en formato que FullCalendar espera.
    reqs = load_requests()
    events = []

    for r in reqs:
        estado = r.get('estado')

        # Solo incluir aprobadas
        if estado not in ["Approved", "Aprobado"]:
            continue

        events.append({
            "title": f"{r.get('nombre')} (Aprobado)",
            "start": r.get('inicio'),
            "end": r.get('fin'),
            "extendedProps": {
                "nombre": r.get("nombre"),
                "area": r.get("area"),
                "estado": estado,
                "dias": r.get("dias")
            },
            "color": "#10b981"  # verde
        })

    return jsonify(events)

# ------------------------------------------------------------------
# VALIDACIÓN FECHA
# ------------------------------------------------------------------

@app.route("/calcular_disponibilidad", methods=["POST"])
def calcular_disponibilidad():
    nombre = request.form.get("nombre_solicitud")
    inicio_raw = request.form.get("inicio_fecha")

    if not nombre or not inicio_raw:
        return jsonify({"permitido": False, "mensaje": "Datos incompletos"})

    # aceptar dd/mm/yyyy o yyyy-mm-dd
    inicio = ddmmyyyy_to_date(inicio_raw) if "/" in inicio_raw else date.fromisoformat(inicio_raw)

    if inicio is None:
        return jsonify({"permitido": False, "mensaje": "Fecha inválida"})

    # Bloquear enero
    if inicio.month == 1:
        return jsonify({"permitido": False, "mensaje": "No se pueden solicitar vacaciones con inicio en enero."})

    # Solo exactamente 15 días
    dias_solicitados = 15

    # Revisar solapamientos por área con solicitudes aprobadas
    employees = load_employees()
    empleado = employees.get(nombre)
    if not empleado:
        return jsonify({"permitido": False, "mensaje": "Empleado no encontrado"})

    area = empleado.get("area", "General")
    fin = calcular_fin(inicio, dias_habiles=dias_solicitados)

    reqs = load_requests()
    for r in reqs:
        if r.get("area") == area and (r.get("estado") in ("Approved", "Aprobado")):
            # comparar rangos
            r_inicio = date.fromisoformat(r.get("inicio"))
            r_fin = date.fromisoformat(r.get("fin"))
            if rango_solapa(inicio, fin, r_inicio, r_fin):
                return jsonify({"permitido": False, "mensaje": f"Ya existe una persona del área {area} de vacaciones en ese periodo."})

    # También validar fecha_15 (tu regla previa)
    fecha_15 = empleado["fecha_15_dias"]
    if empleado["dias_disponibles"] < 15 and inicio < fecha_15:
        return jsonify({"permitido": False, "mensaje": f"Solo puedes solicitar vacaciones a partir del {fecha_15.strftime('%d/%m/%Y')}."})

    return jsonify({"permitido": True, "mensaje": "Solicitud válida", "dias": dias_solicitados, "fin": fin.isoformat()})


# ------------------------------------------------------------------
# ENVIAR SOLICITUD
# ------------------------------------------------------------------

@app.route("/submit", methods=["POST"])
def submit_request():
    nombre = request.form["nombre_solicitud"]
    # aceptar dd/mm/yyyy o yyyy-mm-dd
    inicio_field = request.form["inicio_fecha"]
    inicio = ddmmyyyy_to_date(inicio_field) if "/" in inicio_field else date.fromisoformat(inicio_field)

    employees = load_employees()
    empleado = employees.get(nombre)

    if not empleado:
        return redirect(url_for("employee_request_form", message="Empleado no encontrado"))

    # check january
    if inicio.month == 1:
        return redirect(url_for("employee_request_form", message="No se pueden solicitar vacaciones con inicio en enero."))

    dias_solicitados = 15
    fin = calcular_fin(inicio, dias_habiles=dias_solicitados)

    # verificar solapamiento por area contra aprobadas
    area = empleado.get("area", "General")
    reqs = load_requests()
    for r in reqs:
        if r.get("area") == area and (r.get("estado") in ("Approved", "Aprobado")):
            r_inicio = date.fromisoformat(r.get("inicio"))
            r_fin = date.fromisoformat(r.get("fin"))
            if rango_solapa(inicio, fin, r_inicio, r_fin):
                return redirect(url_for("employee_request_form", message=f"Error: otra persona del área {area} ya tiene vacaciones en ese periodo."))

    # Guardar solicitud con ISO dates, dias = 15
    reqs.append({
        "nombre": nombre,
        "area": area,
        "dias": dias_solicitados,
        "inicio": inicio.isoformat(),
        "fin": fin.isoformat(),
        "estado": "Pending",
        "solicitado_en": datetime.now().isoformat()
    })

    save_requests(reqs)

    return redirect(url_for("employee_request_form", message="Solicitud enviada correctamente"))


# ------------------------------------------------------------------
# ADMIN → APROBAR / RECHAZAR (AJAX-friendly)
# ------------------------------------------------------------------

@app.route("/action", methods=["POST"])
def action_request():
    nombre = request.form["nombre"]
    inicio = request.form["inicio"]
    action = request.form["action"]

    # recibir inicio como ISO
    reqs = load_requests()
    target = None
    for r in reqs:
        if r["nombre"] == nombre and r["inicio"] == inicio:
            target = r
            break

    if not target:
        return jsonify({"success": False, "mensaje": "Solicitud no encontrada."})

    # Si se quiere aprobar: verificar solapamientos por área con otras aprobadas
    if action == "Aprobado":
        area = target.get("area", "General")
        inicio_dt = date.fromisoformat(target.get("inicio"))
        fin_dt = date.fromisoformat(target.get("fin"))

        for r in reqs:
            if r is target:
                continue
            if r.get("area") == area and (r.get("estado") in ("Approved", "Aprobado")):
                r_inicio = date.fromisoformat(r.get("inicio"))
                r_fin = date.fromisoformat(r.get("fin"))
                if rango_solapa(inicio_dt, fin_dt, r_inicio, r_fin):
                    return jsonify({"success": False, "mensaje": f"No se puede aprobar: solapa con {r.get('nombre')} del área {area}."})

        target["estado"] = "Approved"
    else:
        # Denegado
        target["estado"] = "Rejected" if action == "Denegado" else action

    save_requests(reqs)

    return jsonify({"success": True, "mensaje": f"Solicitud {action} correctamente."})


# ------------------------------------------------------------------
# PDF
# ------------------------------------------------------------------

@app.route("/pdf")
def descargar_pdf():
    reqs = load_requests()
    aprobadas = [r for r in reqs if r["estado"] in ("Approved", "Aprobado")]

    path = os.path.join(app.config["PDF_FOLDER"], "vacaciones_aprobadas.pdf")

    doc = SimpleDocTemplate(path)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>Vacaciones Aprobadas</b>", styles["Title"]))
    story.append(Spacer(1, 20))

    for r in aprobadas:
        inicio_dd = iso_to_ddmmyyyy(r.get("inicio"))
        fin_dd = iso_to_ddmmyyyy(r.get("fin"))
        story.append(Paragraph(
            f"{r['nombre']} — {inicio_dd} a {fin_dd} — {r['dias']} días",
            styles["Normal"]
        ))
        story.append(Spacer(1, 10))

    doc.build(story)

    return send_from_directory(app.config["PDF_FOLDER"], "vacaciones_aprobadas.pdf", as_attachment=True)


# ------------------------------------------------------------------

if __name__ == '__main__':
    from waitress import serve
    serve(app, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))


