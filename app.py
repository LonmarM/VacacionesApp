import os
import json
from datetime import date, datetime, timedelta
import pandas as pd
import holidays
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session

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

        empleados[nombre] = {
            "nombre": nombre,
            "codigo": codigo,
            "area": "General",
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


def calcular_fin(inicio, dias_disponibles):
    actuales = inicio
    cont = 0

    while cont < dias_disponibles:
        actuales += timedelta(days=1)

        if actuales.weekday() == 6:
            continue

        if actuales in festivos:
            continue

        cont += 1

    return actuales


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

    for r in requests:
        if r["nombre"] in employees:
            e = employees[r["nombre"]]
            e["inicio_vacaciones"] = r["inicio"]
            e["fin_vacaciones"] = r["fin"]
            e["dias_solicitados_req"] = r["dias"]
            e["estado"] = r["estado"]

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
# VALIDACIÓN FECHA
# ------------------------------------------------------------------

@app.route("/calcular_disponibilidad", methods=["POST"])
def calcular_disponibilidad():
    nombre = request.form.get("nombre_solicitud")
    inicio = request.form.get("inicio_fecha")

    if not nombre or not inicio:
        return {"permitido": False, "mensaje": "Datos incompletos"}

    inicio = date.fromisoformat(inicio)

    empleados = load_employees()
    empleado = empleados.get(nombre)

    if not empleado:
        return {"permitido": False, "mensaje": "Empleado no encontrado"}

    dias_disponibles = int(empleado["dias_disponibles"])
    fecha_15 = empleado["fecha_15_dias"]

    if dias_disponibles < 15 and inicio < fecha_15:
        return {
            "permitido": False,
            "mensaje": f"Solo puedes solicitar vacaciones a partir del {fecha_15}."
        }

    return {"permitido": True, "mensaje": "Solicitud válida"}


# ------------------------------------------------------------------
# ENVIAR SOLICITUD
# ------------------------------------------------------------------

@app.route("/submit", methods=["POST"])
def submit_request():
    nombre = request.form["nombre_solicitud"]
    inicio = date.fromisoformat(request.form["inicio_fecha"])

    empleados = load_employees()
    empleado = empleados.get(nombre)

    if not empleado:
        return "Empleado no encontrado", 400

    dias_disponibles = int(empleado["dias_disponibles"])
    fecha_15 = empleado["fecha_15_dias"]

    if dias_disponibles < 15 and inicio < fecha_15:
        return redirect(url_for(
            "employee_request_form",
            message=f"Solo puedes solicitar vacaciones a partir del {fecha_15}."
        ))

    fin = calcular_fin(inicio, dias_disponibles)

    reqs = load_requests()
    reqs.append({
        "nombre": nombre,
        "area": "General",
        "dias": dias_disponibles,
        "inicio": str(inicio),
        "fin": str(fin),
        "estado": "Pending",
        "solicitado_en": datetime.now().isoformat()
    })

    save_requests(reqs)

    return redirect(url_for("employee_request_form", message="Solicitud enviada correctamente"))


# ------------------------------------------------------------------
# ADMIN → APROBAR / RECHAZAR
# ------------------------------------------------------------------

@app.route("/action", methods=["POST"])
def action_request():
    nombre = request.form["nombre"]
    inicio = request.form["inicio"]
    action = request.form["action"]

    reqs = load_requests()

    for r in reqs:
        if r["nombre"] == nombre and r["inicio"] == inicio:
            r["estado"] = "Approved" if action == "Aprobado" else "Rejected"

    save_requests(reqs)

    return redirect(url_for("admin_view"))


# ------------------------------------------------------------------
# PDF
# ------------------------------------------------------------------

@app.route("/pdf")
def descargar_pdf():
    reqs = load_requests()
    aprobadas = [r for r in reqs if r["estado"] == "Approved"]

    path = os.path.join(app.config["PDF_FOLDER"], "vacaciones_aprobadas.pdf")

    doc = SimpleDocTemplate(path)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>Vacaciones Aprobadas</b>", styles["Title"]))
    story.append(Spacer(1, 20))

    for r in aprobadas:
        story.append(Paragraph(
            f"{r['nombre']} — {r['inicio']} a {r['fin']} — {r['dias']} días",
            styles["Normal"]
        ))
        story.append(Spacer(1, 10))

    doc.build(story)

    return send_from_directory(app.config["PDF_FOLDER"], "vacaciones_aprobadas.pdf", as_attachment=True)


# ------------------------------------------------------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
