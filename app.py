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


def load_employees():
    # Si no existe el excel, devolvemos un ejemplo vacío para evitar crash.
    if not os.path.exists(EMPLOYEE_FILE):
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
        # parse fecha_ingreso como date; si falla, dejar None
        try:
            ingreso = pd.to_datetime(row["fecha_ingreso"], errors="coerce").date()
        except:
            ingreso = None

        dias_sln = int(row["dias_sln"]) if not pd.isna(row["dias_sln"]) else 0

        dias_laborados = (today - pagadas).days - dias_sln if pagadas else 0
        dias_laborados = max(dias_laborados, 0)

        dias_generados = dias_laborados * (15 / 360)

        usados = float(row["dias_otorgados"]) if not pd.isna(row["dias_otorgados"]) else 0
        disponibles = round(dias_generados - usados, 2)

        fecha_15 = calcular_fecha_proyectada_15_dias(pagadas, dias_sln) if pagadas else today

        area = row["tipo"] if ("tipo" in df.columns and not pd.isna(row["tipo"])) else "General"

        empleados[nombre] = {
            "nombre": nombre,
            "codigo": codigo,
            "area": area,
            "fecha_ingreso": ingreso,
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
        try:
            return json.load(f)
        except:
            return []


def save_requests(reqs):
    with open(REQUEST_FILE, "w", encoding="utf-8") as f:
        json.dump(reqs, f, indent=4, ensure_ascii=False)


# ------------------ Helpers para cola por antigüedad ------------------

def employees_sorted_by_area(employees):
    """
    Recibe el dict employees (como lo devuelve load_employees) y devuelve
    un dict: { area: [empleado_dict_ordenado_por_fecha_ingreso_asc, ...], ... }
    """
    areas = {}
    for e in employees.values():
        area = e.get("area", "General")
        areas.setdefault(area, []).append(e)

    # ordenar por fecha_ingreso asc (el más antiguo primero). Si no hay fecha, lo colocamos al final.
    for area, lst in areas.items():
        lst.sort(key=lambda x: x.get("fecha_ingreso") or date.max)
    return areas


def has_previous_employees_requested(nombre, employees_dict, reqs):
    """
    Verifica si todos los empleados anteriores (por fecha_ingreso) en el mismo área
    ya tienen una solicitud (cualquier estado en reqs).
    Retorna (ok: bool, mensaje:str)
    """
    empleado = employees_dict.get(nombre)
    if not empleado:
        return False, "Empleado no encontrado."

    area = empleado.get("area", "General")
    # obtener lista ordenada de empleados por area
    areas = employees_sorted_by_area(employees_dict)
    orden = areas.get(area, [])
    # encontrar índice del empleado en esa lista
    idx = next((i for i, e in enumerate(orden) if e["nombre"] == nombre), None)
    if idx is None:
        return False, "Empleado no encontrado en la lista de área."

    # cargar nombres de quienes ya tienen solicitud en reqs
    nombres_con_solicitud = {r["nombre"] for r in reqs}

    # todos los anteriores deben estar en nombres_con_solicitud
    for prev in orden[:idx]:
        if prev["nombre"] not in nombres_con_solicitud:
            return False, f"El empleado anterior ({prev['nombre']}) aún no ha realizado su solicitud. Debes esperar tu turno."
    return True, "OK"


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
    employees = load_employees()  # dict
    reqs = load_requests()

    # Vincular solicitudes a empleados y mantener fechas en ISO en eventos
    eventos_calendario = []
    for r in reqs:
        if r.get("estado") in ["Pending", "Approved", "Aprobado", "Pendiente"]:
            eventos_calendario.append({
                "title": f"{r.get('nombre')} ({r.get('estado')})",
                "start": r.get("inicio"),
                "end": r.get("fin"),
                "status": r.get("estado"),
                "area": r.get("area")
            })
        # vincular para mostrar en tabla
        if r["nombre"] in employees:
            e = employees[r["nombre"]]
            e["inicio_vacaciones"] = iso_to_ddmmyyyy(r.get("inicio"))
            e["fin_vacaciones"] = iso_to_ddmmyyyy(r.get("fin"))
            e["dias_solicitados_req"] = r.get("dias", 0)
            e["estado"] = r.get("estado", "Sin solicitud")

    # Ordenar empleados por fecha_ingreso asc (más antiguo primero)
    empleados_list = list(employees.values())
    empleados_list.sort(key=lambda x: x.get("fecha_ingreso") or date.max)

    msg = request.args.get("message")

    return render_template(
        "resultados.html",
        results_type="request",
        data=empleados_list,
        eventos=json.dumps(eventos_calendario),
        is_admin=False,
        message=msg
    )


@app.route("/admin")
def admin_view():
    if not session.get("admin_authenticated"):
        return redirect(url_for("admin_login"))

    requests = load_requests()

    eventos_calendario = []
    for r in requests:
        if r.get("estado") in ["Pending", "Approved", "Aprobado", "Pendiente"]:
            eventos_calendario.append({
                "title": f"{r.get('nombre')} ({r.get('estado')})",
                "start": r.get("inicio"),
                "end": r.get("fin"),
                "status": r.get("estado"),
                "area": r.get("area")
            })

    return render_template(
        "resultados.html",
        results_type="admin",
        data=requests,
        eventos=json.dumps(eventos_calendario),
        is_admin=True
    )


# ------------------------------------------------------------------
# ENDPOINT EVENTS (FullCalendar)
# ------------------------------------------------------------------

@app.route("/events")
def events():
    reqs = load_requests()
    events = []

    for r in reqs:
        estado = r.get('estado')
        if estado in ["Approved", "Aprobado"]:
            color = "#10b981"
            label = "Aprobado"
        elif estado in ["Pending", "Pendiente"]:
            color = "#f59e0b"
            label = "Pendiente"
        else:
            continue

        events.append({
            "title": f"{r.get('nombre')} ({label})",
            "start": r.get('inicio'),
            "end": r.get('fin'),
            "extendedProps": {
                "nombre": r.get("nombre"),
                "area": r.get("area"),
                "estado": estado,
                "dias": r.get("dias")
            },
            "color": color
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

    inicio = ddmmyyyy_to_date(inicio_raw) if "/" in inicio_raw else date.fromisoformat(inicio_raw)
    if inicio is None:
        return jsonify({"permitido": False, "mensaje": "Fecha inválida"})

    # Bloquear enero
    if inicio.month == 1:
        return jsonify({"permitido": False, "mensaje": "No se pueden solicitar vacaciones con inicio en enero."})

    dias_solicitados = 15
    fin = calcular_fin(inicio, dias_habiles=dias_solicitados)

    employees = load_employees()
    empleado = employees.get(nombre)
    if not empleado:
        return jsonify({"permitido": False, "mensaje": "Empleado no encontrado"})

    # Validar turno por antigüedad dentro del área
    reqs = load_requests()
    ok, msg = has_previous_employees_requested(nombre, employees, reqs)
    if not ok:
        return jsonify({"permitido": False, "mensaje": msg})

    # Revisar solapamientos por área con solicitudes Pending y Approved
    area = empleado.get("area", "General")
    for r in reqs:
        if r.get("area") == area and (r.get("estado") in ("Approved", "Aprobado")):
            r_inicio = date.fromisoformat(r.get("inicio"))
            r_fin = date.fromisoformat(r.get("fin"))
            if rango_solapa(inicio, fin, r_inicio, r_fin):
                return jsonify({"permitido": False, "mensaje": f"Ya existe una persona del área {area} con solicitud aprobada en ese periodo."})

    # Validar fecha_15
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
    inicio_field = request.form["inicio_fecha"]
    inicio = ddmmyyyy_to_date(inicio_field) if "/" in inicio_field else date.fromisoformat(inicio_field)

    employees = load_employees()
    empleado = employees.get(nombre)

    if not empleado:
        return redirect(url_for("employee_request_form", message="Empleado no encontrado"))

    if inicio.month == 1:
        return redirect(url_for("employee_request_form", message="No se pueden solicitar vacaciones con inicio en enero."))

    dias_solicitados = 15
    fin = calcular_fin(inicio, dias_habiles=dias_solicitados)

    # Validar turno por antigüedad
    reqs = load_requests()
    ok, msg = has_previous_employees_requested(nombre, employees, reqs)
    if not ok:
        return redirect(url_for("employee_request_form", message=msg))

    # verificar solapamiento por area contra Pending y Approved
    area = empleado.get("area", "General")
    for r in reqs:
        if r.get("area") == area and (r.get("estado") in ("Approved", "Aprobado")):
            r_inicio = date.fromisoformat(r.get("inicio"))
            r_fin = date.fromisoformat(r.get("fin"))
            if rango_solapa(inicio, fin, r_inicio, r_fin):
                return redirect(url_for("employee_request_form", message=f"Error: otra persona del área {area} ya tiene solicitud aprobada en ese periodo."))

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

    # Si se quiere aprobar: verificar solapamientos por área con otras aprobadas/pending
    if action == "Aprobado":
        area = target.get("area", "General")
        inicio_dt = date.fromisoformat(target.get("inicio"))
        fin_dt = date.fromisoformat(target.get("fin"))

        # Validación de solapamiento (pero solo para avisar)
        for r in reqs:
            if r is target:
                continue
            if r.get("area") == area and (r.get("estado") in ("Approved", "Aprobado")):
                r_inicio = date.fromisoformat(r.get("inicio"))
                r_fin = date.fromisoformat(r.get("fin"))

                if rango_solapa(inicio_dt, fin_dt, r_inicio, r_fin):

                    # Si viene en la petición el flag "force", aprobamos igual
                    if request.form.get("force") == "1":
                        target["estado"] = "Approved"
                        save_requests(reqs)
                        return jsonify({"success": True, "mensaje": "Aprobada con advertencia"})

                    # Si NO hay force → devolvemos warning
                    return jsonify({
                        "success": False,
                        "warning": True,
                        "mensaje": f"Advertencia: solapa con {r.get('nombre')} del área {area}. ¿Deseas aprobar igualmente?"
                    })
        # Si no hay solapamiento → aprobar normal
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

if __name__ == "__main__":
    # Para producción en Render (usa $PORT). Cambia si prefieres gunicorn u otro WSGI.
        app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
