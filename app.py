import os
from datetime import date, datetime, timedelta
import pandas as pd
from flask import Flask, render_template, request, send_from_directory, jsonify
import holidays
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import json

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PDF_FOLDER'] = 'pdf'

festivos = holidays.Colombia()


# -------------------- UTIL --------------------
def parse_blocked_dates(text):
    """
    Text examples:
      "2025-12-24,2025-12-25"
      "2025-12-24:2025-12-26,2026-01-01"
    Returns a set of date objects.
    """
    if not text:
        return set()
    text = text.strip()
    parts = [p.strip() for p in text.split(",") if p.strip()]
    blocked = set()
    for p in parts:
        if ":" in p:
            a, b = p.split(":", 1)
            start = datetime.strptime(a.strip(), "%Y-%m-%d").date()
            end = datetime.strptime(b.strip(), "%Y-%m-%d").date()
            cur = start
            while cur <= end:
                blocked.add(cur)
                cur += timedelta(days=1)
        else:
            d = datetime.strptime(p, "%Y-%m-%d").date()
            blocked.add(d)
    return blocked


def is_business_day(d):
    # d is datetime.date
    if d.weekday() >= 5:
        return False
    if d in festivos:
        return False
    return True


def next_business_day(d):
    d2 = d
    while not is_business_day(d2):
        d2 += timedelta(days=1)
    return d2


def sumar_dias_habiles_colombia(fecha_inicio, dias, blocked_dates=set()):
    """
    Devuelve la lista de 'dias' fechas hábiles (date objects) empezando desde 'fecha_inicio' (inclusive).
    Si no se pueden obtener (por ejemplo bloqueos) devuelve lista corta (len < dias).
    """
    if dias <= 0:
        return []
    dias_obtenidos = []
    cur = fecha_inicio
    safeguard = 0
    while len(dias_obtenidos) < dias:
        safeguard += 1
        if safeguard > 10000:
            break
        if cur.weekday() < 5 and cur not in festivos and cur not in blocked_dates:
            dias_obtenidos.append(cur)
        cur += timedelta(days=1)
    return dias_obtenidos


# -------------------- LÓGICA DE VACACIONES --------------------
def calcular_dias_vacaciones(fecha_ingreso):
    hoy = date.today()
    años = (hoy - fecha_ingreso).days // 365
    return años * 15


def asignar_calendario_con_restricciones(df, start_date=None, blocked_dates=set(), forced_starts=None):
    """
    Versión corregida: Permite paralelismo entre áreas y busca siempre la fecha más pronta.
    """
    if forced_starts is None:
        forced_starts = {}

    df = df.copy()

    # --- Preprocesamiento igual al anterior ---
    if "area" not in df.columns:
        df["area"] = "General"
    if "dias_tomados" not in df.columns:
        df["dias_tomados"] = 0

    df["fecha_ingreso"] = pd.to_datetime(df["fecha_ingreso"]).dt.date

    hoy = date.today()
    df["dias_trabajados"] = df["fecha_ingreso"].apply(lambda f: (hoy - f).days)
    df["dias_generados"] = df["fecha_ingreso"].apply(calcular_dias_vacaciones)
    df["dias_tomados"] = df["dias_tomados"].fillna(0).astype(int)
    df["dias_disponibles"] = df["dias_generados"] - df["dias_tomados"]
    df["dias_disponibles"] = df["dias_disponibles"].apply(lambda x: int(max(0, x)))

    # Filtros y ordenamiento
    df = df[df["dias_disponibles"] >= 15].reset_index(drop=True)
    df = df.sort_values(by="dias_disponibles").reset_index(drop=True)

    # Control de ocupación por área
    ocupados_por_area = {}  # area -> set(date)

    # Definir la FECHA BASE de inicio para TODOS
    # (Ya no usamos un cursor que avanza globalmente)
    fecha_base = start_date if start_date else date.today()
    fecha_base = next_business_day(fecha_base)

    lista_inicio = []
    lista_fin = []

    for _, row in df.iterrows():
        nombre = row["nombre"]
        area = row["area"]
        dias = int(row["dias_disponibles"])

        # Inicializar el set del área si no existe
        if area not in ocupados_por_area:
            ocupados_por_area[area] = set()

        if dias <= 0:
            lista_inicio.append(None)
            lista_fin.append(None)
            continue

        # 1. Determinar desde dónde empezar a buscar para ESTE empleado
        # Si tiene fecha forzada, usamos esa. Si no, usamos la fecha_base global.
        forced = None
        if str(nombre) in forced_starts:
            try:
                forced = forced_starts[str(nombre)]
                forced = next_business_day(forced)
                candidato = forced
            except:
                candidato = fecha_base
        else:
            candidato = fecha_base

        # 2. Buscar espacio libre
        found = False
        attempts = 0
        
        while not found:
            attempts += 1
            if attempts > 3000: # Evitar bucles infinitos
                lista_inicio.append(None)
                lista_fin.append(None)
                break

            # Ajustar candidato si cae en festivo o bloqueo global inmediato
            candidato = next_business_day(candidato)
            if candidato in blocked_dates:
                candidato += timedelta(days=1)
                continue

            # Generar rango de días necesarios
            dias_lista = sumar_dias_habiles_colombia(candidato, dias, blocked_dates)
            
            # Si no logramos conseguir los días (ej. fin de calendario), avanzamos
            if len(dias_lista) < dias:
                candidato += timedelta(days=1)
                continue

            # VERIFICACIÓN DE CONFLICTOS

            # A) Conflicto GLOBAL (Blocked dates en medio del rango)
            # Aunque sumar_dias salta festivos, verificamos si 'blocked_dates' (ej. cierre de planta) rompe la continuidad deseada o se solapa.
            if any(d in blocked_dates for d in dias_lista):
                # Si choca con bloqueo, movemos candidato después del bloqueo
                max_blocked = max([d for d in dias_lista if d in blocked_dates])
                candidato = max_blocked + timedelta(days=1)
                # Si era forzado, falló la fuerza, volvemos a buscar desde fecha_base
                if forced: 
                    forced = None
                    candidato = fecha_base
                continue

            # B) Conflicto DE ÁREA (Solo importa si chocan con compañeros de SU MISMA área)
            ocupados = ocupados_por_area[area]
            if any(d in ocupados for d in dias_lista):
                # Hay choque en su área: mover candidato después del último ocupado encontrado
                ultimo_ocupado = max([d for d in dias_lista if d in ocupados])
                candidato = ultimo_ocupado + timedelta(days=1)
                
                if forced:
                    forced = None
                    candidato = fecha_base
                continue

            # ¡ENCONTRADO!
            inicio = dias_lista[0]
            fin = dias_lista[-1]

            # Registrar ocupación en SU área
            for d in dias_lista:
                ocupados_por_area[area].add(d)

            lista_inicio.append(inicio)
            lista_fin.append(fin)
            found = True
            
            # IMPORTANTE: NO actualizamos ninguna variable global 'cursor' aquí.
            # El siguiente empleado volverá a empezar desde 'fecha_base'.

    df["inicio_vacaciones"] = lista_inicio
    df["fin_vacaciones"] = lista_fin

    return df


# -------------------- RUTAS --------------------
@app.route("/")
def index():
    # index.html deberá incluir inputs opcionales:
    # start_date (YYYY-MM-DD), blocked_dates (texto), area_filter (opcional para mostrar)
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    archivo = request.files.get("archivo")
    if not archivo:
        return "No subiste archivo", 400

    # leer parámetros opcionales del formulario (aceptamos ambos nombres: fecha_inicio o start_date)
    start_date_txt = request.form.get("start_date", "").strip() or request.form.get("fecha_inicio", "").strip()
    blocked_txt = request.form.get("blocked_dates", "").strip()
    area_filter = request.form.get("area_filter", "").strip()  # opcional para filtrar vista

    # parsear start_date
    start_date = None
    if start_date_txt:
        try:
            start_date = datetime.strptime(start_date_txt, "%Y-%m-%d").date()
        except Exception as e:
            return f"Formato de start_date inválido, debe ser YYYY-MM-DD: {e}", 400

    blocked_dates = parse_blocked_dates(blocked_txt)

    # asegurar folders
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["PDF_FOLDER"], exist_ok=True)

    ruta = os.path.join(app.config["UPLOAD_FOLDER"], archivo.filename)
    archivo.save(ruta)

    # También guardamos una copia conocida como 'ultimo.xlsx' o 'ultimo.csv' para edición posterior
    filename_lower = archivo.filename.lower()
    if filename_lower.endswith(".csv"):
        ultima_ruta = os.path.join(app.config["UPLOAD_FOLDER"], "ultimo.csv")
        archivo.stream.seek(0)
        with open(ultima_ruta, "wb") as f:
            f.write(archivo.read())
    else:
        # intentar excel
        ultima_ruta = os.path.join(app.config["UPLOAD_FOLDER"], "ultimo.xlsx")
        archivo.stream.seek(0)
        with open(ultima_ruta, "wb") as f:
            f.write(archivo.read())

    # leer excel/csv
    try:
        if archivo.filename.lower().endswith(".csv"):
            df = pd.read_csv(ruta)
        else:
            df = pd.read_excel(ruta)
    except Exception as e:
        return f"Error leyendo el archivo: {e}", 500

    # asignar calendario teniendo en cuenta restricciones
    df_final = asignar_calendario_con_restricciones(df, start_date=start_date, blocked_dates=blocked_dates)

    # generar pdf
    ruta_pdf = os.path.join(app.config["PDF_FOLDER"], "vacaciones.pdf")
    generar_pdf(df_final, ruta_pdf)

    # preparar resultados y eventos para la vista
    resultados = df_final.to_dict(orient="records")

    # aplicar filtro por area para la vista si se solicitó
    if area_filter:
        resultados_vista = [r for r in resultados if str(r.get("area", "")).lower() == area_filter.lower()]
    else:
        resultados_vista = resultados

    eventos = []
    for r in resultados:
        if r.get("inicio_vacaciones") and r.get("fin_vacaciones"):
            # asegurar que las fechas sean strings YYYY-MM-DD
            start_str = str(r["inicio_vacaciones"])
            end_obj = r["fin_vacaciones"]
            # if end_obj is Timestamp etc, convert to date
            if hasattr(end_obj, "date"):
                end_str = str(end_obj.date())
                end_dt = end_obj.date()
            else:
                end_str = str(end_obj)
                try:
                    end_dt = datetime.strptime(end_str, "%Y-%m-%d").date()
                except Exception:
                    # fallback
                    end_dt = date.today()
            eventos.append({
                "title": f"{r['nombre']} ({r.get('area','')})",
                "start": start_str,
                # FullCalendar expects end exclusive -> add 1 day
                "end": str(end_dt + timedelta(days=1)),
                "color": "#2563eb"
            })

    # pasar también blocked dates y start_date a la plantilla para mostrar info
    return render_template(
        "resultados.html",
        resultados=resultados_vista,
        eventos=json.dumps(eventos),
        blocked_dates=[d.isoformat() for d in sorted(blocked_dates)],
        start_date=start_date_txt or date.today().isoformat()
    )


@app.route("/editar_fecha", methods=["POST"])
def editar_fecha():
    """
    Recibe form-data: nombre, nueva_fecha (YYYY-MM-DD)
    Lee el último archivo subido (ultimo.xlsx o ultimo.csv), fuerza la fecha para el empleado
    y re-calcula el calendario con forced_starts.
    """
    nombre = request.form.get("nombre")
    nueva_fecha_txt = request.form.get("nueva_fecha")
    if not nombre or not nueva_fecha_txt:
        return jsonify({"ok": False, "msg": "Faltan parámetros 'nombre' o 'nueva_fecha'"}), 400
    try:
        nueva_fecha = datetime.strptime(nueva_fecha_txt, "%Y-%m-%d").date()
    except Exception:
        return jsonify({"ok": False, "msg": "Formato de fecha inválido (YYYY-MM-DD)"}), 400

    # localizar ultimo archivo
    ultima_xlsx = os.path.join(app.config["UPLOAD_FOLDER"], "ultimo.xlsx")
    ultima_csv = os.path.join(app.config["UPLOAD_FOLDER"], "ultimo.csv")
    if os.path.exists(ultima_xlsx):
        df = pd.read_excel(ultima_xlsx)
    elif os.path.exists(ultima_csv):
        df = pd.read_csv(ultima_csv)
    else:
        return jsonify({"ok": False, "msg": "No existe archivo previo (sube primero uno)"}), 400

    # construir forced_starts dict
    forced = {str(nombre): nueva_fecha}

    # Leer blocked dates from form? allow optional blocked_dates param to be passed
    blocked_txt = request.form.get("blocked_dates", "")
    blocked_dates = parse_blocked_dates(blocked_txt)

    # re-run assignment with forced start
    df_final = asignar_calendario_con_restricciones(df, start_date=None, blocked_dates=blocked_dates, forced_starts=forced)

    # generar pdf actualizado
    ruta_pdf = os.path.join(app.config["PDF_FOLDER"], "vacaciones.pdf")
    generar_pdf(df_final, ruta_pdf)

    # opcional: podríamos devolver la tabla actualizada
    resultados = df_final.to_dict(orient="records")
    return jsonify({"ok": True, "resultados": resultados})


def generar_pdf(df, ruta_pdf):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(ruta_pdf)

    story = []
    story.append(Paragraph("<b>Calendario de Vacaciones</b>", styles["Title"]))
    story.append(Spacer(1, 20))

    for _, row in df.iterrows():
        nombre = row.get("nombre")
        inicio = row.get("inicio_vacaciones")
        fin = row.get("fin_vacaciones")
        area = row.get("area", "General")
        if inicio and fin:
            texto = f"<b>{nombre}</b> ({area}) — Inicio: {inicio} — Fin: {fin}"
        else:
            texto = f"<b>{nombre}</b> ({area}) — No tiene días disponibles"
        story.append(Paragraph(texto, styles["Normal"]))
        story.append(Spacer(1, 10))

    doc.build(story)


@app.route("/descargar_pdf")
def descargar_pdf():
    ruta_pdf = os.path.join(app.config["PDF_FOLDER"], "vacaciones.pdf")
    if not os.path.exists(ruta_pdf):
        return "PDF no encontrado. Genera primero el calendario.", 404
    return send_from_directory(app.config["PDF_FOLDER"], "vacaciones.pdf", as_attachment=True)


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
