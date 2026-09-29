from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify, flash
from openpyxl import load_workbook
from pathlib import Path
from shutil import copy2
from datetime import datetime
import os, sys, tempfile

APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
BASE_DATA = Path(os.environ.get("JAC_DATA_DIR", Path.home() / "Documents" / "JAC Nuevo Achi"))
DATA_DIR = BASE_DATA / "datos"
BACKUP_DIR = BASE_DATA / "respaldos"
DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

DATA_FILE = DATA_DIR / "Formato_Caracterizacion_Barrio_Nuevo_Achi.xlsm"
TEMPLATE_FILE = APP_DIR / "data" / "Formato_Caracterizacion_Barrio_Nuevo_Achi.xlsm"

if not DATA_FILE.exists():
    copy2(TEMPLATE_FILE, DATA_FILE)

app = Flask(__name__)
app.secret_key = os.environ.get("JAC_SECRET_KEY", "JAC-Nuevo-Achi-2026-Cambiar")

HEADERS = [
    "Código hogar","Fecha encuesta","Manzana","Bloque","Casa","Nombre jefe del hogar",
    "Tipo de documento jefe","Número documento jefe","Genero","Teléfono","Número de integrantes",
    "Registro de integrantes (ver hoja detallada)","Composición familiar","Edad jefe hogar",
    "Situación laboral","Ingresos mensuales aproximados","Condiciones económicas / observaciones",
    "Tipo de vivienda","Material predominante","Número de habitaciones","Servicios públicos disponibles",
    "Condición de ocupación","Dirección / ubicación del predio","Información predial / matrícula / referencia",
    "Área aproximada del predio (m²)","Uso del predio","Necesidades principales del hogar",
    "Observaciones de la visita","Documentación verificada","Registro fotográfico / código",
    "Estado de caracterización","Fecha de validación comunitaria"
]
MEMBER_HEADERS = [
    "Código hogar","N.º integrante","Nombre completo","Tipo de documento","Número de documento",
    "Parentesco con el jefe del hogar","Edad (años)","Sexo","Situación laboral",
    "Ingreso mensual aproximado ($)","Observaciones"
]
CONTROL_HEADERS = [
    "Código hogar","Jefe del hogar","Encuesta","Visita domiciliaria","Documentos",
    "Predio","Fotografías","Validación comunitaria","Observaciones"
]

def authorized():
    return session.get("auth") is True

def open_book():
    return load_workbook(DATA_FILE, keep_vba=True)

def find_row(ws, code):
    target = str(code).strip().upper()
    for r in range(2, ws.max_row + 1):
        value = ws.cell(r, 1).value
        if value is not None and str(value).strip().upper() == target:
            return r
    return None

def make_backup():
    if DATA_FILE.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        copy2(DATA_FILE, BACKUP_DIR / f"Caracterizacion_{stamp}.xlsm")

def safe_save(wb):
    make_backup()
    fd, tmp = tempfile.mkstemp(prefix="jach_", suffix=".xlsm", dir=DATA_DIR)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, DATA_FILE)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass

def get_households():
    wb = open_book()
    ws = wb["Caracterización"]
    result = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 33)]
        if any(v not in (None, "") for v in vals):
            result.append(dict(zip(HEADERS, vals)))
    return result

@app.route("/", methods=["GET", "POST"])
def login():
    if authorized():
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        user = request.form.get("username", "")
        password = request.form.get("password", "")
        expected_user = os.environ.get("JAC_USER", "admin")
        expected_password = os.environ.get("JAC_PASSWORD", "admin123")
        if user == expected_user and password == expected_password:
            session["auth"] = True
            return redirect(url_for("dashboard"))
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("login.html")

@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.get("/dashboard")
def dashboard():
    if not authorized():
        return redirect(url_for("login"))
    data = get_households()
    complete = sum(
        str(x.get("Estado de caracterización") or "").strip().lower()
        in {"completo", "validado", "validada"}
        for x in data
    )
    return render_template("dashboard.html", total=len(data), complete=complete,
                           pending=len(data)-complete)

@app.get("/api/hogares")
def api_households():
    if not authorized():
        return jsonify({"error": "No autorizado"}), 401
    q = request.args.get("q", "").strip().lower()
    data = get_households()
    if q:
        fields = ["Código hogar", "Nombre jefe del hogar",
                  "Número documento jefe", "Teléfono"]
        data = [x for x in data if any(q in str(x.get(f) or "").lower()
                                       for f in fields)]
    return jsonify(data)

@app.route("/hogar/<code>", methods=["GET", "POST"])
def household(code):
    if not authorized():
        return redirect(url_for("login"))
    wb = open_book()
    ws = wb["Caracterización"]
    row = find_row(ws, code)
    if not row:
        flash("No se encontró el hogar.", "error")
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        for col, header in enumerate(HEADERS, 1):
            if header == "Código hogar":
                continue
            ws.cell(row, col).value = request.form.get(f"f{col}", "").strip() or None
        safe_save(wb)
        flash(f"Hogar {code} actualizado correctamente.", "ok")
        return redirect(url_for("household", code=code))

    values = {h: ws.cell(row, i).value for i, h in enumerate(HEADERS, 1)}

    members = []
    mw = wb["Integrantes del hogar"]
    for r in range(2, mw.max_row + 1):
        if str(mw.cell(r, 1).value or "").strip().upper() == str(code).upper():
            member_data = {h: mw.cell(r, i).value
                            for i, h in enumerate(MEMBER_HEADERS, 1)}
            member_data["_row"] = r
            members.append(member_data)

    control = {}
    cw = wb["Control"]
    cr = find_row(cw, code)
    if cr:
        control = {h: cw.cell(cr, i).value
                   for i, h in enumerate(CONTROL_HEADERS, 1)}

    return render_template("hogar.html", vals=values, headers=HEADERS,
                           members=members, control=control, code=code)

@app.post("/hogar/<code>/integrante")
def add_member(code):
    if not authorized():
        return redirect(url_for("login"))
    wb = open_book()
    ws = wb["Integrantes del hogar"]
    row = ws.max_row + 1
    vals = [
        code, request.form.get("n", ""), request.form.get("nombre", ""),
        request.form.get("tipo", ""), request.form.get("doc", ""),
        request.form.get("parentesco", ""), request.form.get("edad", ""),
        request.form.get("sexo", ""), request.form.get("laboral", ""),
        request.form.get("ingreso", ""), request.form.get("obs", "")
    ]
    for i, value in enumerate(vals, 1):
        ws.cell(row, i).value = value or None
    safe_save(wb)
    flash("Integrante agregado.", "ok")
    return redirect(url_for("household", code=code))


@app.route("/hogar/<code>/integrante/<int:member_row>/editar", methods=["POST"])
def edit_member(code, member_row):
    if not authorized():
        return redirect(url_for("login"))
    wb = open_book()
    ws = wb["Integrantes del hogar"]
    if member_row < 2 or member_row > ws.max_row:
        flash("No se encontró el integrante.", "error")
        return redirect(url_for("household", code=code))
    owner = str(ws.cell(member_row, 1).value or "").strip().upper()
    if owner != str(code).strip().upper():
        flash("El integrante no pertenece a este hogar.", "error")
        return redirect(url_for("household", code=code))
    vals = [
        code, request.form.get("n", ""), request.form.get("nombre", ""),
        request.form.get("tipo", ""), request.form.get("doc", ""),
        request.form.get("parentesco", ""), request.form.get("edad", ""),
        request.form.get("sexo", ""), request.form.get("laboral", ""),
        request.form.get("ingreso", ""), request.form.get("obs", "")
    ]
    for i, value in enumerate(vals, 1):
        ws.cell(member_row, i).value = value or None
    safe_save(wb)
    flash("Integrante actualizado correctamente.", "ok")
    return redirect(url_for("household", code=code))

@app.post("/hogar/<code>/integrante/<int:member_row>/eliminar")
def delete_member(code, member_row):
    if not authorized():
        return redirect(url_for("login"))
    wb = open_book()
    ws = wb["Integrantes del hogar"]
    if member_row < 2 or member_row > ws.max_row:
        flash("No se encontró el integrante.", "error")
        return redirect(url_for("household", code=code))
    owner = str(ws.cell(member_row, 1).value or "").strip().upper()
    if owner != str(code).strip().upper():
        flash("El integrante no pertenece a este hogar.", "error")
        return redirect(url_for("household", code=code))
    ws.delete_rows(member_row, 1)
    safe_save(wb)
    flash("Integrante eliminado correctamente.", "ok")
    return redirect(url_for("household", code=code))

@app.post("/hogar/<code>/control")
def save_control(code):
    if not authorized():
        return redirect(url_for("login"))
    wb = open_book()
    ws = wb["Control"]
    row = find_row(ws, code) or (ws.max_row + 1)
    vals = [code, request.form.get("jefe", "")]
    vals += [request.form.get(k, "") for k in
             ["encuesta","visita","documentos","predio","fotos",
              "validacion","observaciones"]]
    for i, value in enumerate(vals, 1):
        ws.cell(row, i).value = value or None
    safe_save(wb)
    flash("Control documental actualizado.", "ok")
    return redirect(url_for("household", code=code))

@app.get("/descargar")
def download():
    if not authorized():
        return redirect(url_for("login"))
    return send_file(DATA_FILE, as_attachment=True,
                     download_name="Caracterizacion_Barrio_Nuevo_Achi_actualizada.xlsm")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
