from flask import Flask, render_template, request, redirect, session, flash, send_file
import sqlite3
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = "taller_secret_key"

# -------------------
# USUARIO FIJO
# -------------------
USUARIO = "admin"
PASSWORD = "1234"


# -------------------
# CONEXIÓN BD
# -------------------
def conectar():
    return sqlite3.connect("taller.db")


# -------------------
# CREAR / ACTUALIZAR TABLA
# -------------------
def crear_tabla():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vehiculos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            placa TEXT,
            cliente TEXT,
            modelo TEXT,
            diagnostico TEXT,
            estado TEXT,
            monto REAL
        )
    """)

    conn.commit()

    # Revisar si tu tabla antigua no tiene alguna columna
    cursor.execute("PRAGMA table_info(vehiculos)")
    columnas_existentes = [columna[1] for columna in cursor.fetchall()]

    columnas_necesarias = {
        "fecha": "TEXT",
        "placa": "TEXT",
        "cliente": "TEXT",
        "modelo": "TEXT",
        "diagnostico": "TEXT",
        "estado": "TEXT",
        "monto": "REAL DEFAULT 0"
    }

    for columna, tipo in columnas_necesarias.items():
        if columna not in columnas_existentes:
            cursor.execute(f"ALTER TABLE vehiculos ADD COLUMN {columna} {tipo}")

    conn.commit()
    conn.close()


# -------------------
# LOGIN
# -------------------
@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":
        usuario = request.form.get("usuario")
        clave = request.form.get("clave")

        if usuario == USUARIO and clave == PASSWORD:
            session["logueado"] = True
            return redirect("/lista")
        else:
            flash("Usuario o contraseña incorrecta")
            return redirect("/login")

    return render_template("login.html")


# -------------------
# LOGOUT
# -------------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# -------------------
# INICIO
# -------------------
@app.route("/")
def inicio():

    if not session.get("logueado"):
        return redirect("/login")

    return redirect("/lista")


# -------------------
# REGISTRO DE DATOS
# -------------------
@app.route("/lista")
def lista():

    if not session.get("logueado"):
        return redirect("/login")

    return render_template("index.html")


# -------------------
# GUARDAR VEHÍCULO
# -------------------
@app.route("/guardar", methods=["POST"])
def guardar():

    if not session.get("logueado"):
        return redirect("/login")

    fecha = request.form.get("fecha", "")
    placa = request.form.get("placa", "")
    cliente = request.form.get("cliente", "")
    modelo = request.form.get("modelo", "")
    diagnostico = request.form.get("diagnostico", "")
    estado = request.form.get("estado", "")
    monto = request.form.get("monto", "0")

    try:
        monto = float(monto)
    except ValueError:
        monto = 0

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO vehiculos
        (fecha, placa, cliente, modelo, diagnostico, estado, monto)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (fecha, placa, cliente, modelo, diagnostico, estado, monto))

    conn.commit()
    conn.close()

    flash("Registro guardado correctamente")
    return redirect("/lista")


# -------------------
# REPORTE GENERAL
# -------------------
@app.route("/reporte_general")
def reporte_general():

    if not session.get("logueado"):
        return redirect("/login")

    fecha = request.args.get("fecha", "")
    mes = request.args.get("mes", "")
    cliente = request.args.get("cliente", "")

    conn = conectar()
    cursor = conn.cursor()

    query = """
        SELECT id, fecha, placa, cliente, modelo, diagnostico, estado, monto
        FROM vehiculos
        WHERE 1=1
    """

    params = []

    if fecha:
        query += " AND fecha = ?"
        params.append(fecha)

    if mes:
        query += " AND substr(fecha, 1, 7) = ?"
        params.append(mes)

    if cliente:
        query += " AND cliente LIKE ?"
        params.append("%" + cliente + "%")

    query += " ORDER BY fecha DESC, id DESC"

    cursor.execute(query, params)
    datos = cursor.fetchall()

    conn.close()

    return render_template(
        "reporte_general.html",
        datos=datos,
        fecha=fecha,
        mes=mes,
        cliente=cliente
    )


# -------------------
# EXPORTAR EXCEL
# -------------------
@app.route("/exportar_excel")
def exportar_excel():

    if not session.get("logueado"):
        return redirect("/login")

    fecha = request.args.get("fecha", "")
    mes = request.args.get("mes", "")
    cliente = request.args.get("cliente", "")

    conn = conectar()

    query = """
        SELECT fecha, placa, cliente, modelo, diagnostico, estado, monto
        FROM vehiculos
        WHERE 1=1
    """

    params = []

    if fecha:
        query += " AND fecha = ?"
        params.append(fecha)

    if mes:
        query += " AND substr(fecha, 1, 7) = ?"
        params.append(mes)

    if cliente:
        query += " AND cliente LIKE ?"
        params.append("%" + cliente + "%")

    query += " ORDER BY fecha DESC"

    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte General")

    salida.seek(0)

    return send_file(
        salida,
        as_attachment=True,
        download_name="Reporte_Taller.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# -------------------
# ESTADÍSTICA: INGRESO DE VEHÍCULOS
# -------------------
@app.route("/ingreso_vehiculos")
def ingreso_vehiculos():

    if not session.get("logueado"):
        return redirect("/login")

    filtro = request.args.get("filtro", "dia")

    conn = conectar()
    cursor = conn.cursor()

    if filtro == "semana":
        agrupacion = "strftime('%Y - Semana %W', fecha)"
    elif filtro == "mes":
        agrupacion = "strftime('%Y-%m', fecha)"
    else:
        agrupacion = "fecha"

    cursor.execute(f"""
        SELECT {agrupacion} AS periodo, COUNT(*)
        FROM vehiculos
        GROUP BY periodo
        ORDER BY periodo ASC
    """)

    datos = cursor.fetchall()
    conn.close()

    labels = [fila[0] for fila in datos]
    valores = [fila[1] for fila in datos]

    return render_template(
        "ingreso_vehiculos.html",
        filtro=filtro,
        labels=labels,
        valores=valores
    )


# -------------------
# ESTADÍSTICA: ESTADO DE VEHÍCULOS
# -------------------
@app.route("/estado_vehiculos")
def estado_vehiculos():

    if not session.get("logueado"):
        return redirect("/login")

    filtro = request.args.get("filtro", "dia")

    conn = conectar()
    cursor = conn.cursor()

    if filtro == "semana":
        agrupacion = "strftime('%Y - Semana %W', fecha)"
    elif filtro == "mes":
        agrupacion = "strftime('%Y-%m', fecha)"
    else:
        agrupacion = "fecha"

    cursor.execute(f"""
        SELECT {agrupacion} AS periodo, estado, COUNT(*)
        FROM vehiculos
        GROUP BY periodo, estado
        ORDER BY periodo ASC
    """)

    datos_estado = cursor.fetchall()

    cursor.execute("""
        SELECT estado, COUNT(*)
        FROM vehiculos
        GROUP BY estado
    """)

    datos_pastel = cursor.fetchall()
    conn.close()

    periodos = sorted(list(set([fila[0] for fila in datos_estado])))
    estados = sorted(list(set([fila[1] for fila in datos_estado])))

    datasets = []

    for estado in estados:
        valores = []

        for periodo in periodos:
            cantidad = 0

            for fila in datos_estado:
                if fila[0] == periodo and fila[1] == estado:
                    cantidad = fila[2]

            valores.append(cantidad)

        datasets.append({
            "label": estado,
            "data": valores
        })

    labels_pastel = [fila[0] for fila in datos_pastel]
    valores_pastel = [fila[1] for fila in datos_pastel]

    return render_template(
        "estado_vehiculos.html",
        filtro=filtro,
        periodos=periodos,
        datasets=datasets,
        labels_pastel=labels_pastel,
        valores_pastel=valores_pastel
    )


# -------------------
# EDITAR VEHÍCULO
# -------------------
@app.route("/editar/<int:id>", methods=["GET", "POST"])
def editar(id):

    if not session.get("logueado"):
        return redirect("/login")

    conn = conectar()
    cursor = conn.cursor()

    if request.method == "POST":

        fecha = request.form.get("fecha", "")
        placa = request.form.get("placa", "")
        cliente = request.form.get("cliente", "")
        modelo = request.form.get("modelo", "")
        diagnostico = request.form.get("diagnostico", "")
        estado = request.form.get("estado", "")
        monto = request.form.get("monto", "0")

        try:
            monto = float(monto)
        except ValueError:
            monto = 0

        cursor.execute("""
            UPDATE vehiculos
            SET fecha=?, placa=?, cliente=?, modelo=?, diagnostico=?, estado=?, monto=?
            WHERE id=?
        """, (fecha, placa, cliente, modelo, diagnostico, estado, monto, id))

        conn.commit()
        conn.close()

        flash("Registro actualizado correctamente")
        return redirect("/reporte_general")

    cursor.execute("""
        SELECT id, fecha, placa, cliente, modelo, diagnostico, estado, monto
        FROM vehiculos
        WHERE id=?
    """, (id,))

    registro = cursor.fetchone()
    conn.close()

    return render_template("editar.html", registro=registro)


# -------------------
# ELIMINAR VEHÍCULO
# -------------------
@app.route("/eliminar/<int:id>")
def eliminar(id):

    if not session.get("logueado"):
        return redirect("/login")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM vehiculos WHERE id=?", (id,))
    conn.commit()
    conn.close()

    flash("Registro eliminado correctamente")
    return redirect("/reporte_general")


# -------------------
# DASHBOARD ANTIGUO OPCIONAL
# -------------------
@app.route("/dashboard")
def dashboard():

    if not session.get("logueado"):
        return redirect("/login")

    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT estado, COUNT(*)
        FROM vehiculos
        GROUP BY estado
    """)
    datos_estado = cursor.fetchall()

    cursor.execute("""
        SELECT fecha, COUNT(*)
        FROM vehiculos
        GROUP BY fecha
        ORDER BY fecha ASC
    """)
    datos_dia = cursor.fetchall()

    cursor.execute("""
        SELECT SUM(monto)
        FROM vehiculos
    """)
    total_ingresos = cursor.fetchone()[0]

    if total_ingresos is None:
        total_ingresos = 0

    conn.close()

    return render_template(
        "dashboard.html",
        datos_estado=datos_estado,
        datos_dia=datos_dia,
        total_ingresos=total_ingresos
    )


# -------------------
# EJECUTAR APP
# -------------------
if __name__ == "__main__":
    crear_tabla()
    app.run(debug=True)