import sqlite3

conn = sqlite3.connect("taller.db")
cursor = conn.cursor()

cursor.execute("DROP TABLE IF EXISTS vehiculos")

cursor.execute("""
CREATE TABLE vehiculos(
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
conn.close()

print("Base de datos creada correctamente")