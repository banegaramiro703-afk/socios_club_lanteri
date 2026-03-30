from flask import Flask, render_template, request, redirect, url_for
import sqlite3
import pandas as pd
from io import BytesIO
from flask import send_file

app = Flask(__name__)

def conectar_db():
    conn = sqlite3.connect('club.db')
    conn.row_factory = sqlite3.Row
    return conn

# --- 1. RUTA PRINCIPAL ---
@app.route('/')
def index():
    conn = conectar_db()
    socios = conn.execute('SELECT * FROM socios ORDER BY id_socio DESC').fetchall()
    conn.close()
    return render_template('index.html', socios=socios)

# --- 2. AGREGAR SOCIO ---
@app.route('/agregar', methods=['POST'])
def agregar():
    conn = conectar_db()
    # Agregamos el telefono a la consulta de inserción
    conn.execute('INSERT INTO socios (dni, nombre, apellido, telefono) VALUES (?, ?, ?, ?)', 
                 (request.form['dni'], request.form['nombre'], request.form['apellido'], request.form.get('telefono', '')))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

# --- 3. REGISTRAR PAGO ---
@app.route('/registrar_pago', methods=['POST'])
def registrar_pago():
    id_socio, monto = request.form['id_socio'], float(request.form['monto'])
    mes_inicio, anio, tipo = int(request.form['mes']), int(request.form['anio']), request.form['tipo']

    conn = conectar_db()
    if tipo == 'Mensual':
        conn.execute('INSERT INTO pagos (id_socio, monto, mes, anio, tipo) VALUES (?, ?, ?, ?, ?)',
                     (id_socio, monto, mes_inicio, anio, tipo))
    else:
        monto_por_mes = monto / 12
        for i in range(12):
            mes_actual = (mes_inicio + i - 1) % 12 + 1
            anio_actual = anio + (mes_inicio + i - 1) // 12
            conn.execute('INSERT INTO pagos (id_socio, monto, mes, anio, tipo) VALUES (?, ?, ?, ?, ?)',
                         (id_socio, monto_por_mes, mes_actual, anio_actual, tipo))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

# --- 4. HISTORIAL ---
@app.route('/historial/<int:id_socio>')
def historial(id_socio):
    conn = conectar_db()
    socio = conn.execute('SELECT * FROM socios WHERE id_socio = ?', (id_socio,)).fetchone()
    pagos = conn.execute('SELECT * FROM pagos WHERE id_socio = ? ORDER BY anio DESC, mes DESC', (id_socio,)).fetchall()
    conn.close()
    return render_template('historial.html', socio=socio, pagos=pagos)

# --- 5. DEUDORES ---
@app.route('/deudores')
def deudores():
    mes = request.args.get('mes', 3, type=int) 
    anio = request.args.get('anio', 2026, type=int)
    
    conn = conectar_db()
    monto_cuota = conn.execute('SELECT cuota_mensual FROM configuracion LIMIT 1').fetchone()[0]
    
    # Traemos también el teléfono en la consulta
    query = '''
        SELECT id_socio, nombre, apellido, dni, telefono, meses_deuda, (meses_deuda * ?) AS deuda_total
        FROM (
            SELECT s.id_socio, s.nombre, s.apellido, s.dni, s.telefono,
                   (? - (SELECT COUNT(*) FROM pagos p WHERE p.id_socio = s.id_socio AND p.anio = ? AND p.mes <= ?)) AS meses_deuda
            FROM socios s
            WHERE s.estado = 'Activo'
        )
        WHERE meses_deuda > 0
    '''
    lista_deudores = conn.execute(query, (monto_cuota, mes, anio, mes)).fetchall()
    conn.close()
    
    return render_template('deudores.html', deudores=lista_deudores, mes=mes, anio=anio)

# --- 6. EXPORTAR EXCEL ---
@app.route('/exportar_deudores')
def exportar_deudores():
    mes, anio = request.args.get('mes', 3, type=int), request.args.get('anio', 2026, type=int)
    
    conn = conectar_db()
    monto_cuota = conn.execute('SELECT cuota_mensual FROM configuracion LIMIT 1').fetchone()[0]
    
    query = '''
        SELECT dni AS "DNI", nombre AS "Nombre", apellido AS "Apellido", telefono AS "Teléfono",
               meses_deuda AS "Meses Atrasados", (meses_deuda * ?) AS "Deuda Total ($)"
        FROM (
            SELECT s.dni, s.nombre, s.apellido, s.telefono,
                   (? - (SELECT COUNT(*) FROM pagos p WHERE p.id_socio = s.id_socio AND p.anio = ? AND p.mes <= ?)) AS meses_deuda
            FROM socios s WHERE s.estado = 'Activo'
        ) WHERE meses_deuda > 0
    '''
    df = pd.read_sql_query(query, conn, params=(monto_cuota, mes, anio, mes))
    conn.close()
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'Morosos_{mes}_{anio}')
    output.seek(0)
    return send_file(output, download_name=f'Deudores_{mes}_{anio}.xlsx', as_attachment=True)

# --- 7. CONFIGURACIÓN ---
@app.route('/configuracion')
def configuracion():
    conn = conectar_db()
    precio = conn.execute('SELECT cuota_mensual FROM configuracion LIMIT 1').fetchone()[0]
    conn.close()
    return render_template('configuracion.html', precio=precio)

@app.route('/guardar_configuracion', methods=['POST'])
def guardar_configuracion():
    conn = conectar_db()
    conn.execute('UPDATE configuracion SET cuota_mensual = ?', (float(request.form['precio']),))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

# --- 8. RESUMEN ---
@app.route('/resumen')
def resumen():
    anio_actual = 2026
    conn = conectar_db()
    
    total_socios = conn.execute('SELECT COUNT(*) FROM socios WHERE estado = "Activo"').fetchone()[0]
    recaudacion_anual = conn.execute('SELECT SUM(monto) FROM pagos WHERE anio = ?', (anio_actual,)).fetchone()[0] or 0
    
    pagos_meses = conn.execute('SELECT mes, SUM(monto) as total FROM pagos WHERE anio = ? GROUP BY mes', (anio_actual,)).fetchall()
    
    datos_lista = [0] * 12
    for p in pagos_meses:
        datos_lista[p['mes'] - 1] = float(p['total'])
        
    conn.close()
    return render_template('resumen.html', socios=total_socios, recaudacion=recaudacion_anual, anio=anio_actual, datos=datos_lista)

# --- 9. EDITAR DATOS DEL SOCIO ---
@app.route('/editar_socio', methods=['POST'])
def editar_socio():
    id_socio = request.form['id_socio']
    dni = request.form['dni']
    nombre = request.form['nombre']
    apellido = request.form['apellido']
    telefono = request.form.get('telefono', '') # Obtenemos el teléfono nuevo
    
    conn = conectar_db()
    # Actualizamos el teléfono en la base de datos
    conn.execute('UPDATE socios SET dni = ?, nombre = ?, apellido = ?, telefono = ? WHERE id_socio = ?', 
                 (dni, nombre, apellido, telefono, id_socio))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

# --- 10. DAR DE BAJA / ALTA ---
@app.route('/cambiar_estado/<int:id_socio>')
def cambiar_estado(id_socio):
    conn = conectar_db()
    socio = conn.execute('SELECT estado FROM socios WHERE id_socio = ?', (id_socio,)).fetchone()
    nuevo_estado = 'Inactivo' if socio['estado'] == 'Activo' else 'Activo'
    
    conn.execute('UPDATE socios SET estado = ? WHERE id_socio = ?', (nuevo_estado, id_socio))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

if __name__ == '__main__':
    conn = conectar_db()
    
    # Creamos las tablas si no existen (agregando la columna telefono)
    conn.execute('''CREATE TABLE IF NOT EXISTS socios (id_socio INTEGER PRIMARY KEY AUTOINCREMENT, dni TEXT UNIQUE, nombre TEXT, apellido TEXT, telefono TEXT DEFAULT '', estado TEXT DEFAULT 'Activo')''')
    conn.execute('''CREATE TABLE IF NOT EXISTS pagos (id_pago INTEGER PRIMARY KEY AUTOINCREMENT, id_socio INTEGER, monto REAL, mes INTEGER, anio INTEGER, tipo TEXT, fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(id_socio) REFERENCES socios(id_socio))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS configuracion (id INTEGER PRIMARY KEY, cuota_mensual REAL)''')
    
    # Auto-Actualizador: Intenta agregar la columna telefono por si la base de datos ya existía de antes
    try:
        conn.execute("ALTER TABLE socios ADD COLUMN telefono TEXT DEFAULT ''")
        conn.commit()
        print("🔧 Base de datos actualizada: Columna 'telefono' agregada con éxito.")
    except:
        pass # Si la columna ya existe, simplemente sigue de largo sin tirar error
        
    if conn.execute('SELECT COUNT(*) FROM configuracion').fetchone()[0] == 0:
        conn.execute('INSERT INTO configuracion (cuota_mensual) VALUES (1000)')
        conn.commit()
    conn.close()
    app.run(debug=True)