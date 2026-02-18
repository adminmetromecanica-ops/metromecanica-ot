"""
audit_logger.py - Sistema de Auditoría para INACAL
Registra todas las OTs generadas con trazabilidad completa
"""
import sqlite3
import json
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), 'audit_log.db')

def init_audit_db():
    """Inicializa la base de datos de auditoría"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT NOT NULL,
        ot_number TEXT NOT NULL UNIQUE,
        expediente TEXT NOT NULL,
        proforma_number TEXT NOT NULL,
        cliente TEXT NOT NULL,
        ruc_cliente TEXT NOT NULL,
        total_items INTEGER NOT NULL,
        tipo_servicio TEXT,
        fecha_emision TEXT,
        fecha_entrega TEXT,
        estado TEXT DEFAULT 'APROBADA',
        usuario TEXT,
        ip_address TEXT,
        filepath TEXT,
        metadata TEXT
    )
    ''')
    
    # Índices para búsquedas rápidas
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_ot_number ON audit_log(ot_number)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_proforma ON audit_log(proforma_number)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_fecha ON audit_log(timestamp)')
    
    conn.commit()
    conn.close()

def register_ot(ot_data, filepath=None):
    """
    Registra una OT generada en el log de auditoría
    
    Args:
        ot_data: Dict con los datos de la OT
        filepath: Ruta del archivo generado
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
        INSERT INTO audit_log (
            timestamp, ot_number, expediente, proforma_number,
            cliente, ruc_cliente, total_items, tipo_servicio,
            fecha_emision, fecha_entrega, estado, filepath, metadata
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            datetime.now().isoformat(),
            ot_data.get('ot_number', ''),
            ot_data.get('expediente', ''),
            ot_data.get('numero_proforma', ''),
            ot_data.get('cliente', ''),
            ot_data.get('ruc_cliente', ''),
            ot_data.get('total_items', 0),
            ot_data.get('tipo_servicio', ''),
            ot_data.get('fecha_emision', ''),
            ot_data.get('plazo_entrega', ''),
            'APROBADA',
            filepath,
            json.dumps(ot_data, ensure_ascii=False)
        ))
        
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        # OT ya existe - esto está bien, significa que el número es único
        return False
    finally:
        conn.close()

def get_audit_log(start_date=None, end_date=None, cliente=None):
    """
    Obtiene registros del log de auditoría
    
    Args:
        start_date: Fecha inicio (YYYY-MM-DD)
        end_date: Fecha fin (YYYY-MM-DD)
        cliente: Filtrar por cliente
    
    Returns:
        Lista de registros
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = 'SELECT * FROM audit_log WHERE 1=1'
    params = []
    
    if start_date:
        query += ' AND DATE(timestamp) >= ?'
        params.append(start_date)
    
    if end_date:
        query += ' AND DATE(timestamp) <= ?'
        params.append(end_date)
    
    if cliente:
        query += ' AND cliente LIKE ?'
        params.append(f'%{cliente}%')
    
    query += ' ORDER BY timestamp DESC'
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    
    return [dict(row) for row in rows]

def export_audit_csv(output_path, start_date=None, end_date=None):
    """
    Exporta el log de auditoría a CSV para revisión INACAL
    
    Args:
        output_path: Ruta donde guardar el CSV
        start_date: Fecha inicio filtro
        end_date: Fecha fin filtro
    """
    import csv
    
    records = get_audit_log(start_date, end_date)
    
    if not records:
        return False
    
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        fieldnames = [
            'timestamp', 'ot_number', 'expediente', 'proforma_number',
            'cliente', 'ruc_cliente', 'total_items', 'tipo_servicio',
            'fecha_emision', 'fecha_entrega', 'estado'
        ]
        
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        for record in records:
            writer.writerow({k: record.get(k, '') for k in fieldnames})
    
    return True

def get_statistics():
    """Obtiene estadísticas para reportes de auditoría"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    stats = {}
    
    # Total de OTs generadas
    cursor.execute('SELECT COUNT(*) FROM audit_log')
    stats['total_ots'] = cursor.fetchone()[0]
    
    # OTs por mes
    cursor.execute('''
    SELECT strftime('%Y-%m', timestamp) as mes, COUNT(*) as cantidad
    FROM audit_log
    GROUP BY mes
    ORDER BY mes DESC
    LIMIT 12
    ''')
    stats['por_mes'] = cursor.fetchall()
    
    # Clientes únicos
    cursor.execute('SELECT COUNT(DISTINCT ruc_cliente) FROM audit_log')
    stats['clientes_unicos'] = cursor.fetchone()[0]
    
    # Tipos de servicio
    cursor.execute('''
    SELECT tipo_servicio, COUNT(*) as cantidad
    FROM audit_log
    GROUP BY tipo_servicio
    ''')
    stats['por_tipo'] = cursor.fetchall()
    
    conn.close()
    return stats

# Inicializar DB al importar el módulo
init_audit_db()
