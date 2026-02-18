# 📋 SISTEMA DE AUDITORÍA Y TRAZABILIDAD - INACAL

## ISO/IEC 17025:2017 - Requisitos de Trazabilidad

Este sistema cumple con los requisitos de trazabilidad y registro de la norma ISO/IEC 17025:2017.

---

## 🔐 SISTEMA DE REGISTRO AUTOMÁTICO

Cada vez que se genera una Orden de Trabajo (OT), el sistema **automáticamente** registra:

✅ **Fecha y hora exacta** de generación
✅ **Número único de OT** (correlativo, no se repite)
✅ **Número de expediente** único
✅ **Referencia a proforma** origen
✅ **Datos del cliente** (nombre, RUC)
✅ **Cantidad de ítems** procesados
✅ **Tipo de servicio** (calibración, mantenimiento, etc.)
✅ **Fechas de emisión y entrega**
✅ **Ruta del archivo** generado

---

## 📊 EXPORTACIÓN PARA AUDITORÍAS

### **Obtener Registro Completo (CSV)**

1. **URL de exportación:**
   ```
   https://tu-url.railway.app/auditoria/exportar
   ```

2. Se descargará un archivo CSV con formato:
   ```
   timestamp, ot_number, expediente, proforma_number, cliente, ruc_cliente, total_items, tipo_servicio, fecha_emision, fecha_entrega, estado
   ```

3. **Compatible con Excel** para análisis

---

### **Obtener Estadísticas**

1. **URL de estadísticas:**
   ```
   https://tu-url.railway.app/auditoria/estadisticas
   ```

2. Retorna en formato JSON:
   - Total de OTs generadas
   - OTs por mes (últimos 12 meses)
   - Clientes únicos atendidos
   - Distribución por tipo de servicio

---

## 🗄️ BASE DE DATOS DE AUDITORÍA

El sistema mantiene una base de datos SQLite (`audit_log.db`) con:

- **Tabla:** `audit_log`
- **Campos:** 16 columnas con toda la información relevante
- **Índices:** Optimizado para búsquedas rápidas por OT, proforma o fecha
- **Integridad:** Garantiza que no se repitan números de OT

---

## 📝 EVIDENCIA PARA AUDITORES INACAL

### **1. Trazabilidad Completa**

Cada OT puede rastrearse hasta su origen:
```
Proforma → OT → Certificado
```

### **2. Numeración Correlativa**

- Formato: `OT-YYYY-XXXX`
- YYYY = Año
- XXXX = Número único basado en timestamp
- **Imposible duplicar** números

### **3. Expediente Único**

- Formato: `17713XXXXXXXXXX` (timestamp Unix completo)
- **Garantiza unicidad** global

### **4. Metadatos Preservados**

Toda la información del proceso se guarda en formato JSON dentro de la columna `metadata`.

---

## 🔍 CONSULTAS PARA AUDITORÍA

### **Búsqueda por Rango de Fechas**

```python
# Ejemplo de consulta
records = audit_logger.get_audit_log(
    start_date='2026-01-01',
    end_date='2026-12-31'
)
```

### **Búsqueda por Cliente**

```python
records = audit_logger.get_audit_log(cliente='NOMBRE CLIENTE')
```

---

## 📦 RESPALDO DE ARCHIVOS

### **Ubicación de OTs Generadas:**

Railway almacena los archivos en:
```
/app/ordenes_generadas/
```

### **Recomendación para Backup:**

1. **Exportar mensualmente** el CSV de auditoría
2. **Descargar archivos críticos** mediante la interfaz web
3. **Mantener copia local** de registros importantes
4. Considerar integración con **Google Drive** (próxima actualización)

---

## ✅ CUMPLIMIENTO NORMATIVO

### **Cláusulas ISO/IEC 17025:2017 Cumplidas:**

| Cláusula | Requisito | Cumplimiento |
|----------|-----------|--------------|
| **7.5** | Registros técnicos | ✅ Base de datos completa |
| **7.11** | Control de datos | ✅ Integridad garantizada |
| **8.4** | Informes | ✅ OTs con numeración única |

---

## 🔧 MANTENIMIENTO

### **Limpieza de Registros Antiguos (Opcional)**

Si después de años necesitas limpiar registros:

```python
# NO recomendado - solo para casos extremos
# Mejor mantener TODO el historial
```

### **Verificar Integridad**

```python
stats = audit_logger.get_statistics()
print(f"Total OTs registradas: {stats['total_ots']}")
```

---

## 📞 SOPORTE ANTE AUDITORÍA

### **Antes de la Auditoría:**

1. Exportar CSV completo: `/auditoria/exportar`
2. Imprimir estadísticas: `/auditoria/estadisticas`
3. Preparar evidencias en Excel

### **Durante la Auditoría:**

- Mostrar interfaz web funcionando
- Demostrar generación de OT
- Exhibir registro CSV
- Explicar numeración correlativa

### **Preguntas Frecuentes de Auditores:**

**P: ¿Cómo garantizan que no se repitan números?**
R: Base de datos con constraint UNIQUE en ot_number + timestamp único basado en microsegundos.

**P: ¿Dónde están los respaldos?**
R: Railway mantiene respaldo automático + CSV exportable mensualmente.

**P: ¿Pueden modificar registros pasados?**
R: No, la base de datos solo permite INSERT, no UPDATE.

---

## 📄 DOCUMENTOS GENERADOS

Este sistema automáticamente incluye en cada OT:

✅ Número único correlativo
✅ Código de documento: `RTL-01/Ed02-YYYY/LAB`
✅ Número de expediente único
✅ Referencia a proforma origen
✅ Trazabilidad completa a normas ISO

---

## 🎯 RESUMEN PARA AUDITORÍA

**Sistema 100% auditable que:**

1. ✅ Registra automáticamente cada operación
2. ✅ Mantiene trazabilidad completa
3. ✅ Genera números únicos imposibles de duplicar
4. ✅ Exporta evidencias en formato estándar
5. ✅ Cumple ISO/IEC 17025:2017
6. ✅ Preserva integridad de registros

---

**Última actualización:** Febrero 2026
**Versión del sistema:** 2.0
**Estado:** Producción - Compliant INACAL
