"""
extract_proforma.py - Extractor GENÉRICO de proformas Metromecanica
Versión 2.0 - Funciona con cualquier tipo de servicio
"""
import re, sys, json, argparse
import pdfplumber

def clean(t): return " ".join(t.split()).strip()

def find(text, pattern, group=1, default="", flags=re.IGNORECASE|re.DOTALL):
    m = re.search(pattern, text, flags)
    return clean(m.group(group)) if m else default

def extract_proforma(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    # ═══ DATOS BÁSICOS ═══════════════════════════════════════════════════════
    numero_proforma = find(full_text, r'(P\d+\-\d+)', default="SIN-NUM")
    fecha_emision   = find(full_text, r'(\d{2}/\d{2}/\d{4})', default="")

    # ═══ CLIENTE ═════════════════════════════════════════════════════════════
    cliente_raw = find(full_text, r'Se[nñ]or\(es\)\s*:\s*(.+?)(?:Direcci)', default="")
    cliente     = re.split(r'Atenci', cliente_raw, flags=re.IGNORECASE)[0].strip()
    
    dir_raw     = find(full_text, r'Direcci[oó]n\s*:\s*(.+?)(?:R\.U\.C\.)', default="")
    direccion   = re.split(r'@|Atenci', dir_raw, flags=re.IGNORECASE)[0].strip()
    
    ruc_cliente = find(full_text, r'R\.U\.C\.\s*:\s*(\d{11})', default="")

    all_emails      = re.findall(r'[\w.\-]+@[\w.\-]+\.\w+', full_text)
    emails_cliente  = [e for e in all_emails if 'metromecanica' not in e.lower()]
    email_cliente   = emails_cliente[0] if emails_cliente else ""
    
    telefono_cliente= find(full_text, r'Tel[eé]fono\s*:\s*([\d\s]+)', default="").strip()

    # Contacto = columna "Garantia" de la proforma
    cm = re.search(r'\d+\s+a?\s*\d+\s+D[IÍ]AS\s+([\w][\w\s]+?)(?:\n|$)', full_text, re.IGNORECASE)
    contacto = clean(cm.group(1)) if cm else (
        email_cliente.split('@')[0].replace('.',' ').replace('_',' ').title() if email_cliente else ""
    )

    forma_pago    = find(full_text, r'(CR[EÉ]DITO\s+\d+\s+D[IÍ]AS)', default="")
    plazo_entrega = find(full_text, r'(\d+\s+a\s+\d+\s+D[IÍ]AS|\d+\s+a\s+\d+\s+dias)', default="")

    # ═══ EXTRACCIÓN GENÉRICA DE ITEMS ════════════════════════════════════════
    # Detectar todos los items de la tabla
    items_matches = re.findall(
        r'(\d+)\s+(\d+\.\d+)\s+(ZZ|NIU|UND|GLB)\s+(.+?)(?=\d+\.\d+\s+\d+\.\d+)',
        full_text,
        re.DOTALL
    )
    
    items = []
    equipos_set = set()
    tipo_servicio = "GENERICO"
    
    for item_num, cantidad, um, descripcion_raw in items_matches:
        desc_clean = clean(descripcion_raw)
        
        # Detectar tipo de servicio por palabras clave
        if any(kw in desc_clean.upper() for kw in ['BATERIA', 'BATTERY']):
            tipo_servicio = "REEMPLAZO_COMPONENTE"
            # Buscar patron combinado "300 KG Y 500 KG" en todo el texto
            combo = re.search(r'(\d+)\s*KG\s+Y\s+(\d+)\s*KG', full_text, re.IGNORECASE)
            if combo:
                equipos_set.add(f"BALANZA {combo.group(1)} KG")
                equipos_set.add(f"BALANZA {combo.group(2)} KG")
            else:
                # Extraer equipos mencionados en la descripción
                equipos = re.findall(r'BALANZA\s+\d+\s*KG', desc_clean.upper())
                equipos_set.update(equipos)
        
        elif any(kw in desc_clean.upper() for kw in ['CALIBR', 'VERIFIC', 'MICROMETRO', 'DINAMOMETRO', 'MANOMETRO', 'TERMOMETRO', 'DUROMETRO']):
            tipo_servicio = "CALIBRACION"
            # Extraer código del instrumento (ej: IM-012)
            codigo_match = re.search(r'(IM-\d+)', desc_clean)
            codigo = codigo_match.group(1) if codigo_match else f"ITEM-{item_num}"
            
            # Extraer área/ubicación (ej: MAESTRANZA, CONTROL DE CALIDAD)
            area_match = re.search(r'/\s*([A-Z\s]+(?:MAESTRANZA|CALIDAD|MOLINO|AUTOCLAVE|ENSAMBLE)[A-Z\s]*)\s*$', desc_clean)
            area = clean(area_match.group(1)) if area_match else "LABORATORIO"
            equipos_set.add(f"{codigo} - {area}")
        
        items.append({
            "item": int(item_num),
            "cantidad": float(cantidad),
            "um": um,
            "descripcion": desc_clean
        })
    
    # Si no se encontraron items con el patrón, intentar extracción más simple
    if not items:
        # Fallback: buscar línea por línea en la sección de items
        items_section = re.search(r'Item.*?Descripci[oó]n.*?([\d].+?)(?=SON:|Venta Gravada)', full_text, re.DOTALL | re.IGNORECASE)
        if items_section:
            items.append({
                "item": 1,
                "cantidad": 1,
                "um": "UND",
                "descripcion": clean(items_section.group(1)[:200])
            })
    
    # ═══ RESUMEN DEL SERVICIO ════════════════════════════════════════════════
    total_items = len(items)
    equipos_list = sorted(list(equipos_set)) if equipos_set else ["Ver tabla de ítems"]
    
    # Descripción resumida según tipo de servicio
    if tipo_servicio == "CALIBRACION":
        if total_items == 1:
            descripcion_servicio = f"Calibración de {items[0]['descripcion']}"
        else:
            desc_corta = items[0]['descripcion'].split('/')[1].strip() if '/' in items[0]['descripcion'] else items[0]['descripcion'][:50]
            descripcion_servicio = f"Calibración de instrumentos de medición ({total_items} equipos)"
    
    elif tipo_servicio == "REEMPLAZO_COMPONENTE":
        # Extraer el componente principal
        primer_item = items[0]['descripcion']
        componente = re.search(r'COMPRA DE (.+?)(?:PARA|$)', primer_item, re.IGNORECASE)
        if componente:
            descripcion_servicio = f"Suministro e instalación de {componente.group(1).strip()}"
        else:
            descripcion_servicio = primer_item[:100]
    
    else:
        descripcion_servicio = items[0]['descripcion'][:100] if items else "Servicio técnico especializado"
    
    # ═══ ALCANCE DEL SERVICIO (desde observaciones) ══════════════════════════
    obs_text = find(full_text, r'Observaciones\s*:\s*(.*?)(?:"SIRVASE|GRACIAS|www\.)', default="")
    
    alcance_items = []
    actividades = []
    
    for linea in obs_text.split('\n'):
        linea_clean = linea.strip().lstrip('-').strip()
        excluir = ['COSTO', 'IGV', 'PRECIO', 'INCLUYE IGV', 'S/.', 'OBSERVACIONES:', 'ABONAR']
        
        if linea_clean and not any(k in linea_clean.upper() for k in excluir) and len(linea_clean) > 10:
            # Limpiar prefijos comunes
            linea_clean = re.sub(r'^(EL\s+SERVICIO|SE)\s+(REALIZARA?|INCLUYE?)\s+', '', linea_clean, flags=re.IGNORECASE)
            
            # Para actividades (checkboxes)
            if tipo_servicio == "CALIBRACION":
                actividades.append(linea_clean.capitalize())
            else:
                actividades.append(linea_clean.title())
            
            # Para alcance
            alcance_items.append(linea_clean.capitalize())
    
    # Valores por defecto si no hay observaciones
    if not alcance_items:
        if tipo_servicio == "CALIBRACION":
            alcance_items = [
                "Calibración con patrones trazables a INACAL",
                "Emisión de certificado de calibración",
                "Etiquetas de identificación"
            ]
        else:
            alcance_items = ["Servicio completo según especificación técnica"]
    
    if not actividades:
        actividades = alcance_items.copy()
    
    alcance_servicio = " | ".join(alcance_items)

    # ═══ RETORNO ══════════════════════════════════════════════════════════════
    return {
        "aprobada": True,
        "numero_proforma": numero_proforma,
        "fecha_emision": fecha_emision,
        "cliente": cliente,
        "direccion_cliente": direccion,
        "ruc_cliente": ruc_cliente,
        "contacto_cliente": contacto,
        "email_cliente": email_cliente,
        "telefono_cliente": telefono_cliente,
        "forma_pago": forma_pago,
        "plazo_entrega": plazo_entrega,
        
        # Servicio
        "tipo_servicio": tipo_servicio,
        "total_items": total_items,
        "descripcion_servicio": descripcion_servicio,
        "alcance_servicio": alcance_servicio,
        "items": items,
        "equipos": equipos_list,
        "actividades_incluidas": actividades,
        
        # Laboratorio
        "laboratorio": "METROMECANICA - Metrología y Calibración SAC",
        "ruc_laboratorio": "20605421696",
    }

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("--estado", choices=["aprobada","rechazada"], default="aprobada")
    args = parser.parse_args()

    if args.estado == "rechazada":
        print(json.dumps({"aprobada": False, "mensaje": "Proforma RECHAZADA."}, ensure_ascii=False))
        sys.exit(0)

    try:
        data = extract_proforma(args.pdf)
        print(json.dumps(data, ensure_ascii=False, indent=2))
    except Exception as e:
        print(json.dumps({"error": str(e), "aprobada": False}), file=sys.stderr)
        sys.exit(1)
