/**
 * generate_ot_v2.js — Generador de OT con diseño profesional renovado
 * ══════════════════════════════════════════════════════════════════════════════
 * Versión 2.0 — Diseño sofisticado y moderno manteniendo trazabilidad ISO 17025
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  Header, PageNumber, NumberFormat, TabStopType, TabStopPosition
} = require('docx');
const fs   = require('fs');
const path = require('path');

// ═══ PALETA DE COLORES ═══════════════════════════════════════════════════════
const COLOR = {
  // Azules corporativos (degradado de oscuro a claro)
  navy:      "0A2540",  // Azul marino profundo
  primary:   "1E3A5F",  // Azul corporativo principal
  mid:       "2E5C8A",  // Azul medio
  accent:    "4A90E2",  // Azul brillante para destacados
  lightBlue: "E8F2F7",  // Azul muy claro para fondos
  
  // Acentos
  orange:    "FF6B35",  // Naranja vibrante para alertas/destacados
  amber:     "FFA500",  // Ambar para warnings
  
  // Grises técnicos
  slate:     "2D3748",  // Gris oscuro para texto principal
  gray:      "4A5568",  // Gris medio para texto secundario
  lightGray: "E2E8F0",  // Gris claro para bordes
  offWhite:  "F7FAFC",  // Casi blanco para fondos alternos
  
  // Estados
  success:   "48BB78",  // Verde para confirmaciones
  white:     "FFFFFF",
};

// ═══ UTILIDADES ══════════════════════════════════════════════════════════════
const DXA_INCH = 1440;
const PAGE_WIDTH = 11906;  // A4 width
const MARGIN = 850;
const CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN);

function border(color = COLOR.lightGray, size = 6) {
  return { style: BorderStyle.SINGLE, size, color };
}

function borders(color, size) {
  const b = border(color, size);
  return { top: b, bottom: b, left: b, right: b };
}

function noBorders() {
  const b = { style: BorderStyle.NONE, size: 0 };
  return { top: b, bottom: b, left: b, right: b };
}

// ═══ COMPONENTES DE TEXTO ════════════════════════════════════════════════════
function text(content, {
  size = 20,
  bold = false,
  color = COLOR.slate,
  font = "Aptos",
  caps = false,
  italic = false,
  underline = false
} = {}) {
  return new TextRun({
    text: content,
    font, size, bold, color,
    allCaps: caps,
    italics: italic,
    underline: underline ? {} : undefined
  });
}

function para(content, {
  align = AlignmentType.LEFT,
  spacing = {},
  indent = {},
  bold = false,
  size = 20,
  color = COLOR.slate
} = {}) {
  const children = typeof content === 'string'
    ? [text(content, { bold, size, color })]
    : (Array.isArray(content) ? content : [content]);
    
  return new Paragraph({ alignment: align, spacing, indent, children });
}

// ═══ COMPONENTES DE TABLA ════════════════════════════════════════════════════
function cell(content, {
  width,
  borders: b,
  shading,
  vAlign = VerticalAlign.CENTER,
  span,
  margins = { top: 120, bottom: 120, left: 180, right: 180 }
} = {}) {
  const children = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    borders: b || borders(COLOR.lightGray, 4),
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    verticalAlign: vAlign,
    columnSpan: span,
    margins,
    children
  });
}

// ═══ GENERADOR DE NÚMERO CORRELATIVO ═════════════════════════════════════════
function generateOTNumber(proforma_num, fecha_emision) {
  // Extraer año de la fecha de emisión
  const year = fecha_emision ? fecha_emision.split('/')[2] : new Date().getFullYear();
  
  // Generar número correlativo basado en timestamp para unicidad
  const timestamp = Date.now();
  const seq = String(timestamp).slice(-4);
  
  return {
    ot_number: `OT-${year}-${seq}`,
    expediente: `${timestamp}`,
    codigo_doc: `RTL-01/Ed02-${year}/LAB`
  };
}

// ═══ ENCABEZADO DEL DOCUMENTO ════════════════════════════════════════════════
function buildDocumentHeader(data, otInfo) {
  const headerHeight = 1100;
  
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    borders: {
      top: border(COLOR.primary, 12),
      bottom: border(COLOR.primary, 2),
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideH: { style: BorderStyle.NONE },
      insideV: border(COLOR.lightGray, 4)
    },
    rows: [
      new TableRow({
        height: { value: headerHeight, rule: 0 },
        children: [
          // Columna 1: Logo + Empresa
          cell([
            new Paragraph({
              children: [
                new ImageRun({
                  data: fs.readFileSync(path.join(__dirname, 'logo_metromecanica.png')),
                  transformation: {
                    width: 100,
                    height: 100
                  }
                })
              ],
              spacing: { after: 160 },
              alignment: AlignmentType.CENTER
            }),
            para("METROMECANICA", { bold: true, size: 28, color: COLOR.primary, align: AlignmentType.CENTER }),
            para("Metrología y Calibración SAC", { size: 17, color: COLOR.gray, align: AlignmentType.CENTER }),
            para(" ", { size: 8 }),
            para("Laboratorio de Calibración", { size: 15, color: COLOR.mid, bold: true, align: AlignmentType.CENTER }),
            para("ISO/IEC 17025:2017", { size: 14, color: COLOR.accent, align: AlignmentType.CENTER }),
          ], {
            width: CONTENT_WIDTH * 0.35,
            borders: noBorders(),
            shading: COLOR.offWhite,
            vAlign: VerticalAlign.CENTER
          }),
          
          // Columna 2: Título del documento
          cell([
            para("ORDEN DE TRABAJO", {
              bold: true, size: 28, color: COLOR.navy,
              align: AlignmentType.CENTER
            }),
            para("Documento de Control Interno", {
              size: 16, color: COLOR.gray,
              align: AlignmentType.CENTER,
              spacing: { before: 80 }
            }),
          ], {
            width: CONTENT_WIDTH * 0.40,
            borders: noBorders(),
            vAlign: VerticalAlign.CENTER
          }),
          
          // Columna 3: Códigos y metadata
          cell([
            para(`N°: ${otInfo.ot_number}`, { bold: true, size: 20, color: COLOR.primary }),
            para(`Expediente: ${otInfo.expediente}`, { size: 16, color: COLOR.gray }),
            para(`Código: ${otInfo.codigo_doc}`, { size: 16, color: COLOR.gray }),
            para(" ", { size: 6 }),
            para(`Versión: 02`, { size: 15, color: COLOR.gray }),
            para(`Página: 1 de 1`, { size: 15, color: COLOR.gray }),
          ], {
            width: CONTENT_WIDTH * 0.25,
            borders: noBorders(),
            shading: COLOR.lightBlue,
            vAlign: VerticalAlign.CENTER
          })
        ]
      })
    ]
  });
}

// ═══ BANNER DE ESTADO ═════════════════════════════════════════════════════════
function buildStatusBanner(data, otInfo) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE }
    },
    rows: [
      new TableRow({
        children: [
          cell([
            para("ESTADO: APROBADA", {
              bold: true, size: 18, color: COLOR.white,
              align: AlignmentType.CENTER
            })
          ], {
            width: CONTENT_WIDTH * 0.22,
            shading: COLOR.success,
            borders: noBorders()
          }),
          
          cell([
            para(`Emisión: ${data.fecha_emision}`, {
              size: 17, color: COLOR.white,
              align: AlignmentType.CENTER
            })
          ], {
            width: CONTENT_WIDTH * 0.22,
            shading: COLOR.mid,
            borders: noBorders()
          }),
          
          cell([
            para(`Entrega: ${data.plazo_entrega}`, {
              size: 17, color: COLOR.white,
              align: AlignmentType.CENTER
            })
          ], {
            width: CONTENT_WIDTH * 0.22,
            shading: COLOR.mid,
            borders: noBorders()
          }),
          
          cell([
            para(`Ref. Proforma: ${data.numero_proforma}`, {
              size: 17, color: COLOR.white,
              align: AlignmentType.CENTER
            })
          ], {
            width: CONTENT_WIDTH * 0.34,
            shading: COLOR.primary,
            borders: noBorders()
          })
        ]
      })
    ]
  });
}

// ═══ SECCIÓN: INFORMACIÓN COMERCIAL ══════════════════════════════════════════
function buildSeccion1_InfoComercial(data, otInfo) {
  const secWidth = CONTENT_WIDTH;
  const col1 = secWidth * 0.25;
  const col2 = secWidth * 0.75;
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows: [
      // Header de sección
      new TableRow({
        children: [
          cell([
            para("1. INFORMACIÓN COMERCIAL", {
              bold: true, size: 18, color: COLOR.white, caps: true
            })
          ], {
            width: secWidth,
            shading: COLOR.navy,
            borders: borders(COLOR.navy, 8),
            span: 2
          })
        ]
      }),
      
      // RUC
      new TableRow({
        children: [
          cell([para("RUC", { bold: true, size: 18, color: COLOR.primary })], {
            width: col1,
            shading: COLOR.lightBlue
          }),
          cell([para(data.ruc_cliente, { size: 18 })], { width: col2 })
        ]
      }),
      
      // Razón Social
      new TableRow({
        children: [
          cell([para("Razón Social", { bold: true, size: 18, color: COLOR.primary })], {
            width: col1,
            shading: COLOR.lightBlue
          }),
          cell([para(data.cliente, { size: 18 })], { width: col2 })
        ]
      }),
      
      // Dirección
      new TableRow({
        children: [
          cell([para("Dirección Fiscal", { bold: true, size: 18, color: COLOR.primary })], {
            width: col1,
            shading: COLOR.lightBlue
          }),
          cell([para(data.direccion_cliente, { size: 18 })], { width: col2 })
        ]
      }),
      
      // Contacto + Teléfono (fila combinada)
      new TableRow({
        children: [
          cell([para("Contacto", { bold: true, size: 18, color: COLOR.primary })], {
            width: col1,
            shading: COLOR.lightBlue
          }),
          cell([
            para([
              text(`${data.contacto_cliente}  `, { size: 18, bold: true }),
              text(`· Tel: ${data.telefono_cliente}`, { size: 17, color: COLOR.gray })
            ])
          ], { width: col2 })
        ]
      }),
      
      // Email
      new TableRow({
        children: [
          cell([para("Correo Electrónico", { bold: true, size: 18, color: COLOR.primary })], {
            width: col1,
            shading: COLOR.lightBlue
          }),
          cell([para(data.email_cliente, { size: 18, color: COLOR.accent })], { width: col2 })
        ]
      }),
    ]
  });
}

// ═══ SECCIÓN: OBSERVACIONES PARA EL SERVICIO ═════════════════════════════════
function buildSeccion1_5_Observaciones() {
  const secWidth = CONTENT_WIDTH;
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [
          cell([
            para("OBSERVACIONES PARA EL SERVICIO", {
              bold: true, size: 18, color: COLOR.white, caps: true
            })
          ], {
            width: secWidth,
            shading: COLOR.navy,
            borders: borders(COLOR.navy, 8)
          })
        ]
      }),
      
      new TableRow({
        children: [
          cell([
            para("Indicaciones especiales, condiciones del cliente o notas importantes para la ejecución:", {
              size: 16, color: COLOR.gray, italic: true, spacing: { after: 80 }
            }),
            para(" ", { size: 18 }),
            para(" ", { size: 18 }),
            para(" ", { size: 18 }),
            para(" ", { size: 18 }),
            para(" ", { size: 18 }),
          ], {
            width: secWidth,
            shading: COLOR.offWhite,
            borders: borders(COLOR.lightGray, 4)
          })
        ]
      })
    ]
  });
}

// ═══ SECCIÓN: DESCRIPCIÓN DEL SERVICIO ═══════════════════════════════════════
function buildSeccion2_Servicio(data) {
  const secWidth = CONTENT_WIDTH;
  const rows = [];
  const items = data.items || [];
  const total_items = data.total_items || items.length || 1;
  
  // Header
  rows.push(new TableRow({
    children: [
      cell([
        para("2. DESCRIPCIÓN DEL SERVICIO", {
          bold: true, size: 18, color: COLOR.white, caps: true
        })
      ], {
        width: secWidth,
        shading: COLOR.navy,
        borders: borders(COLOR.navy, 8),
        span: 5
      })
    ]
  }));
  
  // Subheaders
  rows.push(new TableRow({
    children: [
      cell([para("N°", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
        width: 600,
        shading: COLOR.lightBlue
      }),
      cell([para("CANT.", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
        width: 900,
        shading: COLOR.lightBlue
      }),
      cell([para("U/M", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
        width: 900,
        shading: COLOR.lightBlue
      }),
      cell([para("DESCRIPCIÓN DEL INSTRUMENTO / COMPONENTE", { bold: true, size: 16, color: COLOR.primary })], {
        width: secWidth - 5400,
        shading: COLOR.lightBlue
      }),
      cell([para("CERTIFICADO ASIGNADO", { bold: true, size: 16, color: COLOR.primary })], {
        width: 3000,
        shading: COLOR.lightBlue
      })
    ]
  }));
  
  // Determinar si es servicio simple o múltiple
  if (total_items <= 3 && items.length > 0) {
    // CASO SIMPLE: Mostrar cada ítem completo con alcance
    items.forEach((item, idx) => {
      const equipos_text = data.equipos.join(" · ");
      
      rows.push(new TableRow({
        children: [
          cell([para(item.item.toString(), { size: 18, align: AlignmentType.CENTER })], { width: 600 }),
          cell([para(item.cantidad.toString(), { size: 18, align: AlignmentType.CENTER })], { width: 900 }),
          cell([para(item.um || "UND", { size: 17, align: AlignmentType.CENTER })], { width: 900 }),
          cell([
            para(item.descripcion, { bold: true, size: 17 }),
            para(" ", { size: 60 }),
            para("Alcance del servicio:", { 
              bold: true, 
              size: 16, 
              color: COLOR.primary, 
              spacing: { before: 80 } 
            }),
            para(data.alcance_servicio || "Servicio completo según especificación", { 
              size: 17, 
              color: COLOR.slate,
              spacing: { before: 60 }
            })
          ], { width: secWidth - 5400 }),
          cell([
            para("", { size: 17 })
          ], { width: 3000, shading: COLOR.offWhite })
        ]
      }));
    });
  } else {
    // CASO MÚLTIPLE: Tabla compacta de todos los ítems
    items.forEach((item, idx) => {
      const shade = idx % 2 === 0 ? COLOR.white : COLOR.offWhite;
      
      rows.push(new TableRow({
        children: [
          cell([para(item.item.toString(), { size: 17, align: AlignmentType.CENTER })], { width: 600, shading: shade }),
          cell([para(item.cantidad.toString(), { size: 17, align: AlignmentType.CENTER })], { width: 900, shading: shade }),
          cell([para(item.um || "UND", { size: 16, align: AlignmentType.CENTER })], { width: 900, shading: shade }),
          cell([para(item.descripcion, { size: 17 })], { width: secWidth - 5400, shading: shade }),
          cell([para("", { size: 16 })], { width: 3000, shading: shade })
        ]
      }));
    });
    
    // Fila de alcance general al final
    rows.push(new TableRow({
      children: [
        cell([
          para("ALCANCE GENERAL DEL SERVICIO", { 
            bold: true, 
            size: 16, 
            color: COLOR.primary 
          }),
          para(data.alcance_servicio || "Servicio completo según especificación", { 
            size: 17, 
            spacing: { before: 80 }
          })
        ], { width: secWidth, span: 5, shading: COLOR.lightBlue })
      ]
    }));
  }
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows
  });
}

// ═══ SECCIÓN: ÁREAS RESPONSABLES ═════════════════════════════════════════════
function buildSeccion3_Areas() {
  const secWidth = CONTENT_WIDTH;
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [
          cell([
            para("3. ÁREAS RESPONSABLES", {
              bold: true, size: 18, color: COLOR.white, caps: true
            })
          ], {
            width: secWidth,
            shading: COLOR.navy,
            borders: borders(COLOR.navy, 8),
            span: 2
          })
        ]
      }),
      
      new TableRow({
        children: [
          cell([
            para("Área Ejecutora", { bold: true, size: 17, color: COLOR.primary }),
            para("laboratorio@metromecanica.com.pe", { size: 17, color: COLOR.accent, spacing: { before: 60 } })
          ], {
            width: secWidth / 2,
            shading: COLOR.lightBlue
          }),
          cell([
            para("Área de Calidad", { bold: true, size: 17, color: COLOR.primary }),
            para("calidad@metromecanica.com.pe", { size: 17, color: COLOR.accent, spacing: { before: 60 } })
          ], {
            width: secWidth / 2,
            shading: COLOR.lightBlue
          })
        ]
      }),
      
      new TableRow({
        children: [
          cell([
            para("Coordinador Asignado", { bold: true, size: 17, color: COLOR.primary }),
            para("_________________________________", { size: 18, spacing: { before: 80 } })
          ], { width: secWidth / 2 }),
          cell([
            para("Fecha de Asignación", { bold: true, size: 17, color: COLOR.primary }),
            para("______ / ______ / 2026", { size: 18, spacing: { before: 80 } })
          ], { width: secWidth / 2 })
        ]
      })
    ]
  });
}

// ═══ SECCIÓN: ACTIVIDADES TÉCNICAS ═══════════════════════════════════════════
function buildSeccion4_Actividades(data) {
  const secWidth = CONTENT_WIDTH;
  const actividades = data.actividades_incluidas || [];
  
  const actRows = actividades.map(act => 
    para([
      text("☐  ", { size: 22, color: COLOR.primary }),
      text(act, { size: 18 })
    ], { spacing: { before: 100 } })
  );
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [
          cell([
            para("4. ACTIVIDADES TÉCNICAS INCLUIDAS", {
              bold: true, size: 18, color: COLOR.white, caps: true
            })
          ], {
            width: secWidth,
            shading: COLOR.navy,
            borders: borders(COLOR.navy, 8)
          })
        ]
      }),
      
      new TableRow({
        children: [
          cell([
            para("El servicio comprende las siguientes actividades:", {
              size: 17, color: COLOR.gray, italic: true,
              spacing: { after: 120 }
            }),
            ...actRows,
            para(" ", { size: 140 }),
            para([
              text("☐  ", { size: 22, color: COLOR.gray }),
              text("Verificación metrológica post-servicio (si aplica)", { size: 17, color: COLOR.gray, italic: true })
            ])
          ], { width: secWidth })
        ]
      })
    ]
  });
}

// ═══ SECCIÓN: REQUISITOS TÉCNICOS ISO 17025 ══════════════════════════════════
function buildSeccion5_RequisitosISO(data) {
  const secWidth = CONTENT_WIDTH;
  const col1 = secWidth * 0.35;
  const col2 = secWidth * 0.65;
  
  const requisitos = [
    ["Tipo de Servicio", "Mantenimiento preventivo de equipos de medición"],
    ["Normas de Referencia", "OIML R 76 / ASTM E617 / Especificaciones del fabricante"],
    ["Equipos Intervenidos", data.equipos.join(" · ")],
    ["Componente Instalado", data.descripcion_componente],
    ["Trazabilidad Metrológica", "Patrones calibrados trazables a INACAL / BIPM"],
    ["Condiciones Ambientales", "Temperatura: 18-28°C  ·  Humedad: 40-70%  ·  Sin vibraciones"],
    ["Registro Ambiental", "Obligatorio (Cláusula 6.3 - ISO/IEC 17025:2017)"],
    ["Personal Competente", "Técnico certificado en calibración de equipos de pesaje"],
    ["EPP Requerido", "Guantes dieléctricos · Lentes de seguridad · Calzado de seguridad"]
  ];
  
  const rows = [
    new TableRow({
      children: [
        cell([
          para("5. REQUISITOS TÉCNICOS — ISO/IEC 17025:2017", {
            bold: true, size: 18, color: COLOR.white, caps: true
          })
        ], {
          width: secWidth,
          shading: COLOR.navy,
          borders: borders(COLOR.navy, 8),
          span: 2
        })
      ]
    })
  ];
  
  requisitos.forEach(([label, value]) => {
    rows.push(new TableRow({
      children: [
        cell([para(label, { bold: true, size: 17, color: COLOR.primary })], {
          width: col1,
          shading: COLOR.lightBlue
        }),
        cell([para(value, { size: 18 })], { width: col2 })
      ]
    }));
  });
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows
  });
}

// ═══ SECCIÓN: REGISTRO DE EJECUCIÓN ══════════════════════════════════════════
function buildSeccion6_Ejecucion() {
  const secWidth = CONTENT_WIDTH;
  const col1 = secWidth * 0.28;
  const col2 = secWidth * 0.22;
  const col3 = secWidth * 0.28;
  const col4 = secWidth * 0.22;
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [
          cell([
            para("6. REGISTRO DE EJECUCIÓN (Cláusula 7.5 - ISO/IEC 17025:2017)", {
              bold: true, size: 18, color: COLOR.white, caps: true
            })
          ], {
            width: secWidth,
            shading: COLOR.navy,
            borders: borders(COLOR.navy, 8),
            span: 4
          })
        ]
      }),
      
      // Fila 1: Fechas
      new TableRow({
        children: [
          cell([para("Fecha de Inicio", { bold: true, size: 17, color: COLOR.primary })], {
            width: col1, shading: COLOR.lightBlue
          }),
          cell([para("______ / ______ / 2026", { size: 18 })], { width: col2 }),
          cell([para("Hora de Inicio", { bold: true, size: 17, color: COLOR.primary })], {
            width: col3, shading: COLOR.lightBlue
          }),
          cell([para("_______ : _______", { size: 18 })], { width: col4 })
        ]
      }),
      
      new TableRow({
        children: [
          cell([para("Fecha de Término", { bold: true, size: 17, color: COLOR.primary })], {
            width: col1, shading: COLOR.lightBlue
          }),
          cell([para("______ / ______ / 2026", { size: 18 })], { width: col2 }),
          cell([para("Hora de Término", { bold: true, size: 17, color: COLOR.primary })], {
            width: col3, shading: COLOR.lightBlue
          }),
          cell([para("_______ : _______", { size: 18 })], { width: col4 })
        ]
      }),
      
      // Fila 2: Condiciones ambientales
      new TableRow({
        children: [
          cell([para("Temperatura (°C)", { bold: true, size: 17, color: COLOR.primary })], {
            width: col1, shading: COLOR.lightBlue
          }),
          cell([para("_________ °C", { size: 18 })], { width: col2 }),
          cell([para("Humedad Relativa (%)", { bold: true, size: 17, color: COLOR.primary })], {
            width: col3, shading: COLOR.lightBlue
          }),
          cell([para("_________ %", { size: 18 })], { width: col4 })
        ]
      }),
      
      // Fila 3: Ubicación
      new TableRow({
        children: [
          cell([para("Lugar de Ejecución", { bold: true, size: 17, color: COLOR.primary })], {
            width: col1, shading: COLOR.lightBlue
          }),
          cell([
            para("☐ Instalaciones del cliente    ☐ Laboratorio", { size: 17 })
          ], { width: col2 + col3 + col4, span: 3 })
        ]
      }),
      
      // Observaciones
      new TableRow({
        children: [
          cell([para("Observaciones Técnicas", { bold: true, size: 17, color: COLOR.primary })], {
            width: col1, shading: COLOR.lightBlue
          }),
          cell([
            para("", { size: 18 }),
            para("", { size: 18 }),
            para("", { size: 18 }),
          ], { width: col2 + col3 + col4, span: 3, shading: COLOR.offWhite })
        ]
      })
    ]
  });
}

// ═══ SECCIÓN: VERIFICACIÓN POST-SERVICIO ═════════════════════════════════════
function buildSeccion7_Verificacion() {
  const secWidth = CONTENT_WIDTH;
  const criterios = [
    "Batería instalada con polaridad correcta",
    "Tensión verificada: 6.0 ± 0.3 V (multímetro calibrado)",
    "Limpieza interior ejecutada según procedimiento",
    "Lubricación de partes móviles realizada",
    "Equipo enciende y opera normalmente",
    "Indicador de nivel de batería en rango óptimo",
    "Lectura de tara verificada (0.000 kg en vacío)",
    "Sin daños en componentes adyacentes",
    "Cubierta cerrada y asegurada correctamente",
    "Documentación técnica completada y firmada"
  ];
  
  const rows = [
    new TableRow({
      children: [
        cell([
          para("7. VERIFICACIÓN POST-SERVICIO — LISTA DE CHEQUEO", {
            bold: true, size: 18, color: COLOR.white, caps: true
          })
        ], {
          width: secWidth,
          shading: COLOR.navy,
          borders: borders(COLOR.navy, 8),
          span: 5
        })
      ]
    }),
    
    new TableRow({
      children: [
        cell([para("CRITERIO DE VERIFICACIÓN", { bold: true, size: 16, color: COLOR.primary })], {
          width: secWidth * 0.50, shading: COLOR.lightBlue
        }),
        cell([para("C", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
          width: secWidth * 0.12, shading: COLOR.lightBlue
        }),
        cell([para("NC", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
          width: secWidth * 0.12, shading: COLOR.lightBlue
        }),
        cell([para("N/A", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
          width: secWidth * 0.12, shading: COLOR.lightBlue
        }),
        cell([para("OBS.", { bold: true, size: 16, color: COLOR.primary, align: AlignmentType.CENTER })], {
          width: secWidth * 0.14, shading: COLOR.lightBlue
        })
      ]
    })
  ];
  
  criterios.forEach((crit, idx) => {
    const shade = idx % 2 === 0 ? COLOR.offWhite : COLOR.white;
    rows.push(new TableRow({
      children: [
        cell([para(crit, { size: 18 })], { width: secWidth * 0.50, shading: shade }),
        cell([para("☐", { size: 22, align: AlignmentType.CENTER })], { width: secWidth * 0.12 }),
        cell([para("☐", { size: 22, align: AlignmentType.CENTER })], { width: secWidth * 0.12 }),
        cell([para("☐", { size: 22, align: AlignmentType.CENTER })], { width: secWidth * 0.12 }),
        cell([para("", { size: 18 })], { width: secWidth * 0.14, shading: COLOR.offWhite })
      ]
    }));
  });
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows
  });
}

// ═══ SECCIÓN: APROBACIONES Y FIRMAS ══════════════════════════════════════════
function buildSeccion6_Firmas() {
  const secWidth = CONTENT_WIDTH;
  const colWidth = secWidth / 3;
  
  const firmaBox = (titulo, campos) => {
    const lines = campos.map(c => para(c, { size: 17, spacing: { before: 100 } }));
    return cell([
      para(titulo, { bold: true, size: 17, color: COLOR.primary, align: AlignmentType.CENTER }),
      para(" ", { size: 200 }),
      para("_________________________________", { size: 18, align: AlignmentType.CENTER }),
      para("Firma / Sello", { size: 15, color: COLOR.gray, align: AlignmentType.CENTER, spacing: { before: 40 } }),
      para(" ", { size: 120 }),
      ...lines
    ], {
      width: colWidth,
      shading: COLOR.offWhite
    });
  };
  
  return new Table({
    width: { size: secWidth, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: [
          cell([
            para("6. APROBACIONES Y FIRMAS (Cláusula 5.5 - ISO/IEC 17025:2017)", {
              bold: true, size: 18, color: COLOR.white, caps: true
            })
          ], {
            width: secWidth,
            shading: COLOR.navy,
            borders: borders(COLOR.navy, 8),
            span: 3
          })
        ]
      }),
      
      new TableRow({
        children: [
          firmaBox("TÉCNICO EJECUTOR", [
            "Nombre: ________________________",
            "Código: ________________________",
            "Fecha:  ________________________"
          ]),
          firmaBox("SUPERVISOR / JEFE DE LABORATORIO", [
            "Nombre: ________________________",
            "Cargo:  ________________________",
            "Fecha:  ________________________"
          ]),
          firmaBox("CONFORMIDAD DEL CLIENTE", [
            "Nombre: ________________________",
            "DNI:    ________________________",
            "Fecha:  ________________________"
          ])
        ]
      })
    ]
  });
}

// ═══ PIE DE PÁGINA ════════════════════════════════════════════════════════════
function buildFooter(otInfo) {
  return new Paragraph({
    border: { top: border(COLOR.lightGray, 4) },
    spacing: { before: 200, after: 100 },
    alignment: AlignmentType.CENTER,
    children: [
      text("METROMECANICA - Metrología y Calibración SAC  ·  ", { size: 16, color: COLOR.gray }),
      text("Psj. 18 de Enero Mz. LL Lte 3, Urb. Bambeta Baja Este, Sector 5, Callao  ·  ", { size: 16, color: COLOR.gray }),
      text("Telef: 940 255 997 / 980 762 761", { size: 16, color: COLOR.gray }),
      text("\n", { size: 12 }),
      text(otInfo.codigo_doc, { size: 14, color: COLOR.mid })
    ]
  });
}

// ═══ GENERADOR PRINCIPAL ═════════════════════════════════════════════════════
function generateOT(data) {
  if (!data.aprobada) {
    console.error("Proforma no aprobada. No se genera OT.");
    process.exit(0);
  }
  
  const otInfo = generateOTNumber(data.numero_proforma, data.fecha_emision);
  
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Aptos", size: 20, color: COLOR.slate }
        }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: 16838 },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
        }
      },
      children: [
        buildDocumentHeader(data, otInfo),
        para(" ", { spacing: { before: 180 } }),
        buildStatusBanner(data, otInfo),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion1_InfoComercial(data, otInfo),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion1_5_Observaciones(),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion2_Servicio(data),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion3_Areas(),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion4_Actividades(data),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion5_RequisitosISO(data),
        para(" ", { spacing: { before: 240 } }),
        buildSeccion6_Firmas(),
        para(" ", { spacing: { before: 200 } }),
        buildFooter(otInfo)
      ]
    }]
  });
  
  return { doc, ot_num: otInfo.ot_number };
}

// ═══ MAIN ═════════════════════════════════════════════════════════════════════
async function main() {
  let data;
  
  if (process.argv.includes('--stdin')) {
    const chunks = [];
    process.stdin.on('data', c => chunks.push(c));
    await new Promise(r => process.stdin.on('end', r));
    data = JSON.parse(Buffer.concat(chunks).toString());
  } else if (process.argv.includes('--file')) {
    const idx  = process.argv.indexOf('--file');
    const file = process.argv[idx + 1];
    data = JSON.parse(fs.readFileSync(file, 'utf8'));
  } else if (process.argv[2] && process.argv[2].startsWith('{')) {
    data = JSON.parse(process.argv[2]);
  } else {
    console.error("Uso: node generate_ot.js '<json>' | --file data.json | --stdin");
    process.exit(1);
  }
  
  const { doc, ot_num } = generateOT(data);
  
  // Crear carpeta ordenes_generadas si no existe
  const outputDir = path.join(__dirname, 'ordenes_generadas');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  const outPath = path.join(outputDir, `${ot_num}.docx`);
  const buffer  = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
  
  console.log(`OK: ${outPath}`);
  return outPath;
}

main().catch(e => { console.error(e); process.exit(1); });
