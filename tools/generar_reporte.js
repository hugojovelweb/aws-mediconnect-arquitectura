/**
 * MediConnect - Reporte Técnico de Arquitectura AWS (Word editable)
 * Kodigo Full Stack Jr 36
 *
 * Uso:   npm install docx   (una sola vez)
 *        node generar_reporte.js
 * Salida: MediConnect_Reporte_Tecnico_AWS.docx
 * Requiere en la misma carpeta: MediConnect_Diagrama_Arquitectura_AWS.png
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak, SimpleField, ExternalHyperlink,
  TabStopType, TabStopPosition,
} = require("docx");

/* ------------------------------------------------------------------ */
/* Modelo de costos (USD / mes) - us-east-1, 730 h/mes                 */
/* ------------------------------------------------------------------ */
const H = 730;
const FARGATE_VCPU = 0.04048, FARGATE_GB = 0.004445;
const task = (vcpu, gb) => H * (vcpu * FARGATE_VCPU + gb * FARGATE_GB);          // 24/7
const taskHours = (vcpu, gb, h) => h * (vcpu * FARGATE_VCPU + gb * FARGATE_GB);
const MONDAY_H = 6 * 4.33;                                                       // 6 h x lunes

const fargate =
  2 * task(1, 2) +          // API Laravel base (2 tareas 1 vCPU / 2 GB)
  2 * task(0.5, 1) +        // Portal paciente base
  2 * task(0.5, 1) +        // Dashboard médico base
  1 * task(0.5, 1) +        // Worker notificaciones
  8 * taskHours(1, 2, MONDAY_H) +      // +8 tareas API los lunes
  6 * taskHours(0.5, 1, MONDAY_H) +    // +6 tareas portal los lunes
  3 * taskHours(0.5, 1, MONDAY_H) +    // +3 tareas dashboard los lunes
  4.33 * taskHours(2, 8, 0.25);        // worker reportes 15 min/semana

// [servicio, configuración / supuesto, costo, tipo]  tipo: F = base casi fijo, V = variable por uso
const costs = [
  ["Amazon Route 53", "1 zona hospedada + 10 M consultas DNS/mes", 4.5, "V"],
  ["Amazon CloudFront", "800 GB salida/mes x $0.085 + 20 M solicitudes HTTPS x $0.01/10 k", 88.0, "V"],
  ["AWS WAF", "2 Web ACL ($5) + 10 reglas ($1) + 30 M solicitudes x $0.60/M", 38.0, "F"],
  ["Elastic Load Balancing (ALB)", "1 ALB Multi-AZ: 730 h x $0.0225 + ~4 LCU promedio x $0.008/h", 39.79, "F"],
  ["NAT Gateway (x2, una por AZ)", "2 x 730 h x $0.045 + 300 GB procesados x $0.045", 79.2, "F"],
  ["Amazon ECS en AWS Fargate", "API 2 tareas (1 vCPU/2 GB), portal 2 y dashboard 2 (0.5 vCPU/1 GB), worker 1; refuerzo lunes 6 h (+8/+6/+3 tareas); worker de reportes 2 vCPU/8 GB x 15 min/semana", 178.36, "F"],
  ["Amazon ECR", "20 GB de imágenes x $0.10", 2.0, "F"],
  ["Amazon RDS for MySQL 8.4 (primaria)", "db.m6g.large Multi-AZ, 200 GB gp3, 100 GB de backup adicional, bajo demanda (valor de la AWS Pricing Calculator)", 277.42, "F"],
  ["Amazon RDS - réplica de lectura", "db.m6g.large Single-AZ, 200 GB gp3 (valor de la AWS Pricing Calculator)", 133.96, "F"],
  ["Amazon ElastiCache for Redis", "2 nodos cache.t4g.small (Multi-AZ): 2 x 730 h x $0.032", 46.72, "F"],
  ["Amazon S3 (archivos médicos y reportes)", "1 TB Standard en us-east-1 + 1 TB de réplica en us-east-2 + 50 GB de reportes, con solicitudes y replicación (valor de la AWS Pricing Calculator)", 47.85, "V"],
  ["Amazon SQS (+ DLQ)", "10 M solicitudes/mes x $0.40/M (1 M gratis)", 3.6, "V"],
  ["Amazon EventBridge Scheduler", "1 programación semanal (dentro del nivel gratuito)", 0.0, "F"],
  ["Amazon SES", "40 000 correos/mes x $0.10 por 1 000", 4.0, "V"],
  ["Amazon SNS - SMS a El Salvador", "800 SMS/mes x $0.337 (cálculo manual: la calculadora no lo incluye; solo recordatorios críticos y alertas de resultados)", 269.6, "V"],
  ["Amazon Chime SDK (videoconsulta)", "8 000 consultas x 20 min x 2 asistentes = 320 000 min x $0.0017", 544.0, "V"],
  ["AWS Secrets Manager", "6 secretos + llamadas API (valor de la AWS Pricing Calculator)", 2.9, "F"],
  ["AWS KMS", "4 claves CMK x $1 + 1 M solicitudes x $0.03/10 k", 7.0, "F"],
  ["Amazon CloudWatch", "50 GB de logs, 20 alarmas y 1 dashboard (valor de la AWS Pricing Calculator)", 42.23, "V"],
  ["AWS CloudTrail", "1 trail de administración (gratis) + eventos de datos S3 (5 M x $0.10/100 k)", 5.0, "V"],
  ["Transferencia de datos a Internet (API)", "200 GB - 100 GB gratis = 100 GB x $0.09 (cálculo manual: la calculadora no lo incluye)", 9.0, "V"],
];
const round2 = (n) => Math.round(n * 100) / 100;
const total = round2(costs.reduce((a, c) => a + c[2], 0));
const totalF = round2(costs.filter((c) => c[3] === "F").reduce((a, c) => a + c[2], 0));
const totalV = round2(costs.filter((c) => c[3] === "V").reduce((a, c) => a + c[2], 0));
const reserve = round2(total * 0.10);
const money = (n) => "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
console.log("Total mensual estimado:", total, " fijo:", totalF, " variable:", totalV);

/* ------------------------------------------------------------------ */
/* Helpers de formato                                                  */
/* ------------------------------------------------------------------ */
const FONT = "Arial";
const NAVY = "232F3E", ORANGE = "FF9900", GRAY = "5F6B7A", LIGHT = "F2F4F7", HEAD_FILL = "232F3E";
const W = 9360; // ancho útil portrait (US Letter, márgenes 1")

function runs(text, base = {}) {
  // soporta **negrita** y `codigo`
  const out = [];
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`)/g).filter((s) => s !== "");
  for (const s of parts) {
    if (s.startsWith("**")) out.push(new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, ...base }));
    else if (s.startsWith("`")) out.push(new TextRun({ text: s.slice(1, -1), font: "Consolas", size: (base.size || 21) - 1, color: "B42318", ...base, bold: false }));
    else out.push(new TextRun({ text: s, font: FONT, ...base }));
  }
  return out;
}
const P = (text, o = {}) =>
  new Paragraph({
    children: runs(text, { size: 21, ...(o.run || {}) }),
    spacing: { after: o.after ?? 120, line: 288 },
    alignment: o.align || AlignmentType.JUSTIFIED,
    ...(o.para || {}),
  });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })], pageBreakBefore: false });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })] });
const B = (text, level = 0) =>
  new Paragraph({ numbering: { reference: "bullets", level }, children: runs(text, { size: 21 }), spacing: { after: 60, line: 276 } });
const N = (text, inst = 1) =>
  new Paragraph({ numbering: { reference: "numbers", level: 0, instance: inst }, children: runs(text, { size: 21 }), spacing: { after: 60, line: 276 } });
const PB = () => new Paragraph({ children: [new PageBreak()] });
const CAP = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 200 }, children: [new TextRun({ text: t, italics: true, size: 18, color: GRAY, font: FONT })] });

const border = { style: BorderStyle.SINGLE, size: 4, color: "C9D1DA" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(text, w, o = {}) {
  const lines = Array.isArray(text) ? text : [text];
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    verticalAlign: o.v || "top",
    children: lines.map((l) => new Paragraph({
      alignment: o.align || AlignmentType.LEFT, spacing: { after: 30 },
      children: runs(l, { size: o.size || 18, color: o.color, bold: o.bold }),
    })),
  });
}
function table(headers, rows, widths, o = {}) {
  const sum = widths.reduce((a, b) => a + b, 0);
  const head = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headers.map((h, i) => cell(h, widths[i], { fill: HEAD_FILL, color: "FFFFFF", bold: true, size: 18, v: "center" })),
  });
  const body = rows.map((r, ri) => new TableRow({
    cantSplit: true,
    children: r.map((c, i) => cell(c, widths[i], {
      fill: o.rowFill ? o.rowFill(r, ri) : (ri % 2 ? LIGHT : undefined),
      align: o.aligns ? o.aligns[i] : undefined,
      bold: o.boldFirst && i === 0,
    })),
  }));
  return new Table({ width: { size: sum, type: WidthType.DXA }, columnWidths: widths, rows: [head, ...body] });
}
const spacer = (n = 120) => new Paragraph({ spacing: { after: n }, children: [] });

function codeBlock(lines) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({
      children: [new TableCell({
        borders, width: { size: W, type: WidthType.DXA },
        shading: { fill: "F6F8FA", type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: lines.map((l) => new Paragraph({ spacing: { after: 0, line: 240 }, children: [new TextRun({ text: l === "" ? " " : l, font: "Consolas", size: 16 })] })),
      })],
    })],
  });
}
const callout = (title, text, fill = "FFF4E0", bar = ORANGE) =>
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.SINGLE, size: 24, color: bar } },
      shading: { fill, type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 100, bottom: 100, left: 160, right: 140 },
      children: [
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: title, bold: true, size: 20, font: FONT })] }),
        ...(Array.isArray(text) ? text : [text]).map((t) => new Paragraph({ spacing: { after: 40 }, children: runs(t, { size: 19 }) })),
      ],
    })] })],
  });

/* ------------------------------------------------------------------ */
/* Encabezado y pie (sello de footer Hugo Jovel Web)                   */
/* ------------------------------------------------------------------ */
const makeHeader = () => new Header({
  children: [new Paragraph({
    style: "HFTxt",
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ORANGE, space: 4 } },
    children: [
      new TextRun({ text: "MediConnect — Reporte técnico de arquitectura AWS", size: 16, color: GRAY, font: FONT }),
      new TextRun({ text: "\tPágina ", size: 16, color: GRAY, font: FONT }),
      new SimpleField("PAGE", "1"),
    ],
  })],
});
const makeFooter = () => new Footer({
  children: [
    new Paragraph({ // línea 1: contexto / stack
      style: "HFTxt", alignment: AlignmentType.CENTER, spacing: { after: 0 },
      children: [new TextRun({ text: "Kodigo Full Stack Jr 36 · AWS · Next.js 16 · Laravel 12 · MySQL 8.4", size: 16, color: GRAY, font: FONT })],
    }),
    new Paragraph({ // línea 2: sello (año dinámico, nombre en negrita y enlazado)
      style: "HFTxt", alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "© ", size: 16, color: GRAY, font: FONT }),
        new SimpleField('DATE \\@ "yyyy"', String(new Date().getFullYear())),
        new TextRun({ text: " ", size: 16, font: FONT }),
        new ExternalHyperlink({
          link: "https://hugojovelweb.tech/",
          children: [new TextRun({ text: "Hugo Jovel Web", bold: true, size: 16, color: "0B5CAD", underline: {}, font: FONT })],
        }),
        new TextRun({ text: ". Todos los derechos reservados.", size: 16, color: GRAY, font: FONT }),
      ],
    }),
  ],
});

/* ------------------------------------------------------------------ */
/* CONTENIDO                                                           */
/* ------------------------------------------------------------------ */
const cover = [
  new Paragraph({ spacing: { before: 1400, after: 120 }, children: [new TextRun({ text: "☁  ACTIVIDAD EVALUADA · EVALUACIÓN GRUPAL", bold: true, size: 20, color: ORANGE, font: FONT })] }),
  new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "MediConnect", bold: true, size: 72, color: NAVY, font: FONT })] }),
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Plataforma Nacional de Telemedicina", size: 36, color: NAVY, font: FONT })] }),
  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ORANGE, space: 6 } }, spacing: { after: 300 }, children: [] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Reporte Técnico de Arquitectura en AWS", bold: true, size: 40, font: FONT })] }),
  new Paragraph({ spacing: { after: 700 }, children: [new TextRun({ text: "Implementar soluciones escalables en la nube seleccionando los servicios de AWS adecuados según las necesidades del negocio", italics: true, size: 24, color: GRAY, font: FONT })] }),
  table(["Dato", "Detalle"], [
    ["Bootcamp", "Kodigo Full Stack Jr 36"],
    ["Tipo de actividad", "Evaluación grupal — equipos de 3 a 5 estudiantes"],
    ["Integrantes del equipo", ["Ixcell Marcela Arévalo Mesa", "Hugo Ernesto Jovel Hernández"]],
    ["Región AWS propuesta", "US East (N. Virginia) — us-east-1, 2 Zonas de Disponibilidad"],
    ["Fecha límite de entrega", "Domingo 15 de noviembre de 2026, 23:59"],
    ["Versión del documento", "1.1 — noviembre de 2026 (estimación de costos validada en AWS Pricing Calculator)"],
  ], [2600, 6760], { boldFirst: true }),
  spacer(300),
  callout("Contenido", [
    "1. Resumen Ejecutivo   ·   2. Análisis de Requerimientos   ·   3. Propuesta de Arquitectura AWS",
    "4. Diagrama de Arquitectura   ·   5. Justificación de Decisiones Técnicas   ·   6. Estrategia de Seguridad",
    "7. Estimación de Costos   ·   8. Conclusiones   ·   Anexos A–C",
  ], "EEF3F8", NAVY),
  PB(),
];

/* ---------------------------- 1. RESUMEN ---------------------------- */
const s1 = [
  H1("1. Resumen Ejecutivo"),
  H2("1.1 Descripción del problema"),
  P("MediConnect es una startup salvadoreña que digitaliza el acceso a la salud mediante videoconsultas, historial clínico digital y recetas electrónicas. Hoy toda la plataforma —portal del paciente, dashboard del médico, API Laravel, base de datos MySQL, archivos médicos y generación de reportes— corre en **un único servidor físico** de un datacenter local. Ese diseño tiene cuatro debilidades estructurales:"),
  B("**Punto único de falla:** un solo servidor sin redundancia; si cae, cae toda la atención médica."),
  B("**Sin elasticidad:** los lunes por la mañana la demanda llega a 5 veces el tráfico diario promedio y el servidor colapsa; el negocio exige soportar hasta 10×."),
  B("**Riesgo para datos sensibles:** los archivos médicos viven en el disco local (sin replicación ni durabilidad garantizada) y la base de datos comparte máquina con la aplicación."),
  B("**Procesos que bloquean el servicio:** el reporte semanal tarda entre 8 y 15 minutos y degrada la API mientras corre."),
  H2("1.2 Síntesis de la solución en AWS"),
  P("Se propone migrar a una arquitectura **contenerizada, multi-AZ y desacoplada** en la región us-east-1, organizada en una VPC con tres capas de subredes (pública, privada de aplicación y privada de datos):"),
  B("**Entrada y rendimiento:** Route 53 + CloudFront (con AWS WAF y certificados ACM) sirven el contenido Next.js desde el borde y protegen el origen; un Application Load Balancer reparte el tráfico entre dos zonas de disponibilidad."),
  B("**Cómputo elástico:** Next.js (portal y dashboard) y Laravel (API) se empaquetan como contenedores en **Amazon ECS con AWS Fargate**, con Auto Scaling por métricas y escalado programado para los lunes."),
  B("**Datos protegidos:** **Amazon RDS for MySQL Multi-AZ** en subredes aisladas (sin ruta a Internet), cifrado con KMS, con réplica de lectura para reportes; **ElastiCache Redis** para sesiones y caché."),
  B("**Archivos médicos:** **Amazon S3** con 99.999999999 % de durabilidad, versionado, cifrado SSE-KMS y replicación entre regiones."),
  B("**Procesamiento desacoplado:** **SQS + EventBridge Scheduler + workers Fargate** ejecutan reportes y notificaciones fuera de la API; **SES** (correo) y **SNS** (SMS) envían los mensajes."),
  B("**Seguridad y gobierno:** IAM con mínimo privilegio, **Secrets Manager** para credenciales, CloudTrail, CloudWatch y AWS Budgets."),
  H2("1.3 Beneficios principales"),
  table(["Beneficio", "Antes (servidor físico)", "Después (AWS)"], [
    ["Escalabilidad", "Colapsa con 5× el tráfico", "Auto Scaling hasta 10× sin degradación; refuerzo programado antes del pico del lunes"],
    ["Disponibilidad", "Servidor único, sin SLA", "Multi-AZ en todas las capas críticas; objetivo ≥ 99.9 % (≈ 43 min de indisponibilidad máx. al mes)"],
    ["Seguridad de datos", "BD y archivos en el mismo host", "BD sin acceso desde Internet, cifrado en reposo y tránsito, credenciales fuera del código"],
    ["Rendimiento", "Latencia variable por distancia y carga", "CDN + caché + balanceo: meta < 2 s en todo el país"],
    ["Reportes", "8–15 min bloqueando la API", "Worker asíncrono sobre réplica de lectura; la API no se ve afectada"],
    ["Durabilidad", "Disco local", "S3 (11 nueves) + versionado + replicación a otra región"],
    ["Costos", "CAPEX fijo e impredecible", "OPEX estimado ≈ " + money(total) + "/mes con presupuestos y alertas (sección 7)"],
  ], [1900, 2700, 4760], { boldFirst: true }),
  spacer(160),
];

/* ---------------------------- 2. REQUERIMIENTOS --------------------- */
const s2 = [
  H1("2. Análisis de Requerimientos"),
  H2("2.1 Requerimientos del negocio vs. desafío técnico"),
  P("La siguiente tabla traduce cada requerimiento obligatorio del CEO en un desafío técnico concreto, la decisión de diseño que lo resuelve y la métrica con la que se verificará."),
  table(["#", "Requerimiento del negocio", "Desafío técnico", "Decisión / servicios AWS", "Métrica de aceptación"], [
    ["R1", "Escalabilidad: hasta 10× el tráfico normal los lunes", "Cómputo y BD deben crecer en minutos; el servidor actual es de capacidad fija", "ECS Fargate + Application Auto Scaling (seguimiento de objetivo + escalado programado); ALB; ElastiCache; réplica de lectura", "p95 de la API < 500 ms con 10× carga en prueba de estrés"],
    ["R2", "Alta disponibilidad ≥ 99.9 %", "Eliminar puntos únicos de falla; recuperación automática ante fallo de AZ", "2 AZ en ALB, NAT, Fargate, RDS Multi-AZ, ElastiCache Multi-AZ; health checks de Route 53 y ALB", "Disponibilidad mensual ≥ 99.9 % (≤ 43 min de caída)"],
    ["R3", "Seguridad de datos médicos: BD nunca accesible desde Internet; credenciales fuera del código", "Aislar la capa de datos; gestionar secretos sin versionarlos", "Subredes privadas de datos sin ruta a IGW/NAT; Security Groups encadenados; Secrets Manager + IAM roles; KMS", "Escaneo externo: 0 puertos de BD abiertos; 0 secretos en el repositorio"],
    ["R4", "Carga rápida (< 2 s) en todo el país", "Reducir latencia hacia usuarios dispersos geográficamente", "CloudFront (CDN) + caché Redis + compresión; imágenes optimizadas por Next.js", "LCP < 2 s medido con Lighthouse/CrUX desde varias ciudades"],
    ["R5", "Reportes desacoplados de la API", "Un trabajo de 8–15 min no debe consumir recursos de la API", "EventBridge Scheduler → SQS → worker Fargate independiente que lee de la réplica; resultado en S3", "Latencia de la API sin variación durante la generación del reporte"],
    ["R6", "Archivos médicos nunca se pierden", "Durabilidad extrema, protección contra borrado/corrupción", "S3 (11 nueves), versionado, replicación entre regiones (CRR), Object Lock en modo gobernanza; RDS con backups y PITR", "RPO ≈ 0 en archivos; RPO ≤ 5 min en BD; RTO ≤ 1 h"],
    ["R7", "Costos predecibles", "Evitar sorpresas por consumo variable", "Modelo de costos por servicio, etiquetas de costo, AWS Budgets con alertas, Savings Plans/instancias reservadas para la carga estable, límite de gasto SMS", "Desviación mensual real vs. estimada < 10 %"],
  ], [500, 1650, 1900, 3010, 2300], { boldFirst: true }),
  spacer(),
  H2("2.2 Componentes críticos"),
  P("No todos los componentes tienen el mismo impacto si fallan. Se clasifican por criticidad para decidir dónde invertir redundancia:"),
  table(["Componente", "Criticidad", "Por qué es crítico", "Riesgo actual → mitigación en AWS"], [
    ["Base de datos MySQL", "Muy alta", "Contiene datos clínicos, recetas y pagos; es el único estado persistente transaccional", "Disco y servidor únicos → RDS Multi-AZ, backups automáticos, PITR, cifrado KMS, aislamiento de red"],
    ["Almacenamiento de archivos", "Muy alta", "Imágenes médicas y resultados que no pueden perderse ni exponerse", "Disco local → S3 versionado + CRR + URLs prefirmadas de corta vida"],
    ["API Laravel", "Alta", "Centraliza autenticación, agendamiento, historiales y notificaciones; todos los clientes dependen de ella", "Proceso único → varias tareas Fargate en 2 AZ detrás de ALB con Auto Scaling"],
    ["Portal del paciente (Next.js)", "Alta", "Es la puerta de entrada de pacientes y de las videoconsultas", "Saturación el lunes → CloudFront + Auto Scaling + escalado programado"],
    ["Dashboard del médico (Next.js)", "Alta", "Sin él los médicos no pueden atender ni emitir recetas", "Igual que el portal; acceso restringido adicional (WAF + MFA)"],
    ["Notificaciones (correo/SMS)", "Media", "Confirmaciones y alertas de resultados; su falla no detiene la atención pero afecta la experiencia", "Envío síncrono → cola SQS con reintentos y DLQ"],
    ["Procesamiento de reportes", "Media", "Genera carga intensiva que degrada el resto del sistema", "Bloquea el servidor → worker aislado sobre réplica de lectura"],
  ], [2100, 1000, 3100, 3160], { boldFirst: true }),
  spacer(),
  H2("2.3 Supuestos de dimensionamiento"),
  P("Para poder estimar capacidad y costos se adoptan los siguientes supuestos (el caso de estudio no entrega cifras absolutas; deben ajustarse con datos reales de MediConnect):"),
  B("50 000 usuarios activos al mes; 8 000 videoconsultas mensuales de 20 minutos promedio."),
  B("El tráfico del lunes por la mañana es 5× el promedio diario observado hoy; la arquitectura se dimensiona para **10×** (factor de seguridad de 2)."),
  B("1 TB de archivos médicos al inicio (imágenes, laboratorios y PDFs) con crecimiento mensual del 5 %."),
  B("Base de datos de ~200 GB con picos de lectura que se descargan a una réplica y a Redis."),
  B("40 000 correos y 800 SMS al mes (los SMS se reservan para recordatorios críticos por su costo en El Salvador)."),
  B("Zona horaria de operación: America/El_Salvador (UTC−6, sin horario de verano), relevante para el escalado programado."),
  PB(),
];

/* ---------------------------- 3. ARQUITECTURA ----------------------- */
const svc = (name, role, config, req) => [name, role, config, req];
const s3sec_raw = [
  H1("3. Propuesta de Arquitectura AWS"),
  P("La solución se despliega en **us-east-1 (N. Virginia)** por su cobertura completa de servicios, menor precio y buena latencia hacia Centroamérica a través de CloudFront. Se utilizan **dos Zonas de Disponibilidad** (us-east-1a y us-east-1b) para balancear disponibilidad y costo. A continuación se describe cada servicio, por qué se eligió y qué requerimiento satisface."),
  H2("3.1 Red y entrada de tráfico"),
  table(["Servicio", "Rol en MediConnect y justificación técnica", "Configuración propuesta", "Requisito"], [
    svc("Amazon VPC", "Red privada lógica que aísla la plataforma. Permite separar por capas y aplicar defensa en profundidad con Security Groups y NACL.", "10.0.0.0/16; 2 AZ; 6 subredes (2 públicas, 2 privadas de app, 2 privadas de datos); VPC Endpoint de S3 (gateway) y Flow Logs", "R2, R3"),
    svc("Amazon Route 53", "DNS administrado con alta disponibilidad (SLA 100 %) y health checks; enruta el dominio a CloudFront y la API al ALB.", "Zona pública mediconnect.sv; registros ALIAS a CloudFront y ALB; health check de /health", "R2, R4"),
    svc("Amazon CloudFront", "CDN que sirve contenido estático y páginas Next.js desde ubicaciones de borde cercanas al usuario y absorbe picos del lunes sin tocar el origen.", "Orígenes: ALB (SSR) y S3 (estáticos); compresión Brotli/Gzip; HTTPS obligatorio; caché por tipo de contenido", "R1, R4"),
    svc("AWS WAF + Shield Standard", "Filtra tráfico malicioso (inyección SQL, XSS, bots, abuso) antes del origen; Shield Standard mitiga DDoS L3/L4 sin costo adicional.", "Web ACL en CloudFront y en el ALB: reglas administradas Common, Known Bad Inputs, SQLi, IP Reputation + límite de tasa por IP", "R3, R2"),
    svc("AWS Certificate Manager", "Certificados TLS gratuitos con renovación automática; evita expiraciones que tumban el servicio.", "Certificados para mediconnect.sv y api.mediconnect.sv; TLS 1.2+", "R3"),
    svc("Application Load Balancer", "Balanceo L7 entre tareas de varias AZ, health checks, enrutamiento por ruta/host (portal, dashboard, API) y terminación TLS.", "En subredes públicas; listeners 443; reglas por host/ruta; SG solo admite 443; deregistration delay de 30 s", "R1, R2"),
    svc("NAT Gateway (x2)", "Permite que los contenedores privados salgan a Internet (APIs externas, pagos) sin exponerlos; uno por AZ para no depender de una sola zona.", "Una por subred pública; el tráfico a S3 usa el VPC Endpoint para ahorrar costo de procesamiento", "R2, R3"),
  ], [1700, 3300, 3160, 1200], { boldFirst: true }),
  spacer(),
  H2("3.2 Cómputo y escalamiento"),
  table(["Servicio", "Rol en MediConnect y justificación técnica", "Configuración propuesta", "Requisito"], [
    svc("Amazon ECS + AWS Fargate", "Ejecuta contenedores sin administrar servidores (sin parches de SO ni capacidad ociosa). Es la forma más directa de llevar Next.js y Laravel a producción y escalar por demanda.", "4 servicios: portal (0.5 vCPU/1 GB), dashboard (0.5/1), API Laravel (1 vCPU/2 GB) y worker; tareas repartidas en 2 AZ; Fargate Spot opcional para workers", "R1, R2, R5"),
    svc("Application Auto Scaling", "Ajusta el número de tareas automáticamente y permite adelantar capacidad antes del pico conocido del lunes.", "Seguimiento de objetivo: CPU 60 % y RequestCountPerTarget; escalado programado lunes 05:30–12:00 (UTC−6): API mín. 10; máximo 20 tareas (10× el valor base)", "R1"),
    svc("Amazon ECR", "Registro privado de imágenes Docker con escaneo de vulnerabilidades y cifrado; integra con IAM.", "Un repositorio por servicio; escaneo al subir; política de ciclo de vida (últimas 20 imágenes)", "R3"),
    svc("Amazon ElastiCache for Redis", "Caché y almacén de sesiones compartido: permite contenedores sin estado y reduce lecturas a MySQL en los picos.", "2 nodos cache.t4g.small, Multi-AZ con failover, cifrado en tránsito y reposo, en subred de datos", "R1, R4"),
  ], [1700, 3300, 3160, 1200], { boldFirst: true }),
  spacer(),
  H2("3.3 Datos y archivos"),
  table(["Servicio", "Rol en MediConnect y justificación técnica", "Configuración propuesta", "Requisito"], [
    svc("Amazon RDS for MySQL 8.4 (Multi-AZ)", "Mantiene el motor MySQL actual (migración con mínimo cambio de código) como servicio administrado: failover automático, backups, parches y cifrado. Se elige 8.4 LTS porque el soporte estándar de 8.0 terminó el 31-jul-2026 y su Extended Support se cobra aparte.", "db.m6g.large Multi-AZ, 200 GB gp3 autoescalable, subred privada de datos, PubliclyAccessible = false, cifrado KMS, backups 14 días + PITR, protección contra borrado", "R2, R3, R6"),
    svc("Réplica de lectura de RDS", "Atiende los reportes y consultas pesadas sin afectar a la instancia primaria.", "db.m6g.large Single-AZ; el worker de reportes usa su endpoint", "R5, R1"),
    svc("Amazon S3", "Almacén de objetos con durabilidad de 99.999999999 %, ideal para imágenes, laboratorios y recetas PDF; elimina el disco local y escala sin límites.", "Buckets separados (médicos, reportes, estáticos); Block Public Access; SSE-KMS; versionado; CRR a us-east-2; Object Lock (gobernanza); ciclo de vida a S3 Standard-IA a los 90 días; acceso solo por URL prefirmada de corta vida", "R6, R3"),
  ], [1700, 3300, 3160, 1200], { boldFirst: true }),
  spacer(),
  H2("3.4 Procesamiento asíncrono, notificaciones y videoconsulta"),
  table(["Servicio", "Rol en MediConnect y justificación técnica", "Configuración propuesta", "Requisito"], [
    svc("Amazon EventBridge Scheduler", "Dispara el reporte semanal sin servidores cron propios ni depender de la API.", "Programación semanal (domingo 23:00 UTC−6) que publica un mensaje en SQS", "R5"),
    svc("Amazon SQS (+ DLQ)", "Cola que desacopla productores y consumidores: absorbe ráfagas y reintenta; los mensajes fallidos van a una cola de mensajes muertos para análisis.", "Colas reports y notifications; visibility timeout 30 min en reports; maxReceiveCount 3 → DLQ; cifrado SSE", "R5, R1"),
    svc("Workers en ECS Fargate", "Procesos de larga duración (8–15 min) que Lambda (límite 15 min) no soportaría con holgura. Son independientes de la API.", "notifications-worker (mín. 1) y reports-worker (mín. 0, escala por profundidad de cola, 2 vCPU/8 GB)", "R5"),
    svc("Amazon SES", "Correo transaccional barato y con alta entregabilidad (confirmaciones, recordatorios, recetas).", "Identidad de dominio verificada con DKIM/SPF/DMARC; salida del sandbox; tasas de rebote monitoreadas", "Notificaciones"),
    svc("Amazon SNS (SMS)", "Envío de SMS a pacientes en El Salvador para recordatorios críticos.", "Mensajes transaccionales; límite mensual de gasto SMS configurado; opt-out gestionado", "Notificaciones, R7"),
    svc("Amazon Chime SDK", "Videoconsultas WebRTC administradas, integrables en Next.js sin operar servidores multimedia; se paga por minuto-asistente.", "Reuniones creadas por la API Laravel con control de asistentes; sin grabación por defecto", "Videoconsulta"),
  ], [1700, 3300, 3160, 1200], { boldFirst: true }),
  spacer(),
  H2("3.5 Seguridad, observabilidad y gobierno"),
  table(["Servicio", "Rol en MediConnect y justificación técnica", "Configuración propuesta", "Requisito"], [
    svc("AWS IAM / IAM Identity Center", "Control de acceso con mínimo privilegio para personas y workloads; roles temporales en vez de claves de larga duración.", "Grupos por función con MFA; roles de tarea ECS; rol OIDC para CI/CD (ver sección 6)", "R3"),
    svc("AWS Secrets Manager", "Guarda credenciales de BD y claves de terceros con rotación automática; evita contraseñas en el código o en la imagen.", "Secretos inyectados en el contenedor al arrancar; rotación cada 30 días de la contraseña de RDS", "R3"),
    svc("AWS KMS", "Claves administradas por el cliente para cifrar RDS, S3, SQS, logs y secretos con trazabilidad de uso.", "CMK por dominio de datos (BD, archivos, mensajería) con rotación anual", "R3, R6"),
    svc("Amazon CloudWatch", "Logs, métricas, alarmas y dashboards; dispara el escalado y alerta al equipo.", "Alarmas: 5xx del ALB, latencia, CPU de RDS, profundidad de SQS, DLQ > 0; retención de logs 90 días", "R2"),
    svc("AWS CloudTrail", "Auditoría de todas las llamadas a la API de AWS y de accesos a objetos S3 sensibles.", "Trail multi-región hacia un bucket de logs inmutable", "R3"),
    svc("AWS Budgets + Cost Explorer", "Control de gasto y alertas anticipadas; sustenta la previsibilidad de costos.", "Presupuesto mensual con alertas al 50/80/100 %; etiquetas Proyecto/Ambiente/Componente", "R7"),
  ], [1700, 3300, 3160, 1200], { boldFirst: true }),
  spacer(),
  H2("3.6 Cómo se cumple cada requerimiento"),
  B("**Escalabilidad (R1):** el cómputo es horizontal y sin estado (sesiones en Redis, archivos en S3). Los lunes, un escalado programado adelanta capacidad a las 05:30 y el Auto Scaling reacciona a CPU y solicitudes por tarea hasta 20 tareas de API (10× la base). CloudFront absorbe estáticos y Redis/réplica protegen a MySQL."),
  B("**Alta disponibilidad (R2):** ALB, NAT, Fargate, RDS y ElastiCache operan en 2 AZ. Si una AZ falla, el ALB deja de enviarle tráfico, ECS relanza tareas en la otra y RDS conmuta a la standby en 60–120 s. Para la ruta crítica de la API (Route 53 × ALB × ECS × RDS Multi-AZ) los SLA publicados componen ≈ 99.93 %, por encima del 99.9 % exigido (los SLA son un piso contractual; la disponibilidad real suele ser mayor)."),
  B("**Seguridad (R3):** la BD vive en subredes sin ruta a Internet y su Security Group solo acepta el puerto 3306 desde los Security Groups de la API y los workers; las credenciales se leen de Secrets Manager en tiempo de ejecución."),
  B("**Tiempo de carga (R4):** CloudFront, compresión, caché Redis y balanceo de carga mantienen el tiempo de respuesta por debajo de 2 s; se valida con Lighthouse y pruebas desde distintos departamentos."),
  B("**Reportes desacoplados (R5):** el reporte corre en un worker independiente, alimentado por SQS y leyendo de la réplica; la API nunca ejecuta la tarea."),
  B("**Durabilidad (R6):** S3 + versionado + replicación entre regiones + Object Lock; RDS con backups automáticos y restauración a un punto en el tiempo."),
  B("**Costos predecibles (R7):** tabla de costos por servicio, presupuestos con alertas y compromisos de uso en la carga estable (sección 7)."),
  PB(),
];

const s3sec = s3sec_raw.slice(0, -1); // sin salto de página final (el cambio de sección ya paginó)
/* ---------------------------- 4. DIAGRAMA --------------------------- */
const img = fs.readFileSync("MediConnect_Diagrama_Arquitectura_AWS.png");
const s4 = [
  H1("4. Diagrama de Arquitectura"),
  P("La Figura 1 muestra la arquitectura completa: VPC con subredes públicas y privadas en dos zonas de disponibilidad, los flujos de tráfico numerados desde el usuario hasta la base de datos y los servicios administrados que la acompañan. El archivo editable **MediConnect_Arquitectura_AWS.drawio** (draw.io con la librería oficial de íconos AWS) se entrega junto con este reporte."),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
    children: [new ImageRun({ type: "png", data: img, transformation: { width: 790, height: 497 },
      altText: { title: "Diagrama de arquitectura AWS", description: "Arquitectura AWS de MediConnect con VPC, subredes públicas y privadas, ECS Fargate, RDS Multi-AZ y S3", name: "diagrama" } })],
  }),
  CAP("Figura 1. Arquitectura AWS de MediConnect (us-east-1, 2 AZ). Para máxima resolución, abrir el archivo .drawio y exportar en PNG/PDF a 300 dpi."),
  H2("4.1 Flujos de tráfico"),
  table(["Paso", "Origen → Destino", "Protocolo / puerto", "Descripción"], [
    ["1", "Usuario → Route 53 → CloudFront", "HTTPS 443", "El navegador resuelve el dominio; AWS WAF inspecciona la solicitud en el borde; el contenido en caché se responde sin llegar al origen."],
    ["2", "CloudFront → Internet Gateway → ALB", "HTTPS 443", "Las solicitudes no cacheables viajan al origen; el Security Group del ALB acepta solo 443."],
    ["3", "ALB → ECS Fargate (portal, dashboard, API)", "HTTP/HTTPS 3000 / 8080", "El ALB enruta por host/ruta hacia tareas en subredes privadas de aplicación de ambas AZ."],
    ["4", "Next.js (SSR) → API Laravel", "HTTPS interno", "El portal y el dashboard consumen la API REST; la sesión y el caché viven en ElastiCache."],
    ["5", "API Laravel → RDS MySQL", "MySQL 3306 (TLS)", "Solo el SG de la API y de los workers puede abrir el puerto; la BD no tiene IP pública ni ruta a Internet."],
    ["6", "API → S3 (URL prefirmada)", "HTTPS vía VPC Endpoint", "El cliente sube/descarga archivos médicos directamente a S3 con una URL temporal; el tráfico no sale a Internet desde la VPC."],
    ["7", "API / EventBridge → SQS → Workers → SES, SNS, S3", "HTTPS (API AWS)", "Los reportes y notificaciones se procesan fuera de la API; el worker lee de la réplica y guarda el resultado en S3."],
  ], [750, 2550, 1800, 4260], { boldFirst: true }),
  spacer(),
  H2("4.2 Plan de direccionamiento y Security Groups"),
  table(["Capa", "Subredes (AZ-a / AZ-b)", "Recursos", "Salida a Internet"], [
    ["Pública", "10.0.0.0/24 / 10.0.1.0/24", "ALB, NAT Gateway", "Directa por Internet Gateway"],
    ["Privada de aplicación", "10.0.10.0/24 / 10.0.11.0/24", "Tareas ECS Fargate (portal, dashboard, API, workers)", "Solo egreso vía NAT Gateway; S3 por VPC Endpoint"],
    ["Privada de datos", "10.0.20.0/24 / 10.0.21.0/24", "RDS MySQL (primaria, standby, réplica), ElastiCache", "Ninguna (sin ruta a IGW ni NAT)"],
  ], [1700, 2500, 3160, 2000], { boldFirst: true }),
  spacer(),
  table(["Security Group", "Entrada permitida", "Origen"], [
    ["sg-alb", "TCP 443", "0.0.0.0/0 (o prefijo administrado de CloudFront para el dominio web)"],
    ["sg-ecs-apps", "TCP 3000 / 8080", "sg-alb"],
    ["sg-ecs-workers", "Ninguna", "—"],
    ["sg-rds", "TCP 3306", "sg-ecs-apps, sg-ecs-workers"],
    ["sg-redis", "TCP 6379", "sg-ecs-apps"],
  ], [2300, 2300, 4760], { boldFirst: true }),
  PB(),
];

/* ---------------------------- 5. JUSTIFICACIÓN ---------------------- */
const s5 = [
  H1("5. Justificación de Decisiones Técnicas"),
  P("Cada servicio seleccionado se contrastó con al menos una alternativa razonable. Se justifican **16 servicios AWS distintos** (el mínimo solicitado es 6), con criterio de requerimiento cubierto, esfuerzo operativo y costo."),
  table(["#", "Servicio seleccionado", "Alternativa(s) descartada(s)", "Motivo de la decisión"], [
    ["1", "Amazon ECS + Fargate", "EC2 con Auto Scaling; Amazon EKS; AWS Lambda", "Fargate elimina la administración de servidores y parches y escala por tarea. EKS añade ~$73/mes de plano de control y complejidad innecesaria para 4 servicios. Lambda no es ideal para Next.js SSR con estado de conexión y limita a 15 min los reportes."],
    ["2", "Amazon RDS for MySQL 8.4 Multi-AZ", "MySQL en EC2; Amazon Aurora MySQL; DynamoDB", "Conserva el motor actual (migración con cambios mínimos en Laravel) y delega failover, parches y backups. Aurora es una evolución válida (ver conclusiones) pero eleva costo y complejidad inicial; DynamoDB exigiría reescribir el modelo relacional clínico."],
    ["3", "Amazon S3", "EBS/EFS en EC2; almacenamiento en el servidor", "Durabilidad de 11 nueves, versionado, replicación y Object Lock a costo por GB bajo; EBS/EFS dependen de instancias y no ofrecen la misma durabilidad."],
    ["4", "Amazon CloudFront", "Servir directamente desde el ALB; otro CDN externo", "Reduce latencia y descarga el origen en los picos del lunes; se integra nativamente con WAF, ACM y S3. Un CDN externo agregaría otro proveedor y otra consola."],
    ["5", "Application Load Balancer", "Network Load Balancer; Nginx propio en EC2; API Gateway", "ALB entrega enrutamiento L7, health checks y TLS administrados. NLB opera en L4; Nginx propio reintroduce un punto único de falla; API Gateway cobra por solicitud y no es necesario para este patrón."],
    ["6", "Amazon SQS + EventBridge Scheduler", "Amazon MQ (RabbitMQ); Kafka (MSK); cron en un servidor", "SQS es serverless, económico y nativo en Laravel (driver sqs). MSK/MQ implican clústeres que operar y costos fijos; un cron en un servidor recrea el problema de dependencia de una sola máquina."],
    ["7", "AWS Secrets Manager", "Variables en .env o en la imagen; SSM Parameter Store", "Rotación automática y auditoría de acceso para credenciales de BD. Parameter Store (SecureString) se usa para configuración no secreta porque no tiene costo por parámetro estándar."],
    ["8", "AWS WAF + Shield Standard", "Solo Security Groups; firewall propio", "Protege la capa 7 (SQLi, XSS, bots) con reglas administradas y sin mantenimiento; Shield Standard viene incluido."],
    ["9", "Amazon ElastiCache for Redis", "Sesiones en MySQL; Memcached", "Saca las sesiones de la BD y reduce lecturas; Redis ofrece persistencia opcional, Multi-AZ y estructuras útiles para colas y límites de tasa."],
    ["10", "Amazon SES", "Servidor SMTP propio; proveedores externos", "$0.10 por 1 000 correos con DKIM/SPF y alta entregabilidad, sin servidores que mantener."],
    ["11", "Amazon SNS (SMS)", "Proveedor SMS externo (p. ej. Twilio)", "Integración nativa con IAM y facturación única; se limita el gasto con un tope mensual. Un proveedor externo podría ser más barato por mensaje y debe evaluarse si el volumen crece."],
    ["12", "Amazon Chime SDK", "Servidor WebRTC propio; Zoom/Meet embebido", "Servicio administrado y pagado por uso (minuto-asistente), sin operar servidores multimedia."],
    ["13", "AWS IAM + KMS", "Credenciales compartidas; cifrado a nivel de aplicación", "Mínimo privilegio con roles temporales y claves con auditoría; es la base de los controles de seguridad médica."],
    ["14", "Amazon CloudWatch + CloudTrail", "Herramientas de monitoreo propias", "Integración nativa con el Auto Scaling, alarmas y auditoría de API sin desplegar agentes adicionales."],
    ["15", "NAT Gateway + VPC Endpoint S3", "NAT instance en EC2; IP pública en cada contenedor", "Administrado y de alta disponibilidad; las IP públicas en contenedores ampliarían la superficie de ataque."],
    ["16", "Amazon Route 53 + ACM", "DNS del registrador; certificados manuales", "Health checks, ALIAS y SLA del 100 %; certificados gratuitos con renovación automática."],
  ], [450, 2300, 2500, 4110], { boldFirst: true }),
  spacer(),
  callout("Decisión de región", [
    "Se eligió **us-east-1** por precio, madurez y disponibilidad de todos los servicios. La región Mexico (Central) reduciría la latencia hacia Centroamérica pero tiene un catálogo de servicios y precios distintos que deben validarse; se deja como alternativa para una fase posterior si las pruebas de latencia lo justifican.",
  ], "EEF3F8", NAVY),
  PB(),
];

/* ---------------------------- 6. SEGURIDAD -------------------------- */
const s6 = [
  H1("6. Estrategia de Seguridad"),
  H2("6.1 Protección de datos sensibles"),
  P("Los datos clínicos, recetas y pagos se protegen con defensa en profundidad, desde la red hasta el objeto almacenado:"),
  table(["Capa", "Control", "Detalle"], [
    ["Red", "Aislamiento por subredes y Security Groups", "La BD está en subredes sin ruta a Internet y con PubliclyAccessible = false; el único origen permitido en 3306 son los SG de la API y los workers."],
    ["Perímetro", "AWS WAF + Shield + CloudFront", "Reglas administradas OWASP, límite de tasa y bloqueo de IP maliciosas; HTTPS obligatorio con TLS 1.2+."],
    ["Cifrado en reposo", "AWS KMS (CMK)", "RDS, snapshots, ElastiCache, S3, SQS y logs cifrados; rotación anual de claves."],
    ["Cifrado en tránsito", "TLS en todos los saltos", "require_secure_transport = ON en MySQL; Redis con TLS; políticas S3 que niegan solicitudes sin TLS (aws:SecureTransport = false)."],
    ["Archivos médicos", "S3 privado + URLs prefirmadas", "Block Public Access activo; descarga mediante URL de 5 minutos generada por la API tras validar permisos del usuario; versionado + Object Lock."],
    ["Aplicación", "Autenticación y autorización", "Laravel con tokens de vida corta (JWT/Sanctum), MFA para médicos y administradores, control por rol y registro de accesos a historiales."],
    ["Detección y auditoría", "CloudTrail, CloudWatch, VPC Flow Logs", "Registro inmutable de acciones; alarmas ante accesos anómalos; GuardDuty recomendado como mejora."],
    ["Respaldo", "RDS automated backups + PITR; S3 CRR", "Backups 14 días; restauración a un punto en el tiempo; copia de archivos en us-east-2."],
  ], [1700, 2500, 5160], { boldFirst: true }),
  callout("Nota de cumplimiento", "Los datos de salud son de máxima sensibilidad. Se toman los controles de HIPAA como referencia de buenas prácticas (servicios elegibles, cifrado, auditoría) y debe validarse con asesoría legal el cumplimiento de la normativa salvadoreña de protección de datos personales y de salud aplicable.", "FDECEC", "DD344C"),
  spacer(),
  H2("6.2 Configuración IAM: usuarios, roles y políticas"),
  H3("Usuarios y grupos (personas)"),
  B("**No se usa el usuario root** para operar: se protege con MFA y se guarda como credencial de emergencia."),
  B("Acceso humano mediante **IAM Identity Center** (SSO) con MFA obligatorio; no se crean usuarios IAM con claves de acceso permanentes."),
  table(["Grupo", "Quiénes", "Permisos"], [
    ["Administradores-Cloud", "Líder técnico / DevOps (2 personas)", "AdministratorAccess solo en cuenta de producción con MFA y sesión de 1 hora; el resto del tiempo usan el rol de desarrollo"],
    ["Desarrolladores", "Equipo Full Stack", "PowerUser sin IAM; lectura de CloudWatch y ECS; sin acceso a datos de producción"],
    ["Operaciones-Soporte", "Soporte 24/7", "Lectura de logs/métricas y reinicio de servicios ECS"],
    ["Finanzas", "Equipo financiero", "Solo lectura de Billing, Cost Explorer y Budgets"],
    ["Auditoría", "Oficial de cumplimiento", "SecurityAudit + lectura de CloudTrail"],
  ], [2300, 2700, 4360], { boldFirst: true }),
  H3("Roles para servicios (workloads)"),
  table(["Rol", "Quién lo asume", "Permisos principales (mínimo privilegio)"], [
    ["mediconnect-ecs-task-execution", "ECS (arranque de la tarea)", "Descargar imágenes de ECR, escribir en CloudWatch Logs y leer SOLO los secretos listados en Secrets Manager"],
    ["mediconnect-api-task-role", "Contenedor API Laravel", "S3 (GetObject/PutObject en el prefijo de archivos médicos), SQS (SendMessage), SES (SendEmail), SNS (Publish SMS), KMS (Decrypt/GenerateDataKey con la CMK asignada)"],
    ["mediconnect-worker-task-role", "Contenedores worker", "SQS (Receive/Delete), S3 (PutObject en bucket de reportes), SES/SNS, KMS de la CMK correspondiente"],
    ["mediconnect-cicd-deploy", "GitHub Actions por OIDC", "Subir imágenes a ECR y actualizar servicios ECS; sin claves de larga duración"],
    ["mediconnect-scheduler-role", "EventBridge Scheduler", "sqs:SendMessage solo a la cola reports"],
  ], [2700, 2200, 4460], { boldFirst: true }),
  H3("Ejemplo de política de mínimo privilegio — rol de tarea de la API"),
  codeBlock([
    "{",
    '  "Version": "2012-10-17",',
    '  "Statement": [',
    "    {",
    '      "Sid": "ArchivosMedicosSoloPrefijo",',
    '      "Effect": "Allow",',
    '      "Action": ["s3:GetObject", "s3:PutObject"],',
    '      "Resource": "arn:aws:s3:::mediconnect-prod-archivos/pacientes/*"',
    "    },",
    "    {",
    '      "Sid": "EncolarTrabajos",',
    '      "Effect": "Allow",',
    '      "Action": "sqs:SendMessage",',
    '      "Resource": [',
    '        "arn:aws:sqs:us-east-1:ACCOUNT_ID:mediconnect-notifications",',
    '        "arn:aws:sqs:us-east-1:ACCOUNT_ID:mediconnect-reports"',
    "      ]",
    "    },",
    "    {",
    '      "Sid": "UsarSoloSuClave",',
    '      "Effect": "Allow",',
    '      "Action": ["kms:Decrypt", "kms:GenerateDataKey"],',
    '      "Resource": "arn:aws:kms:us-east-1:ACCOUNT_ID:key/KEY_ID_ARCHIVOS"',
    "    }",
    "  ]",
    "}",
  ]),
  spacer(),
  H3("Política del bucket: negar tráfico sin TLS y accesos públicos"),
  codeBlock([
    "{",
    '  "Version": "2012-10-17",',
    '  "Statement": [{',
    '    "Sid": "DenegarSinTLS",',
    '    "Effect": "Deny",',
    '    "Principal": "*",',
    '    "Action": "s3:*",',
    '    "Resource": ["arn:aws:s3:::mediconnect-prod-archivos",',
    '                 "arn:aws:s3:::mediconnect-prod-archivos/*"],',
    '    "Condition": { "Bool": { "aws:SecureTransport": "false" } }',
    "  }]",
    "}",
  ]),
  spacer(),
  H2("6.3 Gestión de credenciales y variables de entorno"),
  P("Regla de oro: **ningún secreto vive en el repositorio, en la imagen Docker ni en el código**. El flujo es el siguiente:"),
  N("Las credenciales (usuario/contraseña de RDS, APP_KEY de Laravel, claves de terceros) se crean en **AWS Secrets Manager**, cifradas con KMS. La contraseña de RDS rota automáticamente cada 30 días.", 1),
  N("La **definición de tarea de ECS** referencia el ARN del secreto en el bloque `secrets`; ECS lo inyecta como variable de entorno solo en memoria al arrancar el contenedor.", 1),
  N("La configuración **no secreta** (APP_ENV, nombre de buckets, URL de colas) va en `environment` o en SSM Parameter Store.", 1),
  N("En Next.js solo se exponen al navegador las variables con prefijo `NEXT_PUBLIC_` (URLs públicas); cualquier secreto se lee únicamente en el servidor.", 1),
  N("El repositorio contiene un `.env.example` sin valores reales y `.env` está en `.gitignore`; se activa escaneo de secretos (GitHub secret scanning) y se rotan de inmediato los que se filtren.", 1),
  N("El despliegue usa **OIDC** entre GitHub Actions y AWS (rol temporal), no claves de acceso guardadas como secretos del repositorio.", 1),
  H3("Fragmento de definición de tarea ECS (API Laravel)"),
  codeBlock([
    '"containerDefinitions": [{',
    '  "name": "api-laravel",',
    '  "image": "ACCOUNT_ID.dkr.ecr.us-east-1.amazonaws.com/mediconnect-api:1.0.0",',
    '  "environment": [',
    '    { "name": "APP_ENV", "value": "production" },',
    '    { "name": "DB_HOST", "value": "mediconnect-prod.xxxxxx.us-east-1.rds.amazonaws.com" },',
    '    { "name": "QUEUE_CONNECTION", "value": "sqs" },',
    '    { "name": "AWS_BUCKET", "value": "mediconnect-prod-archivos" }',
    "  ],",
    '  "secrets": [',
    '    { "name": "DB_PASSWORD", "valueFrom": "arn:aws:secretsmanager:us-east-1:ACCOUNT_ID:secret:mediconnect/prod/db:password::" },',
    '    { "name": "APP_KEY",     "valueFrom": "arn:aws:secretsmanager:us-east-1:ACCOUNT_ID:secret:mediconnect/prod/app:APP_KEY::" }',
    "  ]",
    "}]",
  ]),
  spacer(),
  H2("6.4 Continuidad y respuesta a incidentes"),
  B("**RPO/RTO:** RPO ≤ 5 min en BD (PITR) y ≈ 0 en archivos; RTO ≤ 1 h con failover Multi-AZ y restauración automatizada."),
  B("**Alarmas críticas** hacia SNS/correo del equipo: errores 5xx, DLQ con mensajes, CPU/almacenamiento de RDS, intentos de acceso denegados y gasto anómalo."),
  B("**Pruebas periódicas:** simulacro de failover de RDS, restauración de un backup y prueba de carga 10× antes de cada temporada alta."),
  PB(),
];

/* ---------------------------- 7. COSTOS ----------------------------- */
const costRows = costs.map((c) => [c[0], c[1], money(round2(c[2])), c[3] === "F" ? "Base" : "Variable"]);
costRows.push(["Subtotal base (capacidad provisionada)", "", money(totalF), ""]);
costRows.push(["Subtotal variable (depende del uso)", "", money(totalV), ""]);
costRows.push(["TOTAL MENSUAL ESTIMADO", "Escenario base + refuerzo de los lunes", money(total), ""]);
costRows.push(["Reserva de contingencia (10 %)", "Para picos, crecimiento y errores de estimación", money(reserve), ""]);
costRows.push(["PRESUPUESTO MENSUAL RECOMENDADO", "Total + contingencia (valor del AWS Budget)", money(round2(total + reserve)), ""]);
const s7 = [
  H1("7. Estimación de Costos"),
  P("La estimación usa precios de lista de **us-east-1** (USD, 730 horas/mes, bajo demanda) y los supuestos de la sección 2.3. La estimación se reprodujo en la **AWS Pricing Calculator** (enlace al final de esta sección). Los valores de la calculadora son los oficiales; los SMS y la transferencia de datos se calcularon a mano porque la calculadora no los incluye."),
  H2("7.1 Costos por servicio"),
  table(["Servicio", "Configuración y cálculo", "Costo mensual", "Tipo"], costRows, [2500, 4500, 1260, 1100], {
    boldFirst: true, aligns: [undefined, undefined, AlignmentType.RIGHT, undefined],
    rowFill: (r, i) => (String(r[0]).startsWith("TOTAL") || String(r[0]).startsWith("PRESUPUESTO") ? "FFE7B8" : (String(r[0]).startsWith("Subtotal") || String(r[0]).startsWith("Reserva") ? "E8EEF5" : (i % 2 ? LIGHT : undefined))),
  }),
  CAP("Tabla 7.1. Estimación mensual de costos. Valores tomados de la AWS Pricing Calculator; SMS y transferencia de datos calculados manualmente."),
  H2("7.2 Lectura de los resultados"),
  B("El costo total estimado es de **" + money(total) + " al mes** (≈ " + money(round2(total * 12)) + " al año) más una reserva del 10 % para fijar el presupuesto en **" + money(round2(total + reserve)) + "**."),
  B("Los mayores impulsores de costo son **Amazon Chime SDK**, **RDS** (primaria + réplica) y los **SMS a El Salvador** (≈ $0.34 por mensaje): reducir SMS a lo realmente crítico y usar correo/notificaciones push es la palanca más efectiva."),
  B("El cómputo Fargate es económico (≈ " + money(round2(fargate)) + ") porque el refuerzo del lunes solo se paga las horas que se usa; si se llega al límite de 10× durante 6 h cada lunes, el costo adicional es menor a $15 al mes."),
  B("El 10× no se compra de forma permanente: se paga por uso, algo imposible con el servidor físico actual."),
  H2("7.3 Controles para costos predecibles (R7)"),
  B("**AWS Budgets** con alertas al 50 %, 80 % y 100 % del presupuesto y detección de anomalías de costo."),
  B("**Etiquetas obligatorias** (Proyecto, Ambiente, Componente) y informes por etiqueta en Cost Explorer."),
  B("**Compromisos de uso:** Compute Savings Plan (1 año) para Fargate y Reserved Instances/Database Savings Plans de 1 año para RDS, aplicados solo a la carga estable una vez medido el consumo real."),
  B("**Límite de gasto de SMS** en SNS, S3 Lifecycle (Standard-IA a los 90 días), retención de logs a 90 días y VPC Endpoints para reducir el procesamiento de NAT."),
  B("**Ambientes no productivos** (staging) apagados fuera del horario laboral."),
  callout("Enlace de la estimación (AWS Pricing Calculator)", [
    "Estimación compartida en la AWS Pricing Calculator (región US East - N. Virginia). Total oficial de la calculadora más el cálculo manual de SMS y transferencia de datos: " + money(total) + " al mes.",
  ], "E8F5E9", "2E7D32"),
  new Paragraph({ spacing: { before: 120, after: 120 }, children: [
    new TextRun({ text: "Enlace compartido: ", bold: true, size: 20, font: FONT }),
    new ExternalHyperlink({ link: "https://calculator.aws/#/estimate?id=bc30fded3d33d4d662b6d66ff9c0b72faf09d0e7",
      children: [new TextRun({ text: "https://calculator.aws/#/estimate?id=bc30fded3d33d4d662b6d66ff9c0b72faf09d0e7", size: 18, color: "0B5CAD", underline: {}, font: FONT })] }),
  ] }),
  spacer(160),
];

/* ---------------------------- 8. CONCLUSIONES ----------------------- */
const s8 = [
  H1("8. Conclusiones"),
  H2("8.1 Reflexión sobre la selección de servicios cloud"),
  P("El caso MediConnect demuestra que elegir servicios en la nube no consiste en usar la mayor cantidad de herramientas, sino en **traducir cada requerimiento del negocio en una decisión técnica verificable**. La combinación de servicios administrados (RDS, S3, SQS, Fargate) traslada a AWS las tareas de parcheo, replicación y failover, y permite al equipo concentrarse en el producto médico."),
  P("Tres principios guiaron el diseño: **(1)** eliminar los puntos únicos de falla con arquitectura multi-AZ; **(2)** desacoplar lo pesado (reportes, notificaciones) de lo crítico (la API); y **(3)** proteger el dato clínico con capas independientes de red, identidad y cifrado, de modo que un solo error no exponga información de pacientes."),
  H2("8.2 Lecciones aprendidas"),
  N("**El requisito manda el servicio.** Cada elección se justificó contra un requerimiento y una alternativa descartada; así se evitó sobredimensionar (por ejemplo, EKS o Aurora desde el día uno).", 2),
  N("**Escalar no es solo agregar servidores:** hace falta que la aplicación sea sin estado (sesiones en Redis, archivos en S3) para que el escalado horizontal funcione.", 2),
  N("**La seguridad es parte del diseño inicial:** aislar la BD en subredes sin salida y gestionar secretos fuera del código es más barato que remediar una filtración.", 2),
  N("**Los costos ocultos importan:** NAT Gateway, transferencia de datos, SMS y videoconsulta pueden pesar más que el cómputo; modelarlos desde el inicio hace el gasto predecible.", 2),
  N("**Los precios y versiones cambian:** por ejemplo, el fin del soporte estándar de MySQL 8.0 en RDS obliga a elegir 8.4; la estimación debe validarse con la Calculadora vigente antes de comprometer presupuesto.", 2),
  N("**Probar antes de confiar:** el diseño de 10× debe validarse con pruebas de carga y simulacros de failover, no asumirse.", 2),
  H2("8.3 Evolución recomendada"),
  B("**Infraestructura como código** (Terraform o CloudFormation/CDK) para reproducir ambientes y auditar cambios."),
  B("Evaluar **Amazon Aurora MySQL** si la carga de lectura/escritura crece, y una **región secundaria** para recuperación ante desastres."),
  B("Incorporar **GuardDuty, Security Hub y AWS Backup** para detección de amenazas y respaldos centralizados."),
  B("Medir latencia real desde varios departamentos y evaluar la región de México (Central) si fuera necesario."),
  spacer(160),
];

/* ---------------------------- ANEXOS -------------------------------- */
const sA = [
  H1("Anexo A. Cómo se generó la estimación en AWS Pricing Calculator"),
  N("Ingresar a https://calculator.aws → **Create estimate** → seleccionar la región **US East (N. Virginia)**.", 3),
  N("Agregar cada servicio con **Add service** y capturar la configuración de la tabla 7.1 (más abajo, la lista de entradas).", 3),
  N("Nombrar la estimación «MediConnect – Producción» y usar **Share** para obtener el enlace público (o exportar a PDF/CSV).", 3),
  N("El enlace de la estimación compartida está en la sección 7. Su total, más el cálculo manual de SMS y transferencia de datos (que la calculadora no incluye), coincide con la tabla 7.1.", 3),
  table(["Servicio en la Calculadora", "Valores a ingresar"], [
    ["Amazon Route 53", "1 hosted zone; 10 millones de consultas estándar"],
    ["Amazon CloudFront", "Data transfer out 800 GB; 20 millones de solicitudes HTTPS"],
    ["AWS WAF", "2 Web ACL; 10 reglas; 30 millones de solicitudes"],
    ["Elastic Load Balancing", "1 Application Load Balancer; 4 LCU promedio"],
    ["NAT Gateway", "2 gateways; 300 GB procesados por mes"],
    ["AWS Fargate (ECS)", "Linux/x86: API 2 tareas 1 vCPU/2 GB 24×7; portal y dashboard 2 tareas cada uno 0.5 vCPU/1 GB 24×7; worker 1 tarea 0.5 vCPU/1 GB; 17 tareas extra por 26 h/mes; reportes 2 vCPU/8 GB, 15 min/semana"],
    ["Amazon ECR", "20 GB de almacenamiento"],
    ["Amazon RDS for MySQL (primaria)", "db.m6g.large, Multi-AZ, 200 GB gp3, 100 GB de backup adicional, bajo demanda"],
    ["Amazon RDS for MySQL (réplica)", "db.m6g.large, Single-AZ, 200 GB gp3"],
    ["Amazon ElastiCache", "Redis (OSS), 2 nodos cache.t4g.small, Multi-AZ"],
    ["Amazon S3", "Standard 1 TB (us-east-1) + 1 TB (us-east-2) + 50 GB reportes; solicitudes y replicación estimadas"],
    ["Amazon SQS", "10 millones de solicitudes estándar"],
    ["Amazon SES", "40 000 correos/mes"],
    ["Amazon SNS (SMS)", "800 SMS a El Salvador/mes — NO se ingresa en la calculadora; se calcula a mano (800 x $0.337)"],
    ["Amazon Chime SDK", "320 000 minutos-asistente de WebRTC"],
    ["Transferencia de datos a Internet", "NO se ingresa en la calculadora; cálculo manual: (200 GB - 100 GB gratis) x $0.09"],
    ["AWS Secrets Manager / KMS", "6 secretos; 4 claves CMK; 1 millón de solicitudes"],
    ["Amazon CloudWatch / CloudTrail", "50 GB de logs; 20 alarmas; 1 dashboard; 5 millones de eventos de datos S3"],
  ], [3000, 6360], { boldFirst: true }),
  spacer(),
  H1("Anexo B. Cómo editar el diagrama en draw.io"),
  N("Abrir **MediConnect_Arquitectura_AWS.drawio** en https://app.diagrams.net o en draw.io Desktop (File → Open).")
  ,
  N("Si algún ícono apareciera en blanco, activar la librería oficial: **More Shapes → AWS 19** (o «AWS Architecture 2025») y reemplazar el ícono arrastrando el correcto.", 4),
  N("Ajustar etiquetas y flujos; exportar con **File → Export as → PNG** (zoom 200 %, fondo blanco) y reemplazar la imagen del Word (clic derecho → Cambiar imagen).", 4),
  spacer(),
  H1("Anexo C. Fuentes de precios consultadas (noviembre de 2026)"),
  B("AWS Fargate (us-east-1, Linux/x86): $0.04048 por vCPU-hora y $0.004445 por GB-hora — aws.amazon.com/fargate/pricing."),
  B("Application Load Balancer $0.0225/h y NAT Gateway $0.045/h — aws.amazon.com/elasticloadbalancing/pricing y aws.amazon.com/vpc/pricing."),
  B("Amazon RDS for MySQL — valor de la instancia y almacenamiento tomado de la AWS Pricing Calculator (aws.amazon.com/rds/mysql/pricing)."),
  B("Amazon SNS / AWS End User Messaging: SMS a El Salvador ≈ $0.337 por mensaje — aws.amazon.com/sns/sms-pricing."),
  B("Amazon Chime SDK: $0.0017 por minuto-asistente — aws.amazon.com/chime/chime-sdk/pricing."),
  B("Otros precios de lista (S3, CloudFront, WAF, SQS, SES, KMS, Secrets Manager, CloudWatch) tomados de las páginas oficiales de cada servicio; verificar vigencia en la Calculadora."),
];

/* ------------------------------------------------------------------ */
/* Documento                                                           */
/* ------------------------------------------------------------------ */
const pageProps = { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440, header: 500, footer: 450 } } };
const landscapeProps = { page: { size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE }, margin: { top: 900, right: 900, bottom: 900, left: 900, header: 350, footer: 350 } } };

const doc = new Document({
  creator: "Equipo MediConnect — Kodigo Full Stack Jr 36",
  title: "MediConnect — Reporte Técnico de Arquitectura AWS",
  description: "Actividad evaluada: soluciones escalables en la nube con AWS",
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "HFTxt", name: "HF Text", basedOn: "Normal", run: { font: FONT, size: 16, color: GRAY } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 34, bold: true, color: NAVY },
        paragraph: { spacing: { before: 120, after: 200 }, keepNext: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ORANGE, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: "1D4E89" },
        paragraph: { spacing: { before: 240, after: 120 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: NAVY },
        paragraph: { spacing: { before: 180, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } } ] },
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } } ] },
    ],
  },
  sections: [
    { properties: pageProps, headers: { default: makeHeader() }, footers: { default: makeFooter() }, children: [...cover, ...s1, ...s2, ...s3sec] },
    { properties: landscapeProps, headers: { default: makeHeader() }, footers: { default: makeFooter() }, children: [...s4.slice(0, 4)] },
    { properties: pageProps, headers: { default: makeHeader() }, footers: { default: makeFooter() }, children: [...s4.slice(4), ...s5, ...s6, ...s7, ...s8, ...sA] },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("MediConnect_Reporte_Tecnico_AWS.docx", buf);
  console.log("OK -> MediConnect_Reporte_Tecnico_AWS.docx");
});
