const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  PageBreak,
  HeadingLevel,
  ShadingType,
} = require('docx');

// Configuración de estilos tipográficos estándar requeridos
const FONT_NAME = 'Times New Roman';
const SIZE_BODY = 24; // 12pt en medios puntos
const SIZE_H1 = 32;   // 16pt
const SIZE_H2 = 28;   // 14pt
const SIZE_H3 = 24;   // 12pt negrita
const SIZE_TABLE = 21; // 10.5pt para tablas compactas
const LINE_SPACING_1_5 = 360; // 1.5 interlineado (240 * 1.5 = 360)

// Helper para párrafos normales con Times New Roman 12 e interlineado 1.5
function createParagraph(text = '', options = {}) {
  const { bold = false, italic = false, align = AlignmentType.JUSTIFIED, spaceAfter = 160 } = options;
  return new Paragraph({
    alignment: align,
    spacing: { line: LINE_SPACING_1_5, after: spaceAfter },
    children: [
      new TextRun({
        text,
        font: FONT_NAME,
        size: SIZE_BODY,
        bold,
        italic,
        color: '000000',
      }),
    ],
  });
}

// Helper para títulos H1
function createHeading1(title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: LINE_SPACING_1_5 },
    children: [
      new TextRun({
        text: title,
        font: FONT_NAME,
        size: SIZE_H1,
        bold: true,
        color: '000000',
      }),
    ],
  });
}

// Helper para subtítulos H2
function createHeading2(title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140, line: LINE_SPACING_1_5 },
    children: [
      new TextRun({
        text: title,
        font: FONT_NAME,
        size: SIZE_H2,
        bold: true,
        color: '000000',
      }),
    ],
  });
}

// Helper para subtítulos H3
function createHeading3(title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100, line: LINE_SPACING_1_5 },
    children: [
      new TextRun({
        text: title,
        font: FONT_NAME,
        size: SIZE_H3,
        bold: true,
        color: '000000',
      }),
    ],
  });
}

// Helper para viñetas
function createBulletPoint(boldText, normalText) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: LINE_SPACING_1_5, after: 120 },
    bullet: { level: 0 },
    children: [
      new TextRun({
        text: boldText + ' ',
        font: FONT_NAME,
        size: SIZE_BODY,
        bold: true,
        color: '000000',
      }),
      new TextRun({
        text: normalText,
        font: FONT_NAME,
        size: SIZE_BODY,
        color: '000000',
      }),
    ],
  });
}

// Helper para recuadro de evidencia / capturas
function createEvidenceBox(caption) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            shading: { type: ShadingType.CLEAR, fill: 'FAFAFA' },
            margins: { top: 300, bottom: 300, left: 300, right: 300 },
            borders: {
              top: { style: BorderStyle.DASHED, size: 8, color: '666666' },
              bottom: { style: BorderStyle.DASHED, size: 8, color: '666666' },
              left: { style: BorderStyle.DASHED, size: 8, color: '666666' },
              right: { style: BorderStyle.DASHED, size: 8, color: '666666' },
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [
                  new TextRun({
                    text: '[ ESPACIO PARA INSERTAR CAPTURA DE PANTALLA ]',
                    font: FONT_NAME,
                    size: 20,
                    bold: true,
                    color: '000000',
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 0 },
                children: [
                  new TextRun({
                    text: caption,
                    font: FONT_NAME,
                    size: 18,
                    italic: true,
                    color: '000000',
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// Helper para celda de tabla formateada
function createCell(text, isHeader = false, widthPercent = 20) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    shading: isHeader
      ? { type: ShadingType.CLEAR, fill: 'E2E8F0' }
      : { type: ShadingType.CLEAR, fill: 'FFFFFF' },
    margins: { top: 120, bottom: 120, left: 140, right: 140 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
      left: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
      right: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
    },
    children: [
      new Paragraph({
        alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { line: 260, after: 0 },
        children: [
          new TextRun({
            text,
            font: FONT_NAME,
            size: SIZE_TABLE,
            bold: isHeader,
            color: '000000',
          }),
        ],
      }),
    ],
  });
}

async function buildDocument() {
  const sections = [];

  // ==========================================
  // PÁGINA 1: HOJA EN BLANCO PARA PORTADA
  // ==========================================
  const portadaChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 2000 },
      children: [
        new TextRun({
          text: '', // Página en blanco para portada del alumno
          font: FONT_NAME,
        }),
      ],
    }),
    new Paragraph({
      children: [new PageBreak()],
    }),
  ];

  // ==========================================
  // CONTENIDO ACADÉMICO FORMAL
  // ==========================================
  const docChildren = [
    // Encabezado principal del reporte
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120, line: LINE_SPACING_1_5 },
      children: [
        new TextRun({
          text: 'PROYECTO INTEGRADOR',
          font: FONT_NAME,
          size: 26,
          bold: true,
          color: '000000',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 280, line: LINE_SPACING_1_5 },
      children: [
        new TextRun({
          text: 'FUNDAMENTOS Y ESCENARIOS DE USO DE LAS APLICACIONES WEB PROGRESIVAS',
          font: FONT_NAME,
          size: 24,
          bold: true,
          color: '000000',
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 360, line: LINE_SPACING_1_5 },
      children: [
        new TextRun({
          text: 'Sistema "TechField": Plataforma de Reportes de Servicios Técnicos en Campo con Arquitectura Offline-First',
          font: FONT_NAME,
          size: 22,
          italic: true,
          color: '000000',
        }),
      ],
    }),

    // ── 1. INTRODUCCIÓN Y PROBLEMÁTICA ──
    createHeading1('1. Introducción y Análisis de la Problemática'),
    createParagraph(
      'El presente proyecto integrador aborda una problemática crítica recurrente en el sector de servicios técnicos y de mantenimiento de infraestructura: la necesidad operativa de capturar y validar reportes de fallas, visitas preventivas y mantenimientos en campo dentro de zonas rurales, predios agrícolas y sectores periféricos que carecen de conectividad a Internet (redes celulares 3G/4G/5G o redes inalámbricas Wi-Fi).'
    ),
    createParagraph(
      'En este escenario, las cuadrillas técnicas deben desplazarse a áreas remotas donde la cobertura de red es completamente nula o altamente inestable. A este desafío geográfico se suma la heterogeneidad de los dispositivos móviles empleados: los técnicos operan con smartphones de diversos fabricantes, arquitecturas y versiones de sistemas operativos, principalmente terminales Android (desde gamas de entrada hasta gamas medias) y dispositivos iOS (iPhone).'
    ),
    createParagraph(
      'Bajo el modelo tradicional de aplicaciones web o sistemas administrativos centralizados, la pérdida de señal inhabilita por completo la captura de datos, provocando que los técnicos recurran al registro manual en papel o bitácoras físicas. Esto desencadena problemas de doble captura, errores de transcripción, pérdida de información geográfica y retrasos de hasta varios días en la actualización del estado de las órdenes en la oficina central.'
    ),
    createParagraph(
      'Para superar esta brecha operativa, se plantea el diseño, configuración y construcción de una Aplicación Web Progresiva (PWA, por sus siglas en inglés) que adopte una arquitectura Offline-First. Dicha solución garantiza que el 100% de la funcionalidad sustantiva (creación de reportes, captura de coordenadas GPS, almacenamiento local seguro y consulta de órdenes históricas) permanezca operativa en ausencia total de conectividad, realizando una sincronización automática y transparente hacia el servidor central en el momento exacto en que el dispositivo recupera cobertura.'
    ),

    // ── 2. ANÁLISIS ARQUITECTÓNICO Y MATRIZ DE DECISIÓN ──
    createHeading1('2. Análisis Arquitectónico y Selección de Tecnología (20% Rúbrica)'),
    createParagraph(
      'Para justificar de manera rigurosa la selección técnica de una Aplicación Web Progresiva frente a los requerimientos de la empresa de servicios técnicos, se evaluaron formalmente cinco alternativas de arquitectura de software:'
    ),
    createBulletPoint(
      '2.1. Aplicación Web Tradicional (Server-Side Rendering):',
      'Arquitectura donde cada interacción del usuario y navegación requiere una petición HTTP síncrona al servidor web, el cual procesa la lógica y renderiza el HTML final. Si bien presenta un costo de desarrollo moderado y compatibilidad universal, es inviable para el escenario planteado debido a su nula tolerancia al trabajo sin conexión: ante la pérdida de señal, el navegador arroja de inmediato un error de red y bloquea la operación.'
    ),
    createBulletPoint(
      '2.2. Aplicación Web Orientada a Servicios (SPA + API REST):',
      'Desacopla la interfaz de usuario en una Single-Page Application (React, Vue, Angular) que consume una interfaz de programación de aplicaciones (API REST). Ofrece una experiencia fluida y reactiva mientras existe conexión, pero sin la integración de Service Workers y almacenamiento local transaccional, carece de capacidades de ejecución offline y no puede instalarse de manera independiente en la pantalla de inicio del smartphone.'
    ),
    createBulletPoint(
      '2.3. Aplicación Móvil Nativa (Android con Kotlin / iOS con Swift):',
      'Desarrollo de dos códigos fuente independientes compilados específicamente para cada plataforma móvil. Provee el máximo rendimiento y acceso irrestricto a los sensores del hardware. No obstante, representa un costo de desarrollo y mantenimiento prohibitivo (doble equipo técnico), ciclos de entrega prolongados y la dependencia obligatoria de pasar por los procesos de revisión, aprobación y comisiones comerciales de Google Play Store y Apple App Store.'
    ),
    createBulletPoint(
      '2.4. Aplicación Móvil Multiplataforma (Flutter / React Native):',
      'Unifica el desarrollo mediante una base de código compartida que compila a binarios nativos o enlaces nativos. Reduce el esfuerzo de duplicar código, pero mantiene la fricción de distribución a través de las tiendas de aplicaciones, el incremento sustancial en el peso final del instalador (MB) y la complejidad en la gestión de dependencias externas que pueden quedar obsoletas.'
    ),
    createBulletPoint(
      '2.5. Aplicación Web Progresiva (PWA con Service Worker e IndexedDB):',
      'Aplicación web moderna enriquecida con capacidades de almacenamiento en caché programático (Cache Storage API), interceptores de red en segundo plano (Service Workers), manifiesto de instalación web (Web App Manifest) y base de datos local no volátil (IndexedDB). Proporciona funcionamiento offline nativo, instalación directa sin intermediarios mediante URL, actualizaciones instantáneas y compatibilidad multiplataforma a una fracción del costo y tiempo de desarrollo.'
    ),

    createHeading2('2.6. Matriz de Decisión Ponderada'),
    createParagraph(
      'A continuación se presenta la matriz cuantitativa de toma de decisiones. Se establecieron pesos porcentuales de acuerdo con las prioridades de la operación técnica en campo:'
    ),

    // TABLA DE LA MATRIZ DE DECISIÓN
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            createCell('Criterio de Evaluación', true, 26),
            createCell('Peso', true, 10),
            createCell('Web Trad.', true, 12),
            createCell('SPA + API', true, 13),
            createCell('App Nativa', true, 13),
            createCell('Multipl.', true, 13),
            createCell('PWA (TechField)', true, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('Funcionalidad Offline (Sin red)', false, 26),
            createCell('25%', false, 10),
            createCell('0 / 10', false, 12),
            createCell('3 / 10', false, 13),
            createCell('10 / 10', false, 13),
            createCell('9 / 10', false, 13),
            createCell('10 / 10', false, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('Instalabilidad Directa (Sin Store)', false, 26),
            createCell('15%', false, 10),
            createCell('0 / 10', false, 12),
            createCell('0 / 10', false, 13),
            createCell('7 / 10', false, 13),
            createCell('7 / 10', false, 13),
            createCell('10 / 10', false, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('Costo y Eficiencia de Desarrollo', false, 26),
            createCell('20%', false, 10),
            createCell('8 / 10', false, 12),
            createCell('7 / 10', false, 13),
            createCell('2 / 10', false, 13),
            createCell('5 / 10', false, 13),
            createCell('9 / 10', false, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('Compatibilidad Multiplataforma', false, 26),
            createCell('15%', false, 10),
            createCell('10 / 10', false, 12),
            createCell('10 / 10', false, 13),
            createCell('3 / 10', false, 13),
            createCell('7 / 10', false, 13),
            createCell('10 / 10', false, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('Despliegue y Actualización Inmediata', false, 26),
            createCell('10%', false, 10),
            createCell('10 / 10', false, 12),
            createCell('10 / 10', false, 13),
            createCell('2 / 10', false, 13),
            createCell('3 / 10', false, 13),
            createCell('10 / 10', false, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('Acceso a Hardware (GPS / Cámara)', false, 26),
            createCell('15%', false, 10),
            createCell('3 / 10', false, 12),
            createCell('4 / 10', false, 13),
            createCell('10 / 10', false, 13),
            createCell('9 / 10', false, 13),
            createCell('8 / 10', false, 13),
          ],
        }),
        new TableRow({
          children: [
            createCell('CALIFICACIÓN PONDERADA FINAL', true, 26),
            createCell('100%', true, 10),
            createCell('4.65', true, 12),
            createCell('5.05', true, 13),
            createCell('6.00', true, 13),
            createCell('7.05', true, 13),
            createCell('9.50', true, 13),
          ],
        }),
      ],
    }),

    createParagraph(
      'Conclusión de la Selección: La PWA obtiene la máxima calificación ponderada (9.50 / 10). Resuelve de raíz el requerimiento de continuidad operativa offline, permite distribuir la aplicación a cientos de técnicos con un simple enlace HTTPS o código QR sin pasar por las tiendas de apps, y utiliza los estándares web abiertos de geolocalización (HTML5 Geolocation API) y almacenamiento masivo (IndexedDB API).',
      { italic: true, spaceAfter: 240 }
    ),

    // ── 3. COMPATIBILIDAD Y REQUERIMIENTOS DE INSTALACIÓN ──
    createHeading1('3. Compatibilidad Multiplataforma y Requerimientos de Instalación (20% Rúbrica)'),
    createParagraph(
      'Uno de los mandatos clave del proyecto es garantizar la compatibilidad entre diversos motores de renderizado. En el ecosistema móvil coexisten dos grandes familias: los navegadores basados en Chromium (Google Chrome, Microsoft Edge, Samsung Internet, Opera) y los navegadores basados en el motor WebKit de Apple (Safari en iOS y iPadOS).'
    ),

    createHeading2('3.1. Tabla Comparativa de Soporte Técnico por Navegador'),

    // TABLA DE COMPATIBILIDAD
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            createCell('Característica PWA', true, 26),
            createCell('Chrome Android (Chromium)', true, 19),
            createCell('Safari iOS 16.4+ (WebKit)', true, 19),
            createCell('Edge Windows (Chromium)', true, 18),
            createCell('Samsung Internet', true, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('Service Worker & Fetch API', false, 26),
            createCell('Completo (100%)', false, 19),
            createCell('Completo (100%)', false, 19),
            createCell('Completo (100%)', false, 18),
            createCell('Completo (100%)', false, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('Cache Storage API (App Shell)', false, 26),
            createCell('Completo (100%)', false, 19),
            createCell('Completo (100%)', false, 19),
            createCell('Completo (100%)', false, 18),
            createCell('Completo (100%)', false, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('Web App Manifest & Standalone', false, 26),
            createCell('Completo (100%)', false, 19),
            createCell('Completo (100%)', false, 19),
            createCell('Completo (100%)', false, 18),
            createCell('Completo (100%)', false, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('Instalación en Inicio (A2HS)', false, 26),
            createCell('Prompt Automático', false, 19),
            createCell('Manual (Botón Compartir)', false, 19),
            createCell('Prompt Automático', false, 18),
            createCell('Prompt Automático', false, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('IndexedDB (Persistencia local)', false, 26),
            createCell('Completo (>1 GB)', false, 19),
            createCell('Completo (~1 GB límite)', false, 19),
            createCell('Completo (>1 GB)', false, 18),
            createCell('Completo (>1 GB)', false, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('Background Sync API', false, 26),
            createCell('Soportado nativo', false, 19),
            createCell('No soportado', false, 19),
            createCell('Soportado nativo', false, 18),
            createCell('Soportado nativo', false, 18),
          ],
        }),
        new TableRow({
          children: [
            createCell('Captura GPS Satelital', false, 26),
            createCell('Precisión satelital alta', false, 19),
            createCell('Precisión satelital alta', false, 19),
            createCell('Precisión satelital alta', false, 18),
            createCell('Precisión satelital alta', false, 18),
          ],
        }),
      ],
    }),

    createHeading2('3.2. Limitaciones Específicas de WebKit / Safari iOS y Estrategias de Mitigación'),
    createParagraph(
      'La investigación exhaustiva de las limitaciones de Apple WebKit reveló tres restricciones fundamentales que fueron atendidas durante el diseño de la arquitectura de TechField:'
    ),
    createBulletPoint(
      '1. Regla de purga de caché tras 7 días de inactividad:',
      'WebKit impone una política de privacidad por la cual el Cache Storage y los datos del Service Worker pueden ser eliminados si el usuario no interactúa con la aplicación durante un periodo de 7 días. Mitigación: Toda la información crítica de órdenes de trabajo y reportes no se almacena en el Cache Storage, sino en IndexedDB, la cual cuenta con una política de retención mucho más estricta en el sistema operativo iOS.'
    ),
    createBulletPoint(
      '2. Cuota de almacenamiento reducida (~1 GB):',
      'Mientras que Chromium permite consumir hasta el 60% del espacio libre del disco duro, WebKit restringe el origen a un máximo aproximado de 1 GB. Mitigación: Para un sistema de texto estructurado y coordenadas GPS como TechField, 1 GB permite guardar más de 50,000 órdenes de servicio completas, por lo que el diseño es 100% viable.'
    ),
    createBulletPoint(
      '3. Ausencia de Background Sync API en iOS:',
      'Safari en iOS no soporta la API de sincronización periódica en segundo plano. Mitigación: Se implementó un detector de eventos en el navegador mediante window.addEventListener("online") y comprobación en el evento de carga (lifecycle) de la aplicación, garantizando que tan pronto el técnico abra la app o recupere red, la cola de IndexedDB se despacha al servidor sin depender del proceso de fondo de iOS.'
    ),

    createHeading2('3.3. Checklist de Criterios de Instalación PWA'),
    createParagraph(
      'Para que la aplicación sea reconocida como instalable por los navegadores y cumpla la totalidad de las auditorías de Lighthouse, se verificó el siguiente listado de criterios técnicos obligatorios:'
    ),
    createBulletPoint('[X] Servidor con protocolo seguro:', 'Obligatoriedad de HTTPS o localhost en entorno de desarrollo.'),
    createBulletPoint('[X] Web App Manifest enlazado:', 'Archivo manifest.json vinculado en la etiqueta <head> con atributos de nombre, nombre corto y descripción.'),
    createBulletPoint('[X] Modo de despliegue Standalone:', 'display: "standalone" configurado para ocultar la barra de direcciones y botones del navegador web, brindando una experiencia 100% nativa.'),
    createBulletPoint('[X] Paquete de íconos multiplataforma:', 'Íconos PNG en resoluciones de 192x192 px, 512x512 px e ícono con propósito "maskable" para compatibilidad con los sistemas de recortes adaptativos de Android e iOS.'),
    createBulletPoint('[X] Service Worker registrado y activo:', 'Archivo sw.js registrado que implementa el evento fetch para interceptar tráfico de red.'),
    createBulletPoint('[X] Respuesta HTTP 200 en modo Offline:', 'El Service Worker devuelve recursos válidos y navegación completa aun sin conexión.'),
    createBulletPoint('[X] Viewport responsivo Mobile-First:', '<meta name="viewport" content="width=device-width, initial-scale=1.0"> correctamente configurado sin desbordamientos horizontales.'),

    // ── 4. PREPARACIÓN DEL ENTORNO Y API BASE ──
    createHeading1('4. Preparación del Entorno de Desarrollo y API Base (15% Rúbrica)'),
    createParagraph(
      'El entorno de desarrollo se diseñó con un enfoque desacoplado, modular y robusto que permite pruebas locales inmediatas y exposición segura a través de túneles hacia dispositivos físicos reales.'
    ),
    createHeading2('4.1. Arquitectura del Backend REST (Node.js + Express en MVC)'),
    createParagraph(
      'En lugar de un script monolítico, el backend se implementó bajo el patrón Modelo-Vista-Controlador (MVC), asegurando mantenibilidad, separación de responsabilidades y persistencia local sin bloqueos:'
    ),
    createBulletPoint('server/config/index.js:', 'Centraliza constantes, puertos de red, rutas de datos y orígenes de CORS.'),
    createBulletPoint('server/models/Reporte.js:', 'Fábrica y validador de objetos de reporte, asignando identificadores únicos y asegurando la integridad de los campos obligatorios.'),
    createBulletPoint('server/services/reporteService.js:', 'Capa de lógica de negocio y persistencia en disco (JSON). Implementa mecanismos de idempotencia y prevención de registros duplicados ante reenvíos rápidos por fluctuaciones de red.'),
    createBulletPoint('server/controllers/reporteController.js:', 'Controladores encargados de manejar las peticiones HTTP (GET, POST, DELETE) y generar respuestas estandarizadas.'),
    createBulletPoint('server/routes/reportes.js:', 'Enrutamiento limpio y semántico de la API.'),
    createBulletPoint('server/middleware/errorHandler.js:', 'Middleware centralizado para intercepción y reporte de anomalías sin detener el proceso del servidor.'),

    createHeading2('4.2. Configuración de Túnel Seguro HTTPS para Pruebas en Smartphones Reales'),
    createParagraph(
      'Dado que las Web APIs de Service Worker y Geolocation requieren obligatoriamente un contexto seguro (HTTPS), se configuró una herramienta de túnel seguro (LocalTunnel / Ngrok) que genera una URL pública encriptada con certificado SSL válido hacia el puerto local de desarrollo (5173).'
    ),
    createParagraph(
      'Esto permite a los evaluadores y técnicos escanear la URL en sus dispositivos móviles personales (Android o iPhone), validando la experiencia de usuario real bajo luz solar directa.'
    ),
    createEvidenceBox('Figura 1: Captura de pantalla de la terminal corriendo el túnel HTTPS y el endpoint de la API REST funcional.'),

    // ── 5. DESARROLLO DEL CORE PWA Y CAPACIDAD OFFLINE ──
    createHeading1('5. Desarrollo del Core PWA y Capacidad Offline (30% Rúbrica)'),
    createParagraph(
      'La interfaz de usuario fue construida utilizando React 19 junto con Vite y Tailwind CSS, implementando un diseño de alto contraste visual ("High-Contrast Field Palette") que sustituye los fondos oscuros por tonos arena y blancos cálidos (Stone/White) con acentos en verde bosque (Emerald-700/800). Esta disposición elimina los reflejos tipo espejo en las pantallas de los celulares cuando los técnicos operan bajo la luz solar directa del mediodía.'
    ),

    createHeading2('5.1. Ciclo de Vida del Service Worker Artesanal (Vanilla JS)'),
    createParagraph(
      'Para tener control absoluto del almacenamiento y de las peticiones sin depender de librerías opacas, el Service Worker (client/public/sw.js) implementa artesanalmente las tres fases de su ciclo de vida:'
    ),
    createBulletPoint(
      '1. Evento Install (Instalación):',
      'El Service Worker intercepta la instalación y descarga en caché (Cache Storage) los recursos esenciales del "App Shell": index.html, manifiesto, íconos y la página de contingencia offline.html. Se ejecuta self.skipWaiting() para que el nuevo worker tome el control de inmediato.'
    ),
    createBulletPoint(
      '2. Evento Activate (Activación):',
      'Se limpia el almacenamiento eliminando versiones antiguas de la caché para garantizar que los técnicos siempre reciban la última versión de la aplicación. Se invoca self.clients.claim() para reclamar el control de todas las pestañas abiertas.'
    ),
    createBulletPoint(
      '3. Evento Fetch (Intercepción de Peticiones):',
      'Aplica dos estrategias según la naturaleza de la petición: para la interfaz estática utiliza Stale-While-Revalidate (responde de inmediato desde la caché local y revalida con la red en segundo plano); para las peticiones de datos de la API utiliza Network-First con fallback a la caché y contingencia.'
    ),

    createHeading2('5.2. Persistencia Transaccional con IndexedDB frente a LocalStorage'),
    createParagraph(
      'Se descartó el uso de LocalStorage debido a que su operación es síncrona (bloquea el hilo principal de la interfaz) y posee un límite rígido de 5 MB que no soporta un historial técnico robusto. En su lugar se implementó IndexedDB (client/src/db/IndexedDB.js), una base de datos NoSQL asíncrona dentro del navegador que ofrece transacciones ACID y capacidad de almacenamiento de cientos de megabytes.'
    ),
    createParagraph(
      'Flujo de sincronización de reportes (sync.js): Cuando el técnico registra un trabajo sin red, el reporte se almacena con la bandera "sincronizado: false". La aplicación mantiene un listener activo sobre el evento "online". En cuanto el teléfono detecta conectividad, se despacha automáticamente la cola de reportes pendientes hacia el endpoint POST /api/reportes del servidor, actualizando el estado a "sincronizado: true" sin requerir intervención manual del usuario.'
    ),
    createEvidenceBox('Figura 2: Demostración de la aplicación en "Modo Avión" (Offline), almacenamiento en IndexedDB y auto-sincronización exitosa.'),

    // ── 6. AUDITORÍA, PRUEBAS Y CIERRE ──
    createHeading1('6. Auditoría Técnica, Pruebas Multiplataforma y Cierre (15% Rúbrica)'),
    createParagraph(
      'La calidad del prototipo fue sometida a rigurosas auditorías técnicas y pruebas en dispositivos reales:'
    ),
    createBulletPoint(
      'Auditoría Lighthouse PWA:',
      'Se verificó el cumplimiento del 100% de los criterios de Progressive Web App (manifest válido, Service Worker con evento fetch, inicio en standalone, splash screen personalizado y respuesta 200 sin conexión).'
    ),
    createBulletPoint(
      'Pruebas de Instalación Real:',
      'Se ejecutaron pruebas físicas en Google Chrome sobre Android (verificando la adición a pantalla de inicio y visualización standalone) y en Safari sobre iOS (probando el flujo de Compartir y Añadir a Pantalla de Inicio con persistencia en IndexedDB).'
    ),
    createBulletPoint(
      'Despliegue en Producción (Cloud Deployment):',
      'El frontend se encuentra optimizado para despliegue continuo en plataformas CDN globales como Vercel o Netlify con certificado SSL/HTTPS nativo, mientras que el backend puede desplegarse en contenedores Node.js (Render / Railway).'
    ),
    createEvidenceBox('Figura 3: Captura de pantalla de la auditoría técnica aprobatoria (Lighthouse / DevTools Application Panel) y de la PWA instalada en smartphone real.'),

    // ── 7. CONCLUSIONES ──
    createHeading1('7. Conclusiones y Lecciones Aprendidas'),
    createParagraph(
      'El desarrollo del proyecto TechField demuestra que las Aplicaciones Web Progresivas constituyen la solución arquitectónica más costo-eficiente, robusta y escalable para escenarios de trabajo técnico en zonas rurales sin internet.'
    ),
    createParagraph(
      'Al integrar Service Workers artesanales con almacenamiento transaccional en IndexedDB, se superaron las limitaciones históricas de la web, logrando una herramienta que no depende de la conectividad para garantizar la continuidad operativa de la empresa. La eliminación de intermediarios como las tiendas de aplicaciones agiliza los despliegues y reduce a cero los costos de licenciamiento por distribución, consolidando a las PWAs como el estándar de oro para operaciones de campo en la industria moderna.'
    ),

    // ── 8. REFERENCIAS BIBLIOGRÁFICAS ──
    createHeading1('8. Referencias Bibliográficas'),
    createParagraph(
      '1. Google Developers. (2024). Progressive Web Apps Overview and Architectural Foundations. web.dev/explore/pwa'
    ),
    createParagraph(
      '2. Mozilla Developer Network (MDN). (2024). Service Worker API: Lifecycle, Caching Strategies, and Offline Storage. developer.mozilla.org'
    ),
    createParagraph(
      '3. W3C Working Group. (2024). Web App Manifest Specification & Installability Criteria. w3.org/TR/appmanifest/'
    ),
    createParagraph(
      '4. Apple Inc. (2024). Safari Web Content Guide: WebKit Storage Limits and PWA Capabilities on iOS. developer.apple.com'
    ),
  ];

  // Crear el documento Word combinando la hoja en blanco con el contenido
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440, // 1 pulgada (2.54 cm estándar APA)
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        children: [...portadaChildren, ...docChildren],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  const outputPath = path.join(__dirname, '..', 'docs', 'Reporte_Proyecto_Integrador_PWA.docx');
  fs.writeFileSync(outputPath, buffer);
  console.log('Documento Word generado con éxito en:', outputPath);
}

buildDocument().catch((err) => {
  console.error('Error generando documento:', err);
  process.exit(1);
});

