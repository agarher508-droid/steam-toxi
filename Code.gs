const CFG = {
  title: 'TOXIRIA STEAM LAB',
  folderName: 'TOXIRIA STEAM LAB 2026-2027',
  spreadsheetName: 'Registro integral STEAM - CEIP Toxiria',
  sheets: {
    kits: 'Uso de kits',
    sda: 'Situaciones de aprendizaje'
  },
  kits: [
    'Gator-Bot', 'MatataStudio Coding Set Lite', "Let’s Go Code!",
    'Maqueen Plus', 'micro:bit GO', 'Makey Makey', 'Echidna Black V2',
    'Smart Agriculture IoT', 'Wonder Building', 'MakeCode Arcade',
    'Kit textil electrónico', 'Kit de circuitos en papel',
    'Impresora 3D Bambu Lab P1S Combo', 'Plotter Silhouette Cameo 5 Alpha',
    'Bolígrafos de impresión 3D', 'Pirámides holográficas', 'Tablets',
    'Portátiles', 'Micrófonos', 'Chroma y trípode', 'Otros'
  ]
};

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('STEAM_SS_ID');
  let folderId = props.getProperty('STEAM_FOLDER_ID');

  let ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create(CFG.spreadsheetName);
    props.setProperty('STEAM_SS_ID', ss.getId());
  }

  let folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    folder = DriveApp.createFolder(CFG.folderName);
    props.setProperty('STEAM_FOLDER_ID', folder.getId());
  }

  ensureSheet_(ss, CFG.sheets.kits, [
    'ID', 'Fecha registro', 'Fecha uso', 'Curso/grupo', 'Docente',
    'Recurso o kit', 'Cantidad', 'Hora inicio', 'Hora fin',
    'Actividad/propósito', 'Estado del material', 'Observaciones'
  ]);
  ensureSheet_(ss, CFG.sheets.sda, [
    'Código', 'Fecha creación', 'Título', 'Curso', 'Ciclo',
    'Trimestre', 'Responsable', 'Modelo STEAM', 'Reto',
    'Producto final', 'Recursos', 'Estado', 'URL Google Docs',
    'Competencias', 'Criterios de evaluación', 'Saberes básicos',
    'Actividades', 'Evaluación', 'Evidencias'
  ]);

  Logger.log('REGISTRO: ' + ss.getUrl());
  Logger.log('CARPETA: ' + folder.getUrl());
  return {spreadsheetUrl:ss.getUrl(), folderUrl:folder.getUrl()};
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CFG.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('STEAM_SS_ID');
  if (!ssId) throw new Error('Primero ejecuta la función setup desde Apps Script.');
  const ss = SpreadsheetApp.openById(ssId);
  return {
    kits: CFG.kits,
    kitRows: getRows_(ss.getSheetByName(CFG.sheets.kits), 12),
    sdaRows: getRows_(ss.getSheetByName(CFG.sheets.sda), 19),
    spreadsheetUrl: ss.getUrl(),
    project: {
      title: CFG.title,
      year: '2026-2027',
      school: 'CEIP Toxiria · Torredonjimeno (Jaén)'
    }
  };
}

function registrarUsoKit(d) {
  if (!d || !d.fecha || !d.curso || !d.docente || !d.kit || !d.cantidad) {
    throw new Error('Completa fecha, curso, docente, kit y cantidad.');
  }
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.sheets.kits);
  const id = 'KIT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  sh.appendRow([
    id, new Date(), d.fecha, d.curso, d.docente, d.kit,
    Number(d.cantidad), d.inicio || '', d.fin || '', d.actividad || '',
    d.estado || 'Correcto', d.observaciones || ''
  ]);
  return {ok:true, id:id, spreadsheetUrl:ss.getUrl()};
}

function crearSDA(d) {
  if (!d || !d.titulo || !d.curso || !d.responsable) {
    throw new Error('El título, curso y responsable son obligatorios.');
  }
  const props = PropertiesService.getScriptProperties();
  const folder = DriveApp.getFolderById(props.getProperty('STEAM_FOLDER_ID'));
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.sheets.sda);
  const now = new Date();
  const code = 'SDA-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const doc = DocumentApp.create(code + ' - ' + d.titulo);
  const body = doc.getBody();

  body.appendParagraph('CEIP TOXIRIA').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('PLAN STEAM 4.0 · CURSO 2026-2027').setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  body.appendParagraph(d.titulo).setHeading(DocumentApp.ParagraphHeading.HEADING1);

  addSection_(body, '1. IDENTIFICACIÓN', [
    ['Código', code], ['Curso', d.curso], ['Ciclo', d.ciclo],
    ['Trimestre', d.trimestre], ['Responsable', d.responsable],
    ['Modelo STEAM', d.modelo]
  ]);
  addSection_(body, '2. JUSTIFICACIÓN Y RETO', [
    ['Situación o problema', d.reto], ['Pregunta guía', d.pregunta],
    ['Producto final', d.producto]
  ]);
  addSection_(body, '3. OBJETIVOS DE APRENDIZAJE', [['Objetivos', d.objetivos]]);
  addSection_(body, '4. VINCULACIÓN CURRICULAR', [
    ['Competencias específicas', d.competencias],
    ['Criterios de evaluación', d.criterios], ['Saberes básicos', d.saberes]
  ]);
  addSection_(body, '5. SECUENCIACIÓN DE ACTIVIDADES', [['Actividades y sesiones', d.actividades]]);
  addSection_(body, '6. RECURSOS STEAM', [['Materiales y kits', d.recursos]]);
  addSection_(body, '7. METODOLOGÍA Y AGRUPAMIENTOS', [['Metodología', d.metodologia]]);
  addSection_(body, '8. ATENCIÓN A LA DIVERSIDAD', [['Medidas y apoyos', d.inclusion]]);
  addSection_(body, '9. EVALUACIÓN', [
    ['Instrumentos', d.evaluacion], ['Indicadores de logro', d.indicadores]
  ]);
  addSection_(body, '10. PRODUCTOS Y EVIDENCIAS', [['Evidencias previstas', d.evidencias]]);
  addSection_(body, '11. VALORACIÓN FINAL', [['Resultados y mejoras', 'Pendiente de completar tras la aplicación.']]);
  body.appendParagraph('© Ana Garrido Herrada').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(folder);

  const url = doc.getUrl();
  sh.appendRow([
    code, now, d.titulo, d.curso, d.ciclo || '', d.trimestre || '',
    d.responsable, d.modelo || '', d.reto || '', d.producto || '',
    d.recursos || '', 'Borrador', url, d.competencias || '',
    d.criterios || '', d.saberes || '', d.actividades || '',
    d.evaluacion || '', d.evidencias || ''
  ]);
  return {ok:true, code:code, url:url, spreadsheetUrl:ss.getUrl()};
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('STEAM_SS_ID');
  if (!id) throw new Error('Ejecuta primero setup().');
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setWrap(true);
    sh.autoResizeColumns(1, Math.min(headers.length, 12));
  }
}

function getRows_(sh, width) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, width).getDisplayValues().reverse().slice(0, 100);
}

function addSection_(body, title, fields) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  fields.forEach(function(f) {
    body.appendParagraph(f[0]).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(f[1] || 'Pendiente de completar.');
  });
}
