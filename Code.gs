/**
 * 2H.O — Backend de Apps Script para el brief y la solicitud de llamada.
 *
 * Este script NO sirve el sitio (index.html vive en GitHub/GitHub Pages).
 * Solo funciona como API: recibe peticiones POST en JSON y responde JSON.
 *
 * Publicación:
 * 1) En script.google.com crea un proyecto y pega este archivo como Code.gs.
 * 2) Implementar > Nueva implementación > Aplicación web.
 *    - Ejecutar como: Yo (el propietario del script, dueño del correo
 *      que aparece como TEAM_EMAIL).
 *    - Quién tiene acceso: Cualquier usuario.
 * 3) Copia la URL /exec resultante y pégala en index.html, en
 *    CONFIG.appsScriptUrl.
 * 4) La primera ejecución crea automáticamente una hoja de cálculo
 *    "2H.O — Leads piloto" (guarda su ID en Propiedades del script) donde
 *    quedan registrados los inicios y los envíos.
 */

const TEAM_EMAIL = '2horas.online01@gmail.com';
const WHATSAPP_NUMBER = '573173089057';
const SPREADSHEET_ID_PROPERTY = 'SPREADSHEET_ID';
const STARTS_SHEET_NAME = 'Inicios';
const LEADS_SHEET_NAME = 'Leads';

const NEED_LABELS = {
  Digitalizar: 'Digitalizar (mostrar servicios o catálogo)',
  Captar: 'Captar (recibir consultas)',
  Automatizar: 'Automatizar (organizar el primer contacto)',
  Validar: 'Validar (recoger interés)'
};

/* ================= Web app entry point (API JSON, sin frontend) ================= */

function doGet() {
  return jsonOutput_({ ok: true, service: '2H.O backend', usage: 'POST { action: "beginLead"|"submitLead", payload: {...} }' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action;
    const payload = body.payload || {};

    if (action === 'beginLead') return jsonOutput_(beginLead(payload));
    if (action === 'submitLead') return jsonOutput_(submitLead(payload));
    if (action === 'beginLeadExp4') return jsonOutput_(beginLeadExp4(payload));
    if (action === 'pingStepExp4') return jsonOutput_(pingStepExp4(payload));
    if (action === 'submitLeadExp4') return jsonOutput_(submitLeadExp4(payload));
    return jsonOutput_({ ok: false, message: 'Acción no reconocida.' });
  } catch (err) {
    console.error('doPost: ' + err);
    return jsonOutput_({ ok: false, message: 'No se pudo procesar la solicitud.' });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Acciones (llamadas desde doPost) ================= */

function beginLead(payload) {
  try {
    payload = payload || {};
    const leadId = String(payload.leadId || '').trim();
    if (!leadId) return { ok: false };

    const sheet = getStartsSheet_();
    const lastRow = sheet.getLastRow();
    const existingIds = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      : [];
    if (existingIds.indexOf(leadId) === -1) {
      sheet.appendRow([
        new Date(),
        leadId,
        payload.source || '',
        payload.medium || '',
        payload.campaign || ''
      ]);
    }
    return { ok: true };
  } catch (err) {
    console.error('beginLead: ' + err);
    return { ok: false };
  }
}

function submitLead(payload) {
  try {
    payload = payload || {};
    const kind = payload.kind;
    const values = payload.values || {};
    const requestId = String(payload.requestId || '').trim();
    const leadId = String(payload.leadId || '').trim();

    if (kind !== 'brief' && kind !== 'call') {
      return { ok: false, message: 'Tipo de solicitud no reconocido.' };
    }
    if (!payload.consent) {
      return { ok: false, message: 'Debes aceptar el contacto para continuar.' };
    }
    if (!isValidEmail_(values.email)) {
      return { ok: false, message: 'Revisa el correo ingresado, no parece válido.' };
    }
    if (!String(values.business || '').trim() || !String(values.person || '').trim()) {
      return { ok: false, message: 'Faltan datos obligatorios.' };
    }

    // Evita procesar dos veces la misma solicitud (reintentos de red).
    const cache = CacheService.getScriptCache();
    const cacheKey = 'req_' + requestId;
    if (requestId && cache.get(cacheKey)) {
      return { ok: true, deduped: true };
    }

    appendLeadRow_(kind, leadId, requestId, values);
    sendTeamNotification_(kind, leadId, values);
    sendUserConfirmation_(kind, values);

    if (requestId) cache.put(cacheKey, '1', 21600); // 6 horas
    return { ok: true };
  } catch (err) {
    console.error('submitLead: ' + err);
    return { ok: false, message: 'Ocurrió un error al procesar tu solicitud. Intenta de nuevo o escríbenos por WhatsApp.' };
  }
}

/* ================= Experimento 4 — Formulario v2 (popup, 15% dto.) ================= */

const EXP4_SHEET_NAME = 'Exp4_Seguimiento';
const EXP4_HEADERS = ['lead_id','inicio_utc','fuente','medio','campana','ultima_pregunta_vista','completo','negocio','descripcion_negocio','categoria','objetivo_principal','red_social_o_maps','manual_marca','manual_marca_archivo','whatsapp','claridad_formulario','fecha_completado_utc','brief_util','fecha_envio_vista_previa','notas'];

function getExp4Sheet_(){
  return getSheetByName_(EXP4_SHEET_NAME, EXP4_HEADERS);
}
function exp4Index_(){
  const map = {};
  EXP4_HEADERS.forEach(function(key, i){ map[key] = i; });
  return map;
}
function exp4FindRow_(sheet, leadId){
  const data = sheet.getDataRange().getValues();
  const idx = exp4Index_();
  for (let r = 1; r < data.length; r++) {
    if (data[r][idx.lead_id] === leadId) return r + 1;
  }
  return -1;
}
function exp4SafeCell_(value, limit){
  if (value === undefined || value === null) return '';
  return String(value).trim().slice(0, limit || 300);
}

function beginLeadExp4(payload){
  try {
    payload = payload || {};
    const leadId = exp4SafeCell_(payload.leadId, 100);
    if (!/^[a-zA-Z0-9-]{6,100}$/.test(leadId)) return { ok: false };
    const sheet = getExp4Sheet_();
    if (exp4FindRow_(sheet, leadId) !== -1) return { ok: true };
    const idx = exp4Index_();
    const row = new Array(EXP4_HEADERS.length).fill('');
    row[idx.lead_id] = leadId;
    row[idx.inicio_utc] = new Date().toISOString();
    row[idx.fuente] = exp4SafeCell_(payload.source, 80) || 'directo_desconocido';
    row[idx.medio] = exp4SafeCell_(payload.medium, 60) || 'sin_etiqueta';
    row[idx.campana] = exp4SafeCell_(payload.campaign, 80) || 'sin_etiqueta';
    row[idx.ultima_pregunta_vista] = 1;
    row[idx.completo] = 'no';
    sheet.appendRow(row);
    return { ok: true };
  } catch (err) {
    console.error('beginLeadExp4: ' + err);
    return { ok: false };
  }
}

function pingStepExp4(payload){
  try {
    payload = payload || {};
    const leadId = exp4SafeCell_(payload.leadId, 100);
    const step = Number(payload.step);
    if (!leadId || !isFinite(step)) return { ok: false };
    const sheet = getExp4Sheet_();
    const r = exp4FindRow_(sheet, leadId);
    if (r === -1) return { ok: false };
    const idx = exp4Index_();
    sheet.getRange(r, idx.ultima_pregunta_vista + 1).setValue(step);
    return { ok: true };
  } catch (err) {
    console.error('pingStepExp4: ' + err);
    return { ok: false };
  }
}

function submitLeadExp4(payload){
  try {
    payload = payload || {};
    if (payload.consent !== true || !/^[a-zA-Z0-9-]{6,100}$/.test(payload.requestId || ''))
      return { ok: false, message: 'Solicitud inválida. Revisa el formulario.' };
    const v = payload.values || {};
    const limits = { business: 100, businessDescription: 400, category: 60, objective: 80, socialLink: 200, brandManual: 60, brandManualLink: 300, phone: 25, clarity: 2 };
    const clean = {};
    for (const key in limits) {
      if (v[key] !== undefined && typeof v[key] !== 'string') return { ok: false, message: 'Formato no válido.' };
      clean[key] = String(v[key] || '').trim().slice(0, limits[key]);
    }
    const required = ['business', 'businessDescription', 'category', 'objective', 'brandManual', 'phone'];
    for (const key of required) { if (!clean[key]) return { ok: false, message: 'Completa los campos obligatorios.' }; }
    if (!/^\+?[0-9\s-]{7,25}$/.test(clean.phone)) return { ok: false, message: 'Introduce un WhatsApp válido.' };

    // Evita procesar dos veces la misma solicitud (reintentos de red).
    const cache = CacheService.getScriptCache();
    const requestId = String(payload.requestId || '').trim();
    const cacheKey = 'exp4req_' + requestId;
    if (requestId && cache.get(cacheKey)) return { ok: true, deduped: true };

    const fileName = exp4SafeCell_(v.brandManualFileName, 150);
    const fileType = exp4SafeCell_(v.brandManualFileType, 100);
    const fileData = typeof v.brandManualFileData === 'string' ? v.brandManualFileData : '';
    if (fileData.length > 11 * 1024 * 1024) return { ok: false, message: 'El archivo adjunto es demasiado grande.' };
    let attachments = [];
    let manualNote = 'No';
    if (fileData) {
      const bytes = Utilities.base64Decode(fileData);
      attachments.push(Utilities.newBlob(bytes, fileType || 'application/octet-stream', fileName || 'manual-marca'));
      manualNote = 'Sí (archivo adjunto: ' + (fileName || 'archivo') + ')';
    } else if (clean.brandManualLink) {
      manualNote = 'Sí (link: ' + clean.brandManualLink + ')';
    } else if (clean.brandManual.indexOf('Sí') === 0) {
      manualNote = 'Sí (sin archivo)';
    }

    const leadId = exp4SafeCell_(payload.leadId, 100);
    const sheet = getExp4Sheet_();
    let r = leadId ? exp4FindRow_(sheet, leadId) : -1;
    const idx = exp4Index_();
    if (r === -1) {
      const row = new Array(EXP4_HEADERS.length).fill('');
      row[idx.lead_id] = leadId || Utilities.getUuid();
      row[idx.inicio_utc] = new Date().toISOString();
      sheet.appendRow(row);
      r = sheet.getLastRow();
    }
    sheet.getRange(r, idx.negocio + 1).setValue(clean.business);
    sheet.getRange(r, idx.descripcion_negocio + 1).setValue(clean.businessDescription);
    sheet.getRange(r, idx.categoria + 1).setValue(clean.category);
    sheet.getRange(r, idx.objetivo_principal + 1).setValue(clean.objective);
    sheet.getRange(r, idx.red_social_o_maps + 1).setValue(clean.socialLink);
    sheet.getRange(r, idx.manual_marca + 1).setValue(clean.brandManual);
    sheet.getRange(r, idx.manual_marca_archivo + 1).setValue(manualNote);
    sheet.getRange(r, idx.whatsapp + 1).setValue(clean.phone);
    sheet.getRange(r, idx.claridad_formulario + 1).setValue(clean.clarity || '');
    sheet.getRange(r, idx.completo + 1).setValue('si');
    sheet.getRange(r, idx.fecha_completado_utc + 1).setValue(new Date().toISOString());

    sendExp4TeamNotification_(clean, manualNote, payload.requestId, attachments);
    if (requestId) cache.put(cacheKey, '1', 21600); // 6 horas
    return { ok: true };
  } catch (err) {
    console.error('submitLeadExp4: ' + err);
    return { ok: false, message: 'No se pudo confirmar el envío. Revisa los datos o contacta al equipo por WhatsApp para verificar si llegó.' };
  }
}

function sendExp4TeamNotification_(clean, manualNote, requestId, attachments){
  const labels = { business: 'Negocio', businessDescription: 'Qué hace el negocio', category: 'Categoría', objective: 'Objetivo principal', socialLink: 'Instagram / Facebook / Maps', phone: 'WhatsApp', clarity: 'Claridad del formulario (1-5)' };
  const rows = Object.keys(labels).map(function(key){ return [labels[key], clean[key] || 'No indicado']; });
  rows.push(['Manual de marca', manualNote]);
  const bodyRows = rows
    .map(([label, value]) => '<tr><td style="padding:6px 12px 6px 0;color:#6b6f5c;white-space:nowrap">' + escapeHtml_(label) + '</td><td style="padding:6px 0;font-weight:600">' + escapeHtml_(value || '—') + '</td></tr>')
    .join('');
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#12130e">' +
    '<p>Nuevo brief con descuento del 15% reservado.</p>' +
    '<table cellpadding="0" cellspacing="0">' + bodyRows + '</table>' +
    '<p style="margin-top:16px;font-size:12px;color:#6b6f5c">Referencia: ' + escapeHtml_(requestId) + '</p>' +
    '</div>';
  const mailOptions = { to: TEAM_EMAIL, subject: 'Nuevo brief (Exp. 4 — 15% dto.) — 2HO', htmlBody: html, name: '2HO · Experimento 4' };
  if (attachments && attachments.length) mailOptions.attachments = attachments;
  MailApp.sendEmail(mailOptions);
}

/* ================= Sheet helpers ================= */

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SPREADSHEET_ID_PROPERTY);
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      console.error('No se pudo abrir la hoja guardada, se creará una nueva: ' + err);
    }
  }
  const ss = SpreadsheetApp.create('2H.O — Leads piloto');
  props.setProperty(SPREADSHEET_ID_PROPERTY, ss.getId());
  return ss;
}

function getSheetByName_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    const blank = ss.getSheets().find(s => s.getName() !== name && s.getLastRow() === 0 && s.getLastColumn() <= 1);
    if (blank) ss.deleteSheet(blank);
  }
  return sheet;
}

function getStartsSheet_() {
  return getSheetByName_(STARTS_SHEET_NAME, ['Fecha', 'Lead ID', 'Fuente', 'Medio', 'Campaña']);
}

function getLeadsSheet_() {
  return getSheetByName_(LEADS_SHEET_NAME, [
    'Fecha', 'Tipo', 'Lead ID', 'Request ID', 'Negocio', 'Oferta', 'Necesidad',
    'Objetivo', 'Referencia (Instagram)', 'Estilo', 'Nombre', 'Correo',
    'WhatsApp', 'Fecha llamada', 'Hora llamada'
  ]);
}

function appendLeadRow_(kind, leadId, requestId, values) {
  getLeadsSheet_().appendRow([
    new Date(),
    kind === 'brief' ? 'Brief' : 'Llamada',
    leadId,
    requestId,
    values.business || '',
    values.offer || '',
    values.need || '',
    values.goal || '',
    values.reference || '',
    values.style || '',
    values.person || '',
    values.email || '',
    values.phone || '',
    values.date || '',
    values.time || ''
  ]);
}

/* ================= Emails ================= */

function sendTeamNotification_(kind, leadId, values) {
  const isBrief = kind === 'brief';
  const subject = isBrief
    ? 'Nuevo brief: ' + values.business
    : 'Nueva solicitud de llamada: ' + values.business;

  const rows = isBrief
    ? [
        ['Negocio', values.business],
        ['Oferta', values.offer],
        ['Necesidad', NEED_LABELS[values.need] || values.need],
        ['Referencia (Instagram)', values.reference],
        ['Estilo', values.style],
        ['Nombre', values.person],
        ['Correo', values.email],
        ['WhatsApp', values.phone || '—']
      ]
    : [
        ['Negocio', values.business],
        ['Nombre', values.person],
        ['Correo', values.email],
        ['WhatsApp', values.phone || '—'],
        ['Fecha propuesta', values.date],
        ['Hora propuesta', values.time]
      ];

  const bodyRows = rows
    .map(([label, value]) => '<tr><td style="padding:6px 12px 6px 0;color:#6b6f5c;white-space:nowrap">' + escapeHtml_(label) + '</td><td style="padding:6px 0;font-weight:600">' + escapeHtml_(value || '—') + '</td></tr>')
    .join('');

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#12130e">' +
    '<p>Lead ID: <code>' + escapeHtml_(leadId) + '</code></p>' +
    '<table cellpadding="0" cellspacing="0">' + bodyRows + '</table>' +
    '</div>';

  MailApp.sendEmail({
    to: TEAM_EMAIL,
    subject: subject,
    htmlBody: html,
    name: '2H.O — Notificaciones',
    replyTo: values.email
  });
}

function sendUserConfirmation_(kind, values) {
  const isBrief = kind === 'brief';
  const subject = isBrief
    ? 'Recibimos tu brief, ' + values.person + ' — 2H.O'
    : 'Tu llamada con 2H.O está en camino, ' + values.person;

  const html = isBrief
    ? buildBriefEmail_(values)
    : buildCallEmail_(values);

  MailApp.sendEmail({
    to: values.email,
    subject: subject,
    htmlBody: html,
    name: '2H.O',
    replyTo: TEAM_EMAIL
  });
}

function buildBriefEmail_(values) {
  const summaryRows = [
    ['Negocio', values.business],
    ['Oferta', values.offer],
    ['Lo que buscas', NEED_LABELS[values.need] || values.need],
    ['Tu Instagram hoy', values.reference],
    ['Estilo preferido', values.style],
    ['WhatsApp', values.phone || 'No indicado']
  ];
  const summaryHtml = summaryTable_(summaryRows);
  const whatsappText = 'Hola, 2HO. Soy ' + values.person + ' de ' + values.business + '. Acabo de enviar mi brief y quiero conversar sobre mi web.';

  const body =
    '<p style="margin:0 0 16px">Hola ' + escapeHtml_(values.person) + ', ¡gracias por completar tu brief! Este es el resumen que le llegó a nuestro equipo:</p>' +
    summaryHtml +
    '<p style="margin:24px 0 8px;font-weight:700;font-size:16px">¿Qué sigue?</p>' +
    '<p style="margin:0 0 16px">Nuestro equipo revisa tu brief y te escribe con la primera vista previa de tu web en un plazo de <strong>24 a 48 horas</strong>. La página no se genera automáticamente al enviar el formulario: la prepara una persona del equipo.</p>' +
    '<p style="margin:0 0 20px">¿Tienes una duda puntual antes de eso, quieres ajustar algo de tu brief o prefieres contarnos por voz? Escríbenos por WhatsApp cuando quieras:</p>' +
    ctaButton_(waHref_(whatsappText), 'Hablar por WhatsApp') +
    '<p style="margin:24px 0 0;font-size:13px;color:#6b6f5c">¿Prefieres agendar una llamada guiada en vez de esperar? Vuelve a la página y elige "Agendar llamada" para proponer un horario; te confirmaremos por correo o WhatsApp antes de llamarte.</p>';

  return emailShell_('Tu brief llegó a 2H.O', body);
}

function buildCallEmail_(values) {
  const summaryRows = [
    ['Negocio', values.business],
    ['Fecha propuesta', formatCallDate_(values.date)],
    ['Hora propuesta', values.time || '—'],
    ['WhatsApp', values.phone || 'No indicado']
  ];
  const summaryHtml = summaryTable_(summaryRows);
  const whatsappText = 'Hola, 2HO. Soy ' + values.person + ' de ' + values.business + '. Propuse una llamada y quiero confirmar el horario.';

  const body =
    '<p style="margin:0 0 16px">Hola ' + escapeHtml_(values.person) + ', recibimos tu propuesta de horario para hablar sobre tu web. Este es el resumen:</p>' +
    summaryHtml +
    '<p style="margin:24px 0 8px;font-weight:700;font-size:16px">¿Cómo te contactaremos?</p>' +
    '<p style="margin:0 0 16px">El horario que propusiste <strong>todavía no es una reserva confirmada</strong>. Nuestro equipo va a escribirte por correo o WhatsApp para confirmar disponibilidad antes de llamarte al número que dejaste. Si ese horario no funciona, te propondremos uno alterno.</p>' +
    '<p style="margin:0 0 20px">Si necesitas adelantar algo, cambiar el horario o tienes una pregunta puntual mientras tanto, escríbenos directo por WhatsApp:</p>' +
    ctaButton_(waHref_(whatsappText), 'Hablar por WhatsApp') +
    '<p style="margin:24px 0 0;font-size:13px;color:#6b6f5c">Ten a mano tu teléfono en la fecha y hora propuestas: así confirmamos y avanzamos más rápido en la llamada.</p>';

  return emailShell_('Tu solicitud de llamada llegó a 2H.O', body);
}

/* ================= Email building blocks ================= */

function emailShell_(heading, innerHtml) {
  return (
    '<div style="background:#eeebdc;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
    '<div style="max-width:520px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid #e3e0d1">' +
      '<div style="background:#12130e;padding:22px 28px">' +
        '<span style="color:#d7f23a;font-weight:800;font-size:18px;letter-spacing:.02em">2H.O</span>' +
        '<span style="color:#a9ad97;font-size:12px;display:block;margin-top:2px">Tu web lista en 24–48 horas</span>' +
      '</div>' +
      '<div style="padding:28px">' +
        '<h1 style="margin:0 0 16px;font-size:20px;color:#12130e">' + escapeHtml_(heading) + '</h1>' +
        '<div style="font-size:14px;line-height:1.55;color:#12130e">' + innerHtml + '</div>' +
      '</div>' +
      '<div style="background:#f5f6ee;padding:16px 28px;font-size:12px;color:#6b6f5c;text-align:center">' +
        'Recibiste este correo porque completaste un formulario en la página de 2H.O.<br>' +
        'Dudas sobre tus datos: <a href="mailto:' + TEAM_EMAIL + '" style="color:#6b6f5c">' + TEAM_EMAIL + '</a>' +
      '</div>' +
    '</div>' +
    '</div>'
  );
}

function summaryTable_(rows) {
  const trs = rows
    .map(([label, value]) => (
      '<tr>' +
      '<td style="padding:9px 12px;background:#f5f6ee;border-bottom:1px solid #eeebdc;color:#6b6f5c;font-size:12px;text-transform:uppercase;letter-spacing:.03em;white-space:nowrap;vertical-align:top">' + escapeHtml_(label) + '</td>' +
      '<td style="padding:9px 12px;background:#f5f6ee;border-bottom:1px solid #eeebdc;font-weight:600">' + escapeHtml_(value || '—') + '</td>' +
      '</tr>'
    ))
    .join('');
  return '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border-radius:10px;overflow:hidden;margin:8px 0 4px">' + trs + '</table>';
}

function ctaButton_(href, label) {
  return (
    '<table cellpadding="0" cellspacing="0"><tr><td style="border-radius:999px;background:#3ddc7a">' +
    '<a href="' + href + '" style="display:inline-block;padding:12px 24px;font-weight:700;font-size:14px;color:#12130e;text-decoration:none">' + escapeHtml_(label) + '</a>' +
    '</td></tr></table>'
  );
}

function waHref_(text) {
  return 'https://wa.me/' + WHATSAPP_NUMBER + '?text=' + encodeURIComponent(text);
}

const WEEKDAYS_ES = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MONTHS_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

function formatCallDate_(isoDate) {
  if (!isoDate) return '—';
  const parts = String(isoDate).split('-');
  if (parts.length !== 3) return isoDate;
  const date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  if (isNaN(date.getTime())) return isoDate;
  return WEEKDAYS_ES[date.getDay()] + ' ' + date.getDate() + ' de ' + MONTHS_ES[date.getMonth()] + ' de ' + date.getFullYear();
}

/* ================= Utils ================= */

function isValidEmail_(email) {
  return typeof email === 'string' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
