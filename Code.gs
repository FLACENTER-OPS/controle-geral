const DEAL_MASTER_HEADERS = [
  'deal_id', 'stock', 'vin', 'vehicle', 'seller', 'customer_name', 'sale_date',
  'lender', 'deal_type', 'trade_in', 'gps_required', 'parcelado', 'out_of_state',
  'company_deal', 'review_status', 'commission_status', 'dmv_status',
  'parcelamento_status', 'title_status', 'omnia_status', 'folder_status',
  'envelope_status', 'created_at', 'updated_at', 'hub_status'
];

const DEAL_ALERTS_HEADERS = [
  'deal_id', 'alert_type', 'priority', 'status', 'created_at', 'resolved_at'
];

const DEAL_DOCUMENTS_HEADERS = [
  'deal_id', 'document_name', 'printed_status', 'dealer_center_status', 'status', 'created_at', 'updated_at'
];

const DEAL_HISTORY_HEADERS = [
  'deal_id', 'action', 'user', 'notes', 'timestamp'
];

const DEAL_NOTES_HEADERS = [
  'deal_id', 'note_text', 'user', 'note_type', 'created_at'
];

const DEAL_MASTER_STATUS_FIELDS = [
  'review_status',
  'commission_status',
  'dmv_status',
  'parcelamento_status',
  'title_status',
  'omnia_status',
  'folder_status',
  'envelope_status'
];

function doGet(e) {
  try {
    const ss = getCentralSpreadsheet();
    const action = e && e.parameter ? e.parameter.action : '';

    if (action === 'getDealDetails') {
      return getDealDetailsResponse(ss, e.parameter.deal_id || e.parameter.id || '');
    }

    if (action === 'getDealNotes') {
      return getDealNotesResponse(ss, e.parameter.deal_id || e.parameter.id || '');
    }

    const masterSheet = ensureSheet(ss, 'DEAL_MASTER', DEAL_MASTER_HEADERS);
    const alertsSheet = ensureSheet(ss, 'DEAL_ALERTS', DEAL_ALERTS_HEADERS);
    const deals = getDealMasterRows(masterSheet);
    const alertsByDeal = getOpenAlertsByDeal(alertsSheet);

    const enrichedDeals = deals.map(deal => {
      const alerts = alertsByDeal[String(deal.deal_id)] || [];
      return {
        ...deal,
        alerts,
        alert_count: alerts.length,
        critical_alert_count: alerts.filter(isCriticalAlert).length,
      };
    });

    return jsonOutput({
      success: true,
      source: 'DEAL_MASTER',
      alerts_source: 'DEAL_ALERTS',
      hub_ready: true,
      deals: enrichedDeals,
    });
  } catch (err) {
    return jsonOutput({ success: false, error: err.message });
  }
}

function getDealDetailsResponse(ss, dealId) {
  if (!dealId) throw new Error('deal_id obrigatorio');

  const masterSheet = ensureSheet(ss, 'DEAL_MASTER', DEAL_MASTER_HEADERS);
  const alertsSheet = ensureSheet(ss, 'DEAL_ALERTS', DEAL_ALERTS_HEADERS);
  const historySheet = ensureSheet(ss, 'DEAL_HISTORY', DEAL_HISTORY_HEADERS);
  const documentsSheet = ensureSheet(ss, 'DEAL_DOCUMENTS', DEAL_DOCUMENTS_HEADERS);
  const notesSheet = ensureSheet(ss, 'DEAL_NOTES', DEAL_NOTES_HEADERS);
  const deal = getDealMasterById(masterSheet, dealId);

  if (!deal) throw new Error('Deal nao encontrado no DEAL_MASTER: ' + dealId);

  return jsonOutput({
    success: true,
    source: 'DEAL_MASTER',
    deal,
    alerts: getRowsByDealId(alertsSheet, dealId).filter(alert => !isResolvedAlert(alert)),
    history: getRowsByDealId(historySheet, dealId),
    documents: getRowsByDealId(documentsSheet, dealId),
    notes: getRowsByDealId(notesSheet, dealId),
  });
}

function getDealNotesResponse(ss, dealId) {
  if (!dealId) throw new Error('deal_id obrigatorio');
  const notesSheet = ensureSheet(ss, 'DEAL_NOTES', DEAL_NOTES_HEADERS);
  return jsonOutput({
    success: true,
    notes: getRowsByDealId(notesSheet, dealId),
  });
}

function getDealMasterById(sheet, dealId) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(String);

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(dealId)) return rowToObject(headers, values[i]);
  }
  return null;
}

function getRowsByDealId(sheet, dealId) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);

  return values.slice(1)
    .filter(row => String(row[0]) === String(dealId))
    .map(row => rowToObject(headers, row));
}

function doPost(e) {
  try {
    const ss = getCentralSpreadsheet();
    const payload = JSON.parse(e.postData.contents || '{}');
    const masterSheet = ensureSheet(ss, 'DEAL_MASTER', DEAL_MASTER_HEADERS);
    const historySheet = ensureSheet(ss, 'DEAL_HISTORY', DEAL_HISTORY_HEADERS);

    if (payload.action === 'updateStatus') {
      updateDealMasterStatus(masterSheet, historySheet, payload);
      return jsonOutput({ success: true });
    }

    if (payload.action === 'updateDocumentStatus') {
      const documentsSheet = ensureSheet(ss, 'DEAL_DOCUMENTS', DEAL_DOCUMENTS_HEADERS);
      const document = updateDealDocumentStatus(documentsSheet, historySheet, payload);
      return jsonOutput({ success: true, document });
    }

    if (payload.action === 'addDealNote') {
      const notesSheet = ensureSheet(ss, 'DEAL_NOTES', DEAL_NOTES_HEADERS);
      const note = addDealNote(notesSheet, historySheet, payload);
      return jsonOutput({ success: true, note });
    }

    if (payload.action === 'saveDealMaster' || payload.action === 'save') {
      upsertDealMaster(masterSheet, historySheet, payload);
      return jsonOutput({ success: true, id: payload.deal_id || payload.dealId || payload.id });
    }

    return jsonOutput({ success: false, error: 'Acao invalida' });
  } catch (err) {
    return jsonOutput({ success: false, error: err.message });
  }
}

function getCentralSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('CENTRAL_DB_ID');
  if (!id) throw new Error('CENTRAL_DB_ID nao configurado nas Script Properties.');
  return SpreadsheetApp.openById(id);
}

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  headers.forEach((header, index) => {
    if (currentHeaders[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
  sheet.setFrozenRows(1);
  return sheet;
}

function getDealMasterRows(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);

  return values.slice(1)
    .filter(row => row[0])
    .map(row => rowToObject(headers, row))
    .sort((a, b) => String(b.updated_at || b.created_at || '').localeCompare(String(a.updated_at || a.created_at || '')));
}

function getOpenAlertsByDeal(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};
  const headers = values[0].map(String);
  const alertsByDeal = {};

  values.slice(1).forEach(row => {
    const alert = rowToObject(headers, row);
    const dealId = String(alert.deal_id || '');
    if (!dealId || isResolvedAlert(alert)) return;
    if (!alertsByDeal[dealId]) alertsByDeal[dealId] = [];
    alertsByDeal[dealId].push(alert);
  });

  return alertsByDeal;
}

function updateDealMasterStatus(masterSheet, historySheet, payload) {
  const dealId = payload.deal_id || payload.dealId || payload.id;
  const field = payload.field;
  if (!dealId) throw new Error('deal_id obrigatorio');
  if (!DEAL_MASTER_STATUS_FIELDS.includes(field)) throw new Error('Campo de status invalido: ' + field);

  const row = findDealMasterRow(masterSheet, dealId);
  if (row < 1) throw new Error('Deal nao encontrado no DEAL_MASTER: ' + dealId);

  const column = DEAL_MASTER_HEADERS.indexOf(field) + 1;
  masterSheet.getRange(row, column).setValue(payload.value || '');
  if (field === 'dmv_status') {
    const hubStatusColumn = DEAL_MASTER_HEADERS.indexOf('hub_status') + 1;
    const currentHubStatus = String(masterSheet.getRange(row, hubStatusColumn).getValue() || '').toLowerCase();
    if (payload.value === 'done') {
      masterSheet.getRange(row, hubStatusColumn).setValue('completed');
    } else if (currentHubStatus === 'completed') {
      masterSheet.getRange(row, hubStatusColumn).setValue('active');
    }
  }
  masterSheet.getRange(row, DEAL_MASTER_HEADERS.indexOf('updated_at') + 1).setValue(new Date().toISOString());
  appendHistory(historySheet, dealId, 'update_status', payload.user || '', field + ' = ' + (payload.value || ''));
}

function updateDealDocumentStatus(documentsSheet, historySheet, payload) {
  const dealId = payload.deal_id || payload.dealId || payload.id;
  const documentName = payload.document_name || payload.documentName;
  const field = payload.field;
  const value = payload.value || '';

  if (!dealId) throw new Error('deal_id obrigatorio');
  if (!documentName) throw new Error('document_name obrigatorio');
  if (!['printed_status', 'dealer_center_status'].includes(field)) {
    throw new Error('Campo de documento invalido: ' + field);
  }
  if (!['pending', 'ok', 'missing', ''].includes(value)) {
    throw new Error('Status de documento invalido: ' + value);
  }

  const row = findDealDocumentRow(documentsSheet, dealId, documentName);
  if (row < 1) throw new Error('Documento nao encontrado no DEAL_DOCUMENTS: ' + documentName);

  const headers = getSheetHeaders(documentsSheet, DEAL_DOCUMENTS_HEADERS);
  const currentValues = documentsSheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const document = rowToObject(headers, currentValues);
  document[field] = value;
  document.status = calculateDocumentStatus(document.printed_status, document.dealer_center_status);
  document.updated_at = new Date().toISOString();

  documentsSheet.getRange(row, DEAL_DOCUMENTS_HEADERS.indexOf(field) + 1).setValue(value);
  documentsSheet.getRange(row, DEAL_DOCUMENTS_HEADERS.indexOf('status') + 1).setValue(document.status);
  documentsSheet.getRange(row, DEAL_DOCUMENTS_HEADERS.indexOf('updated_at') + 1).setValue(document.updated_at);

  appendHistory(
    historySheet,
    dealId,
    'update_document',
    payload.user || '',
    documentName + ': ' + field + ' = ' + value + ', status = ' + document.status
  );

  return document;
}

function addDealNote(notesSheet, historySheet, payload) {
  const dealId = payload.deal_id || payload.dealId || payload.id;
  const noteText = String(payload.note_text || payload.noteText || '').trim();
  const user = payload.user || '';
  const createdAt = new Date().toISOString();

  if (!dealId) throw new Error('deal_id obrigatorio');
  if (!noteText) throw new Error('note_text obrigatorio');

  const note = {
    deal_id: dealId,
    note_text: noteText,
    user,
    note_type: 'internal',
    created_at: createdAt,
  };

  notesSheet.appendRow(DEAL_NOTES_HEADERS.map(header => note[header] || ''));
  appendHistory(historySheet, dealId, 'add_internal_note', user, noteText);
  return note;
}

function calculateDocumentStatus(printedStatus, dealerCenterStatus) {
  const printed = String(printedStatus || 'pending').toLowerCase();
  const dealerCenter = String(dealerCenterStatus || 'pending').toLowerCase();
  if (printed === 'missing' || dealerCenter === 'missing') return 'critical';
  if (printed === 'ok' && dealerCenter === 'ok') return 'ok';
  return 'pending';
}

function upsertDealMaster(masterSheet, historySheet, payload) {
  const now = new Date().toISOString();
  const dealId = payload.deal_id || payload.dealId || payload.id;
  if (!dealId) throw new Error('deal_id obrigatorio');

  const row = findDealMasterRow(masterSheet, dealId);
  const existing = row > 0 ? masterSheet.getRange(row, 1, 1, DEAL_MASTER_HEADERS.length).getValues()[0] : [];
  const existingByHeader = rowToObject(DEAL_MASTER_HEADERS, existing);
  const merged = {
    ...existingByHeader,
    deal_id: dealId,
    stock: pick(payload.stock, payload.stockNumber, payload.stock_number, existingByHeader.stock),
    vin: pick(payload.vin, existingByHeader.vin),
    vehicle: pick(payload.vehicle, payload.veiculo, existingByHeader.vehicle),
    seller: pick(payload.seller, payload.vendedor, existingByHeader.seller),
    customer_name: pick(payload.customer_name, payload.customerName, payload.cliente, existingByHeader.customer_name),
    sale_date: pick(payload.sale_date, payload.data, payload.dataVenda, existingByHeader.sale_date),
    lender: pick(payload.lender, existingByHeader.lender),
    deal_type: pick(payload.deal_type, payload.finance, existingByHeader.deal_type),
    trade_in: pick(payload.trade_in, payload.tradein, existingByHeader.trade_in),
    gps_required: pick(payload.gps_required, payload.gps, existingByHeader.gps_required),
    parcelado: pick(payload.parcelado, existingByHeader.parcelado),
    out_of_state: pick(payload.out_of_state, existingByHeader.out_of_state),
    company_deal: pick(payload.company_deal, payload.companyDeal, existingByHeader.company_deal),
    review_status: pick(payload.review_status, existingByHeader.review_status, 'pending'),
    commission_status: pick(payload.commission_status, existingByHeader.commission_status, 'pending'),
    dmv_status: pick(payload.dmv_status, existingByHeader.dmv_status, 'pending'),
    parcelamento_status: pick(payload.parcelamento_status, existingByHeader.parcelamento_status),
    title_status: pick(payload.title_status, existingByHeader.title_status),
    omnia_status: pick(payload.omnia_status, existingByHeader.omnia_status),
    folder_status: pick(payload.folder_status, existingByHeader.folder_status),
    envelope_status: pick(payload.envelope_status, existingByHeader.envelope_status),
    created_at: existingByHeader.created_at || now,
    updated_at: now,
    hub_status: pick(payload.hub_status, existingByHeader.hub_status, payload.dmv_status === 'done' ? 'completed' : 'active'),
  };

  const values = DEAL_MASTER_HEADERS.map(header => merged[header] || '');
  if (row > 0) {
    masterSheet.getRange(row, 1, 1, values.length).setValues([values]);
    appendHistory(historySheet, dealId, 'update_master', payload.user || '', 'Deal atualizado pelo controle-geral');
  } else {
    masterSheet.appendRow(values);
    appendHistory(historySheet, dealId, 'create_master', payload.user || '', 'Deal criado pelo controle-geral');
  }
}

function getSheetHeaders(sheet, defaultHeaders) {
  if (sheet.getLastRow() < 1) return defaultHeaders;
  return sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), defaultHeaders.length)).getValues()[0].map(String);
}

function findDealMasterRow(sheet, dealId) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(dealId)) return i + 1;
  }
  return -1;
}

function findDealDocumentRow(sheet, dealId, documentName) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(dealId) && String(values[i][1]) === String(documentName)) return i + 1;
  }
  return -1;
}

function appendHistory(sheet, dealId, action, user, notes) {
  sheet.appendRow([dealId || '', action || '', user || '', notes || '', new Date().toISOString()]);
}

function rowToObject(headers, row) {
  return headers.reduce((acc, header, index) => {
    acc[header] = row[index] === undefined ? '' : row[index];
    return acc;
  }, {});
}

function isResolvedAlert(alert) {
  const status = String(alert.status || '').toLowerCase();
  return Boolean(alert.resolved_at) || ['resolved', 'resolvido', 'closed', 'fechado'].includes(status);
}

function isCriticalAlert(alert) {
  return ['critical', 'critico', 'critico', 'high', 'alta'].includes(String(alert.priority || '').toLowerCase());
}

function pick() {
  for (let i = 0; i < arguments.length; i++) {
    if (arguments[i] !== undefined && arguments[i] !== null && arguments[i] !== '') return arguments[i];
  }
  return '';
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
