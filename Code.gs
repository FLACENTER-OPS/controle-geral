const DEAL_MASTER_HEADERS = [
  'deal_id', 'stock', 'vin', 'vehicle', 'seller', 'customer_name', 'sale_date',
  'lender', 'deal_type', 'trade_in', 'gps_required', 'parcelado', 'out_of_state',
  'company_deal', 'review_status', 'commission_status', 'dmv_status',
  'parcelamento_status', 'title_status', 'omnia_status', 'folder_status',
  'envelope_status', 'created_at', 'updated_at', 'hub_status', 'document_status'
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
  const health = syncDealHealth(ss, dealId);
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
    health,
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
      const health = syncDealHealth(ss, payload.deal_id || payload.dealId || payload.id);
      return jsonOutput({ success: true, health });
    }

    if (payload.action === 'updateDocumentStatus') {
      const documentsSheet = ensureSheet(ss, 'DEAL_DOCUMENTS', DEAL_DOCUMENTS_HEADERS);
      const document = updateDealDocumentStatus(documentsSheet, historySheet, payload);
      const health = syncDealHealth(ss, payload.deal_id || payload.dealId || payload.id);
      return jsonOutput({ success: true, document, health });
    }

    if (payload.action === 'syncDealHealth') {
      const health = syncDealHealth(ss, payload.deal_id || payload.dealId || payload.id);
      return jsonOutput({ success: true, health });
    }

    if (payload.action === 'addDealNote') {
      const notesSheet = ensureSheet(ss, 'DEAL_NOTES', DEAL_NOTES_HEADERS);
      const note = addDealNote(notesSheet, historySheet, payload);
      return jsonOutput({ success: true, note });
    }

    if (payload.action === 'saveDealMaster' || payload.action === 'save') {
      upsertDealMaster(masterSheet, historySheet, payload);
      const id = payload.deal_id || payload.dealId || payload.id;
      const health = syncDealHealth(ss, id);
      return jsonOutput({ success: true, id, health });
    }

    return jsonOutput({ success: false, error: 'Ação inválida' });
  } catch (err) {
    return jsonOutput({ success: false, error: err.message });
  }
}

function getCentralSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('CENTRAL_DB_ID');
  if (!id) throw new Error('CENTRAL_DB_ID não configurado nas Script Properties.');
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
  if (!dealId) throw new Error('deal_id obrigatório');
  if (!DEAL_MASTER_STATUS_FIELDS.includes(field)) throw new Error('Campo de status inválido: ' + field);

  const row = findDealMasterRow(masterSheet, dealId);
  if (row < 1) throw new Error('Deal não encontrado no DEAL_MASTER: ' + dealId);

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

  if (!dealId) throw new Error('deal_id obrigatório');
  if (!documentName) throw new Error('document_name obrigatório');
  if (!['printed_status', 'dealer_center_status'].includes(field)) {
    throw new Error('Campo de documento inválido: ' + field);
  }
  if (!['pending', 'ok', 'missing', ''].includes(value)) {
    throw new Error('Status de documento inválido: ' + value);
  }

  const row = findDealDocumentRow(documentsSheet, dealId, documentName);
  if (row < 1) throw new Error('Documento não encontrado no DEAL_DOCUMENTS: ' + documentName);

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

function syncDealHealth(ss, dealId) {
  if (!dealId) throw new Error('deal_id obrigatorio');

  const masterSheet = ensureSheet(ss, 'DEAL_MASTER', DEAL_MASTER_HEADERS);
  const documentsSheet = ensureSheet(ss, 'DEAL_DOCUMENTS', DEAL_DOCUMENTS_HEADERS);
  const alertsSheet = ensureSheet(ss, 'DEAL_ALERTS', DEAL_ALERTS_HEADERS);
  const historySheet = ensureSheet(ss, 'DEAL_HISTORY', DEAL_HISTORY_HEADERS);
  const row = findDealMasterRow(masterSheet, dealId);
  if (row < 1) throw new Error('Deal nao encontrado no DEAL_MASTER: ' + dealId);

  const deal = rowToObject(DEAL_MASTER_HEADERS, masterSheet.getRange(row, 1, 1, DEAL_MASTER_HEADERS.length).getValues()[0]);
  const documents = getRowsByDealId(documentsSheet, dealId);
  const now = new Date().toISOString();
  const changes = [];

  if (String(deal.dmv_status || '').toLowerCase() === 'done') {
    if (String(deal.hub_status || '').toLowerCase() !== 'completed') {
      setMasterField(masterSheet, row, 'hub_status', 'completed');
      deal.hub_status = 'completed';
      changes.push('hub_status = completed');
    }
    closeAlertsByTypes(alertsSheet, dealId, ['DMV nao enviado', 'DMV não enviado', 'DMV pending', 'DMV pendente'], now);
  }

  const documentStatus = calculateDealDocumentStatus(documents);
  if (documentStatus === 'complete' && String(deal.document_status || '').toLowerCase() !== 'complete') {
    setMasterField(masterSheet, row, 'document_status', 'complete');
    deal.document_status = 'complete';
    changes.push('document_status = complete');
  } else if (documentStatus !== 'complete' && String(deal.document_status || '').toLowerCase() === 'complete') {
    setMasterField(masterSheet, row, 'document_status', '');
    deal.document_status = '';
    changes.push('document_status = pending');
  }

  if (hasCriticalDocument(documents)) {
    ensureOpenAlert(alertsSheet, dealId, 'Documento critico', 'critical', 'open', now);
  }

  const aging = getAgingInfo(deal.sale_date);
  if (aging && aging.days > 60) {
    ensureOpenAlert(alertsSheet, dealId, 'Aging critico', 'critical', 'open', now);
  }

  const parcelado = String(deal.parcelado || '').toLowerCase();
  const parcelamentoStatus = String(deal.parcelamento_status || '').toLowerCase();
  if (['sim', 'yes', 'true'].includes(parcelado) && !['done', 'created', 'complete', 'completed', 'not_needed'].includes(parcelamentoStatus)) {
    ensureOpenAlert(alertsSheet, dealId, 'Parcelamento financeiro', 'high', 'open', now);
  }

  if (changes.length) {
    setMasterField(masterSheet, row, 'updated_at', now);
    appendHistory(historySheet, dealId, 'sync_deal_health', '', changes.join(', '));
  }

  const openAlerts = getRowsByDealId(alertsSheet, dealId).filter(alert => !isResolvedAlert(alert));
  return {
    health_score: calculateHealthScore(openAlerts, aging, documents),
    aging,
    document_status: deal.document_status || (documentStatus === 'complete' ? 'complete' : ''),
    document_complete: documentStatus === 'complete',
    alert_count: openAlerts.length,
    critical_alert_count: openAlerts.filter(isCriticalAlert).length,
  };
}

function calculateDealDocumentStatus(documents) {
  if (!documents.length) return '';
  return documents.every(document => {
    const printed = String(document.printed_status || '').toLowerCase();
    const dealerCenter = String(document.dealer_center_status || '').toLowerCase();
    const status = String(document.status || '').toLowerCase();
    return printed === 'ok' && dealerCenter === 'ok' && (status === 'ok' || status === 'complete');
  }) ? 'complete' : 'pending';
}

function hasCriticalDocument(documents) {
  return documents.some(document => {
    const status = String(document.status || '').toLowerCase();
    const printed = String(document.printed_status || '').toLowerCase();
    const dealerCenter = String(document.dealer_center_status || document.dealercenter_status || '').toLowerCase();
    return status === 'critical' || printed === 'missing' || dealerCenter === 'missing';
  });
}

function getAgingInfo(saleDate) {
  const date = toDateValue(saleDate);
  if (!date) return null;
  const today = new Date();
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const days = Math.max(0, Math.floor((end - start) / 86400000));
  const tone = days <= 21 ? 'green' : days <= 45 ? 'yellow' : days <= 60 ? 'orange' : 'red';
  return { days, tone };
}

function calculateHealthScore(openAlerts, aging, documents) {
  if (openAlerts.some(isCriticalAlert) || (aging && aging.days > 60) || hasCriticalDocument(documents)) return 'critical';
  if (openAlerts.length || (aging && aging.days > 21)) return 'attention';
  return 'ok';
}

function ensureOpenAlert(sheet, dealId, alertType, priority, status, now) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(dealId) && String(values[i][1]) === String(alertType)) {
      const alert = rowToObject(DEAL_ALERTS_HEADERS, values[i]);
      if (!isResolvedAlert(alert)) return;
    }
  }
  sheet.appendRow([dealId, alertType, priority || 'high', status || 'open', now || new Date().toISOString(), '']);
}

function closeAlertsByTypes(sheet, dealId, alertTypes, now) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const alert = rowToObject(DEAL_ALERTS_HEADERS, values[i]);
    if (String(alert.deal_id) === String(dealId) && alertTypes.includes(String(alert.alert_type)) && !isResolvedAlert(alert)) {
      sheet.getRange(i + 1, DEAL_ALERTS_HEADERS.indexOf('status') + 1).setValue('resolved');
      sheet.getRange(i + 1, DEAL_ALERTS_HEADERS.indexOf('resolved_at') + 1).setValue(now || new Date().toISOString());
    }
  }
}

function setMasterField(sheet, row, field, value) {
  const column = DEAL_MASTER_HEADERS.indexOf(field) + 1;
  if (column > 0) sheet.getRange(row, column).setValue(value);
}

function toDateValue(value) {
  if (!value) return null;
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value === 'number') {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const text = String(value).trim();
  if (!text) return null;
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
  match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) return new Date(Number(match[3]), Number(match[1]) - 1, Number(match[2]), 12);
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
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
  if (!dealId) throw new Error('deal_id obrigatório');

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
  return ['critical', 'critico', 'crítico', 'high', 'alta'].includes(String(alert.priority || '').toLowerCase());
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
