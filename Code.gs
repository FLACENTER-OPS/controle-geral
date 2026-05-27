const DEAL_MASTER_HEADERS = [
  'deal_id', 'stock', 'vin', 'vehicle', 'seller', 'customer_name', 'sale_date',
  'lender', 'deal_type', 'trade_in', 'gps_required', 'parcelado', 'out_of_state',
  'company_deal', 'review_status', 'commission_status', 'dmv_status',
  'parcelamento_status', 'title_status', 'omnia_status', 'folder_status',
  'envelope_status', 'created_at', 'updated_at'
];

const DEAL_ALERTS_HEADERS = [
  'deal_id', 'alert_type', 'priority', 'status', 'created_at', 'resolved_at'
];

const DEAL_HISTORY_HEADERS = [
  'deal_id', 'action', 'user', 'notes', 'timestamp'
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const payload = JSON.parse(e.postData.contents || '{}');
    const masterSheet = ensureSheet(ss, 'DEAL_MASTER', DEAL_MASTER_HEADERS);
    const historySheet = ensureSheet(ss, 'DEAL_HISTORY', DEAL_HISTORY_HEADERS);

    if (payload.action === 'updateStatus') {
      updateDealMasterStatus(masterSheet, historySheet, payload);
      return jsonOutput({ success: true });
    }

    if (payload.action === 'saveDealMaster' || payload.action === 'save') {
      upsertDealMaster(masterSheet, historySheet, payload);
      return jsonOutput({ success: true, id: payload.deal_id || payload.dealId || payload.id });
    }

    return jsonOutput({ success: false, error: 'Ação inválida' });
  } catch (err) {
    return jsonOutput({ success: false, error: err.message });
  }
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
  masterSheet.getRange(row, DEAL_MASTER_HEADERS.indexOf('updated_at') + 1).setValue(new Date().toISOString());
  appendHistory(historySheet, dealId, 'update_status', payload.user || '', field + ' = ' + (payload.value || ''));
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

function findDealMasterRow(sheet, dealId) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(dealId)) return i + 1;
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
