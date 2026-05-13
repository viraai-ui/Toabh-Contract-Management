const MAIN_SHEET_NAME = PropertiesService.getScriptProperties().getProperty('MAIN_SHEET_NAME') || 'Contract Links';
const RENEWALS_SHEET_NAME = PropertiesService.getScriptProperties().getProperty('RENEWALS_SHEET_NAME') || 'Renewals';
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
const CONTRACT_GENERATOR_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('CONTRACT_GENERATOR_WEBHOOK_URL') || '';
const ZOHO_SIGN_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('ZOHO_SIGN_WEBHOOK_URL') || '';
const ZOHO_STATUS_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('ZOHO_STATUS_WEBHOOK_URL') || '';

const MAIN_COLUMNS = [
  'Name', 'Email', 'Phone', 'Contract Link', 'Zoho Request ID', 'Zoho Status', 'Zoho Sent At', 'Zoho Error', 'Signed PDF URL', 'Version',
  'Contract Signed On', 'Contract Start Date', 'Contract Validity', 'Contract Expiry Date', 'Days Left',
  'AI Scan Status', 'AI Scan Notes', 'Renewal Status', 'Renewal Sheet Row ID', 'Last Synced At', 'Notes'
];

const RENEWAL_COLUMNS = [
  'Renewal ID', 'Original Contract Row ID', 'Name', 'Email', 'Phone', 'Old Version', 'New Version',
  'Old Contract Link', 'Old Signed PDF URL', 'Old Expiry Date', 'Renewal Status', 'Renewal Started On',
  'Editable Data JSON', 'New Contract Link', 'New Zoho Request ID', 'New Zoho Status',
  'New Signed PDF URL', 'New Contract Signed On', 'New Contract Expiry Date', 'Notes', 'Error'
];

const EDITABLE_DATA_HEADERS = ['Editable Data JSON', 'Editable Data JSON / Editable Fields'];

const RENEWAL_FINAL_STATUSES = {
  SIGNED: 'Signed/Renewed',
  RENEWED: 'Signed/Renewed',
  CANCELLED: 'Cancelled',
  ON_HOLD: 'On Hold',
  NOT_RENEWING: 'Not Renewing'
};

function setupSheetsAndTriggers() {
  ensureSheetColumns_();
  removeManagedTriggers_();
  ScriptApp.newTrigger('handleSheetEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger('runMaintenance').timeBased().everyHours(1).create();
}

function removeManagedTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'handleSheetEdit' || handler === 'runMaintenance') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function ensureSheetColumns_() {
  const ss = SpreadsheetApp.getActive();
  ensureHeaders_(ss.getSheetByName(MAIN_SHEET_NAME) || ss.insertSheet(MAIN_SHEET_NAME), MAIN_COLUMNS);
  ensureHeaders_(ss.getSheetByName(RENEWALS_SHEET_NAME) || ss.insertSheet(RENEWALS_SHEET_NAME), RENEWAL_COLUMNS);
}

function ensureHeaders_(sheet, headers) {
  const existing = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  headers.forEach((header, index) => {
    if (existing[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
  if (sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1);
}

function handleSheetEdit(e) {
  ensureSheetColumns_();
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  if (row === 1) return;

  if (sheet.getName() === MAIN_SHEET_NAME) {
    handleMainSheetEdit_(e, sheet, row);
    return;
  }

  if (sheet.getName() === RENEWALS_SHEET_NAME) {
    handleRenewalSheetEdit_(sheet, row);
  }
}

function handleMainSheetEdit_(e, sheet, row) {
  const map = getHeaderMap_(sheet);
  if (e.range.getColumn() !== map['Signed PDF URL']) return;

  const value = String(e.value || '').trim();
  if (!isGoogleDriveLink_(value)) return;

  const previousValue = String(e.oldValue || '').trim();
  if (!sheet.getRange(row, map['AI Scan Status']).getValue() || previousValue !== value) {
    sheet.getRange(row, map['AI Scan Status']).setValue('Pending');
    sheet.getRange(row, map['AI Scan Notes']).setValue(previousValue && previousValue !== value ? 'Signed PDF URL changed. Pending rescan.' : 'Signed PDF detected. Pending AI scan.');
  }
  scanPendingContracts({ onlyRow: row, force: false });
}

function handleRenewalSheetEdit_(sheet, rowId) {
  const map = getHeaderMap_(sheet);
  const row = getRowObject_(sheet, map, rowId);
  if (shouldFinalizeRenewal_(row)) {
    finalizeRenewalToMain_(rowId, { source: 'sheet_edit' });
  }
}

function runMaintenance() {
  ensureSheetColumns_();
  scanPendingContracts({ force: false });
  syncDerivedFields_();
  syncRenewalFinalizations_();
}

function doGet(e) {
  ensureSheetColumns_();
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || 'contracts');
  try {
    return handleGetAction_(action, params);
  } catch (error) {
    return jsonOutput_({ ok: false, message: error.message || 'Request failed.' });
  }
}

function doPost(e) {
  ensureSheetColumns_();
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const params = e && e.parameter ? e.parameter : {};
    const action = String(params.action || body.action || '');
    switch (action) {
      case 'add-contract':
        return jsonOutput_(callLegacyFunction_('addContract', body));
      case 'resend-contract':
        return jsonOutput_(callLegacyFunction_('resendContractApi', body));
      case 'rescanWithAI':
      case 'rescan-ai':
        return jsonOutput_(rescanWithAI_(Number(body.rowId)));
      case 'startRenewal':
      case 'start-renewal':
        return jsonOutput_(startRenewal_(Number(body.rowId)));
      case 'addNote':
      case 'update-note':
        return jsonOutput_(addNote_(Number(body.rowId), String(body.note || ''), String(body.scope || 'main')));
      case 'setContractRenewalStatus':
        return jsonOutput_(setContractRenewalStatus_(Number(body.rowId), String(body.status || ''), String(body.note || '')));
      case 'updateRenewal':
      case 'update-renewal':
        return jsonOutput_(updateRenewal_(Number(body.rowId), body.updates || {}));
      case 'regenerateRenewalContract':
      case 'generate-renewal-contract':
        return jsonOutput_(regenerateRenewalContract_(Number(body.rowId)));
      case 'sendRenewalForSigning':
      case 'send-renewal-for-signing':
        return jsonOutput_(sendRenewalForSigning_(Number(body.rowId)));
      case 'refreshRenewalZohoStatus':
      case 'refresh-renewal-status':
        return jsonOutput_(refreshRenewalZohoStatus_(Number(body.rowId)));
      case 'markRenewalSigned':
        return jsonOutput_(markRenewalSigned_(Number(body.rowId), body));
      case 'markRenewalRenewed':
        return jsonOutput_(markRenewalRenewed_(Number(body.rowId), body));
      case 'setRenewalOnHold':
        return jsonOutput_(setRenewalOnHold_(Number(body.rowId), String(body.note || '')));
      case 'setRenewalNotRenewing':
        return jsonOutput_(setRenewalNotRenewing_(Number(body.rowId), String(body.note || '')));
      case 'cancelRenewal':
        return jsonOutput_(cancelRenewal_(Number(body.rowId), String(body.note || '')));
      default:
        return jsonOutput_({ ok: false, message: 'Unsupported action.' });
    }
  } catch (error) {
    return jsonOutput_({ ok: false, message: error.message || 'Action failed.' });
  }
}

function handleGetAction_(action, params) {
  if (action === 'dashboard') {
    return jsonOutput_({
      ok: true,
      contracts: getSignedContracts_(params.expiringFilter, params.from, params.to),
      renewals: getRenewals_(),
      generatedAt: new Date().toISOString()
    });
  }

  if (action === 'contracts') {
    if (hasLegacyFunction_('getContracts')) return jsonOutput_(callLegacyFunction_('getContracts', params));
    return jsonOutput_({ ok: true, contracts: getContractsForApi_(params) });
  }

  if (action === 'documents') {
    if (hasLegacyFunction_('getDocuments')) return jsonOutput_(callLegacyFunction_('getDocuments', params));
    return jsonOutput_({ ok: true, documents: {} });
  }

  if (action === 'documents_all') {
    if (hasLegacyFunction_('getDocumentsAll')) return jsonOutput_(callLegacyFunction_('getDocumentsAll', params));
    return jsonOutput_({ ok: true, documents: [] });
  }

  if (action === 'renewals') {
    return jsonOutput_({ ok: true, renewals: getRenewals(), generatedAt: new Date().toISOString() });
  }

  return jsonOutput_({ ok: false, message: 'Unsupported action.' });
}

function hasLegacyFunction_(name) {
  return typeof this[name] === 'function';
}

function callLegacyFunction_(name, payload) {
  if (!hasLegacyFunction_(name)) {
    throw new Error('Missing legacy handler: ' + name);
  }
  return this[name](payload || {});
}

function getSignedContracts_(expiringFilter, from, to) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const rows = getRowsAsObjects_(sheet, map);

  return rows
    .filter((row) => isGoogleDriveLink_(row['Signed PDF URL']))
    .filter((row) => contractMatchesFilter_(row, expiringFilter, from, to))
    .map((row) => toContractRecord_(row, map));
}

function getContractsForApi_(params) {
  var signedOnly = String((params && params.signedOnly) || '').trim() === '1';
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  return getRowsAsObjects_(sheet, map)
    .filter((row) => !signedOnly || isGoogleDriveLink_(row['Signed PDF URL']))
    .map((row) => ({
      rowId: row.__rowId,
      name: row['Name'] || '',
      email: row['Email'] || '',
      phone: row['Phone'] || '',
      contractLink: row['Contract Link'] || '',
      zohoRequestId: row['Zoho Request ID'] || '',
      zohoStatus: row['Zoho Status'] || '',
      zohoSentAt: formatDateValue_(row['Zoho Sent At'], true),
      zohoError: row['Zoho Error'] || '',
      signedPdfUrl: row['Signed PDF URL'] || '',
      version: row['Version'] || '',
      contractSignedOn: formatDateValue_(row['Contract Signed On']),
      contractStartDate: formatDateValue_(row['Contract Start Date']),
      contractValidity: row['Contract Validity'] || '',
      contractExpiryDate: formatDateValue_(row['Contract Expiry Date']),
      daysLeft: valueToNumber_(row['Days Left']),
      aiScanStatus: row['AI Scan Status'] || '',
      aiScanNotes: row['AI Scan Notes'] || '',
      renewalStatus: row['Renewal Status'] || '',
      renewalSheetRowId: row['Renewal Sheet Row ID'] || '',
      lastSyncedAt: formatDateValue_(row['Last Synced At'], true),
      notes: row['Notes'] || ''
    }));
}

function getRenewals_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  return getRowsAsObjects_(sheet, map).map((row) => {
    const editable = parseEditableFields_(getEditableDataValue_(row));
    return {
      rowId: row.__rowId,
      renewalId: row['Renewal ID'] || '',
      originalContractRowId: row['Original Contract Row ID'] || '',
      name: row['Name'] || '',
      email: row['Email'] || '',
      phone: row['Phone'] || '',
      oldVersion: row['Old Version'] || '',
      newVersion: row['New Version'] || '',
      oldContractLink: row['Old Contract Link'] || '',
      oldSignedPdfUrl: row['Old Signed PDF URL'] || '',
      oldExpiryDate: formatDateValue_(row['Old Expiry Date']),
      renewalStatus: row['Renewal Status'] || '',
      renewalStartedOn: formatDateValue_(row['Renewal Started On']),
      editableDataJson: getEditableDataValue_(row),
      editableFields: editable,
      newContractLink: row['New Contract Link'] || '',
      newZohoRequestId: row['New Zoho Request ID'] || '',
      newZohoStatus: row['New Zoho Status'] || '',
      newSignedPdfUrl: row['New Signed PDF URL'] || '',
      newContractSignedOn: formatDateValue_(row['New Contract Signed On']),
      newContractExpiryDate: formatDateValue_(row['New Contract Expiry Date']),
      notes: row['Notes'] || '',
      error: row['Error'] || '',
      canEdit: !isRenewalFinalStatus_(row['Renewal Status']),
      isReadyForFinalization: shouldFinalizeRenewal_(row)
    };
  });
}

function getRenewals() {
  return getRenewals_();
}

function rescanWithAI_(rowId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  if (rowId < 2) throw new Error('Invalid row ID.');
  sheet.getRange(rowId, map['AI Scan Status']).setValue('Rescan Requested');
  sheet.getRange(rowId, map['AI Scan Notes']).setValue('Manual rescan requested');
  scanPendingContracts({ onlyRow: rowId, force: true });
  return { ok: true };
}

function addNote_(rowId, note, scope) {
  const isRenewal = scope === 'renewal';
  const sheet = SpreadsheetApp.getActive().getSheetByName(isRenewal ? RENEWALS_SHEET_NAME : MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const current = String(sheet.getRange(rowId, map['Notes']).getValue() || '').trim();
  const stamped = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const next = current ? current + '\n' + '[' + stamped + '] ' + note : '[' + stamped + '] ' + note;
  sheet.getRange(rowId, map['Notes']).setValue(next);
  if (!isRenewal && map['Last Synced At']) sheet.getRange(rowId, map['Last Synced At']).setValue(new Date());
  return { ok: true };
}

function setContractRenewalStatus_(rowId, status, note) {
  const allowed = ['On Hold', 'Not Renewing'];
  if (allowed.indexOf(status) === -1) throw new Error('Unsupported contract renewal status.');
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  sheet.getRange(rowId, map['Renewal Status']).setValue(status);
  if (note) addNote_(rowId, note, 'main');
  if (map['Last Synced At']) sheet.getRange(rowId, map['Last Synced At']).setValue(new Date());
  return { ok: true, rowId: rowId, status: status };
}

function startRenewal_(rowId) {
  const ss = SpreadsheetApp.getActive();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const renewalSheet = ss.getSheetByName(RENEWALS_SHEET_NAME);
  const mainMap = getHeaderMap_(mainSheet);
  const renewalMap = getHeaderMap_(renewalSheet);
  const row = getRowObject_(mainSheet, mainMap, rowId);
  if (!isGoogleDriveLink_(row['Signed PDF URL'])) throw new Error('Only signed contracts can start renewal.');

  const existingRenewalRowId = findActiveRenewalRowIdByOriginal_(rowId);
  if (existingRenewalRowId) {
    return { ok: true, renewalRowId: existingRenewalRowId, renewalId: renewalSheet.getRange(existingRenewalRowId, renewalMap['Renewal ID']).getValue(), reused: true };
  }

  const nextVersion = incrementVersion_(row['Version'] || 'V1');
  const renewalId = buildRenewalId_(renewalSheet);
  const editablePayload = JSON.stringify({
    name: row['Name'] || '',
    email: row['Email'] || '',
    phone: row['Phone'] || '',
    version: nextVersion,
    noKycRequired: true,
    contractStartDate: formatDateValue_(row['Contract Start Date']),
    contractValidity: row['Contract Validity'] || '',
    contractExpiryDate: formatDateValue_(row['Contract Expiry Date'])
  });

  const renewalRowId = renewalSheet.getLastRow() + 1;
  const renewalValues = blankRowForHeaders_(RENEWAL_COLUMNS);
  renewalValues[renewalMap['Renewal ID'] - 1] = renewalId;
  renewalValues[renewalMap['Original Contract Row ID'] - 1] = rowId;
  renewalValues[renewalMap['Name'] - 1] = row['Name'] || '';
  renewalValues[renewalMap['Email'] - 1] = row['Email'] || '';
  renewalValues[renewalMap['Phone'] - 1] = row['Phone'] || '';
  renewalValues[renewalMap['Old Version'] - 1] = row['Version'] || '';
  renewalValues[renewalMap['New Version'] - 1] = nextVersion;
  renewalValues[renewalMap['Old Contract Link'] - 1] = row['Contract Link'] || '';
  renewalValues[renewalMap['Old Signed PDF URL'] - 1] = row['Signed PDF URL'] || '';
  renewalValues[renewalMap['Old Expiry Date'] - 1] = row['Contract Expiry Date'] || '';
  renewalValues[renewalMap['Renewal Status'] - 1] = 'Draft Created';
  renewalValues[renewalMap['Renewal Started On'] - 1] = new Date();
  renewalValues[getEditableDataColumnIndex_(renewalMap) - 1] = editablePayload;
  renewalSheet.getRange(renewalRowId, 1, 1, renewalValues.length).setValues([renewalValues]);

  updateMainSheetRenewalState_(rowId, 'Draft Created', renewalRowId);
  triggerContractGenerator_(renewalRowId, false);
  return { ok: true, renewalRowId: renewalRowId, renewalId: renewalId };
}

function updateRenewal_(rowId, updates) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  validateRenewalRowId_(rowId);
  const row = getRowObject_(sheet, map, rowId);
  if (isRenewalFinalStatus_(row['Renewal Status'])) throw new Error('Finalized renewals cannot be edited.');

  const editable = parseEditableFields_(getEditableDataValue_(row));
  const normalized = updates || {};
  const fields = ['name', 'email', 'phone', 'version', 'contractStartDate', 'contractValidity', 'contractExpiryDate', 'noKycRequired'];
  fields.forEach((field) => {
    if (Object.prototype.hasOwnProperty.call(normalized, field)) {
      editable[field] = normalized[field];
    }
  });

  if (Object.prototype.hasOwnProperty.call(normalized, 'name')) sheet.getRange(rowId, map['Name']).setValue(String(normalized.name || ''));
  if (Object.prototype.hasOwnProperty.call(normalized, 'email')) sheet.getRange(rowId, map['Email']).setValue(String(normalized.email || ''));
  if (Object.prototype.hasOwnProperty.call(normalized, 'phone')) sheet.getRange(rowId, map['Phone']).setValue(String(normalized.phone || ''));
  if (Object.prototype.hasOwnProperty.call(normalized, 'version')) sheet.getRange(rowId, map['New Version']).setValue(String(normalized.version || ''));

  sheet.getRange(rowId, getEditableDataColumnIndex_(map)).setValue(JSON.stringify(editable));
  sheet.getRange(rowId, map['Renewal Status']).setValue('Edited');
  clearRenewalError_(sheet, map, rowId);
  setRenewalNote_(sheet, map, rowId, 'Renewal draft updated from dashboard.');
  updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), 'Renewal Started', rowId);
  return { ok: true, rowId: rowId };
}

function regenerateRenewalContract_(rowId) {
  validateRenewalRowId_(rowId);
  triggerContractGenerator_(rowId, true);
  return { ok: true, rowId: rowId };
}

function triggerContractGenerator_(renewalRowId, isRegeneration) {
  const renewalSheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(renewalSheet);
  const row = getRowObject_(renewalSheet, map, renewalRowId);
  if (!CONTRACT_GENERATOR_WEBHOOK_URL) {
    setRenewalError_(renewalSheet, map, renewalRowId, 'Missing CONTRACT_GENERATOR_WEBHOOK_URL script property.');
    return;
  }

  const editableData = parseEditableFields_(getEditableDataValue_(row));
  renewalSheet.getRange(renewalRowId, map['Renewal Status']).setValue(isRegeneration ? 'Regenerating Contract' : 'Generating Contract');
  clearRenewalError_(renewalSheet, map, renewalRowId);

  try {
    const response = UrlFetchApp.fetch(CONTRACT_GENERATOR_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        renewalId: row['Renewal ID'],
        originalContractRowId: row['Original Contract Row ID'],
        name: editableData.name || row['Name'],
        email: editableData.email || row['Email'],
        phone: editableData.phone || row['Phone'],
        oldVersion: row['Old Version'],
        newVersion: editableData.version || row['New Version'],
        noKycRequired: editableData.noKycRequired !== false,
        editableData: editableData,
        regenerate: !!isRegeneration
      })
    });

    const payload = safeJsonParse_(response.getContentText()) || {};
    const status = payload.status || 'Contract Generated';
    renewalSheet.getRange(renewalRowId, map['Renewal Status']).setValue(status);
    if (payload.newContractLink) renewalSheet.getRange(renewalRowId, map['New Contract Link']).setValue(payload.newContractLink);
    if (payload.newZohoRequestId) renewalSheet.getRange(renewalRowId, map['New Zoho Request ID']).setValue(payload.newZohoRequestId);
    if (payload.newZohoStatus) renewalSheet.getRange(renewalRowId, map['New Zoho Status']).setValue(payload.newZohoStatus);
    if (payload.newSignedPdfUrl) renewalSheet.getRange(renewalRowId, map['New Signed PDF URL']).setValue(payload.newSignedPdfUrl);
    if (payload.newContractSignedOn) renewalSheet.getRange(renewalRowId, map['New Contract Signed On']).setValue(parseDateValue_(payload.newContractSignedOn) || payload.newContractSignedOn);
    if (payload.newContractExpiryDate) renewalSheet.getRange(renewalRowId, map['New Contract Expiry Date']).setValue(parseDateValue_(payload.newContractExpiryDate) || payload.newContractExpiryDate);
    setRenewalNote_(renewalSheet, map, renewalRowId, isRegeneration ? 'Renewal contract regenerated.' : 'Renewal contract generated.');
    updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), status, renewalRowId);
    maybeFinalizeRenewal_(renewalRowId, { source: 'generator' });
  } catch (error) {
    renewalSheet.getRange(renewalRowId, map['Renewal Status']).setValue('Failed');
    setRenewalError_(renewalSheet, map, renewalRowId, error.message || 'Generator trigger failed');
    updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), 'Failed', renewalRowId);
  }
}

function sendRenewalForSigning_(rowId) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const row = getRowObject_(sheet, map, rowId);
  if (!row['New Contract Link']) throw new Error('Generate the renewal contract before sending for signing.');
  if (!ZOHO_SIGN_WEBHOOK_URL) throw new Error('Missing ZOHO_SIGN_WEBHOOK_URL script property.');

  sheet.getRange(rowId, map['Renewal Status']).setValue('Sending For Signing');
  clearRenewalError_(sheet, map, rowId);

  try {
    const response = UrlFetchApp.fetch(ZOHO_SIGN_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        renewalId: row['Renewal ID'],
        originalContractRowId: row['Original Contract Row ID'],
        name: row['Name'],
        email: row['Email'],
        phone: row['Phone'],
        newContractLink: row['New Contract Link'],
        newVersion: row['New Version'],
        newZohoRequestId: row['New Zoho Request ID'] || ''
      })
    });
    const payload = safeJsonParse_(response.getContentText()) || {};
    if (payload.newZohoRequestId) sheet.getRange(rowId, map['New Zoho Request ID']).setValue(payload.newZohoRequestId);
    sheet.getRange(rowId, map['New Zoho Status']).setValue(payload.newZohoStatus || 'Sent for Signature');
    sheet.getRange(rowId, map['Renewal Status']).setValue(payload.status || 'Sent for Signature');
    setRenewalNote_(sheet, map, rowId, 'Renewal sent for signing.');
    updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), payload.status || 'Renewal Started', rowId);
    maybeFinalizeRenewal_(rowId, { source: 'send_for_signing' });
    return { ok: true, rowId: rowId, zohoRequestId: payload.newZohoRequestId || row['New Zoho Request ID'] || '' };
  } catch (error) {
    sheet.getRange(rowId, map['Renewal Status']).setValue('Send Failed');
    setRenewalError_(sheet, map, rowId, error.message || 'Zoho send failed');
    throw error;
  }
}

function refreshRenewalZohoStatus_(rowId) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const row = getRowObject_(sheet, map, rowId);
  if (!row['New Zoho Request ID']) throw new Error('Missing Zoho request ID for this renewal.');
  if (!ZOHO_STATUS_WEBHOOK_URL) throw new Error('Missing ZOHO_STATUS_WEBHOOK_URL script property.');

  try {
    const response = UrlFetchApp.fetch(ZOHO_STATUS_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        renewalId: row['Renewal ID'],
        requestId: row['New Zoho Request ID']
      })
    });
    const payload = safeJsonParse_(response.getContentText()) || {};
    if (payload.newZohoStatus) sheet.getRange(rowId, map['New Zoho Status']).setValue(payload.newZohoStatus);
    if (payload.newSignedPdfUrl) sheet.getRange(rowId, map['New Signed PDF URL']).setValue(payload.newSignedPdfUrl);
    if (payload.newContractSignedOn) sheet.getRange(rowId, map['New Contract Signed On']).setValue(parseDateValue_(payload.newContractSignedOn) || payload.newContractSignedOn);
    if (payload.newContractExpiryDate) sheet.getRange(rowId, map['New Contract Expiry Date']).setValue(parseDateValue_(payload.newContractExpiryDate) || payload.newContractExpiryDate);
    if (payload.status) sheet.getRange(rowId, map['Renewal Status']).setValue(payload.status);
    setRenewalNote_(sheet, map, rowId, 'Zoho status refreshed.');
    maybeFinalizeRenewal_(rowId, { source: 'zoho_refresh' });
    return { ok: true, rowId: rowId };
  } catch (error) {
    setRenewalError_(sheet, map, rowId, error.message || 'Zoho status refresh failed');
    throw error;
  }
}

function markRenewalSigned_(rowId, body) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  if (body && body.newSignedPdfUrl) sheet.getRange(rowId, map['New Signed PDF URL']).setValue(String(body.newSignedPdfUrl));
  if (body && body.newContractSignedOn) sheet.getRange(rowId, map['New Contract Signed On']).setValue(parseDateValue_(body.newContractSignedOn) || body.newContractSignedOn);
  if (body && body.newContractExpiryDate) sheet.getRange(rowId, map['New Contract Expiry Date']).setValue(parseDateValue_(body.newContractExpiryDate) || body.newContractExpiryDate);
  sheet.getRange(rowId, map['New Zoho Status']).setValue('Signed');
  sheet.getRange(rowId, map['Renewal Status']).setValue(RENEWAL_FINAL_STATUSES.SIGNED);
  setRenewalNote_(sheet, map, rowId, 'Marked signed from dashboard.');
  finalizeRenewalToMain_(rowId, { source: 'manual_signed' });
  return { ok: true, rowId: rowId };
}

function markRenewalRenewed_(rowId, body) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  if (body && body.newSignedPdfUrl) sheet.getRange(rowId, map['New Signed PDF URL']).setValue(String(body.newSignedPdfUrl));
  if (body && body.newContractSignedOn) sheet.getRange(rowId, map['New Contract Signed On']).setValue(parseDateValue_(body.newContractSignedOn) || body.newContractSignedOn);
  if (body && body.newContractExpiryDate) sheet.getRange(rowId, map['New Contract Expiry Date']).setValue(parseDateValue_(body.newContractExpiryDate) || body.newContractExpiryDate);
  sheet.getRange(rowId, map['Renewal Status']).setValue(RENEWAL_FINAL_STATUSES.RENEWED);
  setRenewalNote_(sheet, map, rowId, 'Marked renewed from dashboard.');
  finalizeRenewalToMain_(rowId, { source: 'manual_renewed' });
  return { ok: true, rowId: rowId };
}

function setRenewalOnHold_(rowId, note) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  sheet.getRange(rowId, map['Renewal Status']).setValue(RENEWAL_FINAL_STATUSES.ON_HOLD);
  if (note) setRenewalNote_(sheet, map, rowId, note);
  const row = getRowObject_(sheet, map, rowId);
  updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), RENEWAL_FINAL_STATUSES.ON_HOLD, rowId);
  return { ok: true, rowId: rowId };
}

function setRenewalNotRenewing_(rowId, note) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  sheet.getRange(rowId, map['Renewal Status']).setValue(RENEWAL_FINAL_STATUSES.NOT_RENEWING);
  if (note) setRenewalNote_(sheet, map, rowId, note);
  const row = getRowObject_(sheet, map, rowId);
  updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), RENEWAL_FINAL_STATUSES.NOT_RENEWING, rowId);
  return { ok: true, rowId: rowId };
}

function cancelRenewal_(rowId, note) {
  validateRenewalRowId_(rowId);
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  sheet.getRange(rowId, map['Renewal Status']).setValue(RENEWAL_FINAL_STATUSES.CANCELLED);
  if (note) setRenewalNote_(sheet, map, rowId, note);
  const row = getRowObject_(sheet, map, rowId);
  updateMainSheetRenewalState_(Number(row['Original Contract Row ID']), RENEWAL_FINAL_STATUSES.CANCELLED, rowId);
  return { ok: true, rowId: rowId };
}

function scanPendingContracts(options) {
  const force = options && options.force;
  const onlyRow = options && options.onlyRow;
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const rows = getRowsAsObjects_(sheet, map)
    .filter((row) => !onlyRow || row.__rowId === onlyRow)
    .filter((row) => isGoogleDriveLink_(row['Signed PDF URL']))
    .filter((row) => {
      const status = String(row['AI Scan Status'] || '').trim();
      return force || !status || status === 'Pending' || status === 'Failed' || status === 'Rescan Requested';
    });

  rows.forEach((row) => {
    try {
      const scan = scanContractPdf_(String(row['Signed PDF URL']));
      const startDate = parseDateValue_(scan.contractStartDate);
      const signedOn = parseDateValue_(scan.contractSignedOn);
      let expiryDate = parseDateValue_(scan.contractExpiryDate);
      if (!expiryDate && startDate && scan.contractValidity) {
        expiryDate = deriveExpiryDateFromValidity_(startDate, scan.contractValidity);
      }
      const daysLeft = expiryDate ? calculateDaysLeft_(expiryDate) : '';
      const notes = [];
      if (scan.mismatchOrMissingInfo) notes.push(scan.mismatchOrMissingInfo);
      if (!signedOn) notes.push('Contract signed date not confidently found.');
      if (!startDate) notes.push('Contract start/effective date not confidently found.');
      if (!scan.contractValidity) notes.push('Contract validity not confidently found.');
      if (!expiryDate) notes.push('Contract expiry date could not be derived.');
      if (scan.talentName && row['Name'] && normalizeText_(scan.talentName) !== normalizeText_(row['Name'])) {
        notes.push('Talent name mismatch: sheet=' + row['Name'] + ' pdf=' + scan.talentName);
      }

      const writeResult = writeAiFields_(sheet, map, row.__rowId, {
        contractSignedOn: signedOn,
        contractStartDate: startDate,
        contractValidity: scan.contractValidity || '',
        contractExpiryDate: expiryDate,
        daysLeft: daysLeft,
      }, force, notes);

      if (writeResult.mismatches.length) {
        Array.prototype.push.apply(notes, writeResult.mismatches);
      }

      if (!expiryDate) {
        sheet.getRange(row.__rowId, map['AI Scan Status']).setValue('Failed');
        sheet.getRange(row.__rowId, map['AI Scan Notes']).setValue(notes.join(' | ') || 'Contract expiry date could not be confidently calculated.');
        sheet.getRange(row.__rowId, map['Renewal Status']).setValue(deriveRenewalStatus_(valueToNumber_(row['Days Left']), row['Renewal Status']));
        sheet.getRange(row.__rowId, map['Last Synced At']).setValue(new Date());
        return;
      }
      sheet.getRange(row.__rowId, map['AI Scan Status']).setValue('Scanned');
      sheet.getRange(row.__rowId, map['AI Scan Notes']).setValue(notes.join(' | '));
      sheet.getRange(row.__rowId, map['Renewal Status']).setValue(deriveRenewalStatus_(daysLeft, row['Renewal Status']));
      sheet.getRange(row.__rowId, map['Last Synced At']).setValue(new Date());
    } catch (error) {
      sheet.getRange(row.__rowId, map['AI Scan Status']).setValue('Failed');
      sheet.getRange(row.__rowId, map['AI Scan Notes']).setValue(error.message || 'Unknown scan error');
      sheet.getRange(row.__rowId, map['Renewal Status']).setValue(deriveRenewalStatus_(valueToNumber_(row['Days Left']), row['Renewal Status']));
      sheet.getRange(row.__rowId, map['Last Synced At']).setValue(new Date());
    }
  });
}

function scanContractPdf_(signedPdfUrl) {
  if (!GEMINI_API_KEY) throw new Error('Missing GEMINI_API_KEY in script properties.');
  const fileId = extractDriveFileId_(signedPdfUrl);
  if (!fileId) throw new Error('Invalid Google Drive file URL.');

  const blob = DriveApp.getFileById(fileId).getBlob();
  const requestBody = {
    contents: [{
      role: 'user',
      parts: [
        {
          text: 'Read this signed contract PDF. Return only strict JSON with these exact keys: talentName, contractSignedOn, contractStartDate, contractValidity, contractExpiryDate, mismatchOrMissingInfo. Rules: contractSignedOn should be the actual signing/completion date if visible in the signed contract or certificate; contractStartDate should be the agreement/effective/start date mentioned in the contract; contractValidity should be the exact validity term text like "3 years" if visible; contractExpiryDate should be the final expiry date if explicitly visible, otherwise leave empty; mismatchOrMissingInfo should briefly note uncertainty, missing dates, unreadable pages, or mismatch warnings. Use ISO YYYY-MM-DD whenever possible. Do not include markdown or extra text.'
        },
        {
          inlineData: {
            mimeType: blob.getContentType() || 'application/pdf',
            data: Utilities.base64Encode(blob.getBytes())
          }
        }
      ]
    }],
    generationConfig: {
      responseMimeType: 'application/json'
    }
  };

  const response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=' + encodeURIComponent(GEMINI_API_KEY), {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(requestBody)
  });

  const payload = JSON.parse(response.getContentText());
  const text = payload && payload.candidates && payload.candidates[0] && payload.candidates[0].content && payload.candidates[0].content.parts && payload.candidates[0].content.parts[0] && payload.candidates[0].content.parts[0].text;
  if (!text) throw new Error('Empty AI scan response.');
  return JSON.parse(text);
}

function syncDerivedFields_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  getRowsAsObjects_(sheet, map).forEach((row) => {
    const expiryDate = parseDateValue_(row['Contract Expiry Date']);
    const daysLeft = expiryDate ? calculateDaysLeft_(expiryDate) : '';
    sheet.getRange(row.__rowId, map['Days Left']).setValue(daysLeft);
    sheet.getRange(row.__rowId, map['Renewal Status']).setValue(deriveRenewalStatus_(daysLeft === '' ? valueToNumber_(row['Days Left']) : daysLeft, row['Renewal Status']));
  });
}

function syncRenewalFinalizations_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  getRowsAsObjects_(sheet, map).forEach((row) => {
    if (shouldFinalizeRenewal_(row)) {
      finalizeRenewalToMain_(row.__rowId, { source: 'maintenance' });
    }
  });
}

function maybeFinalizeRenewal_(rowId, context) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const row = getRowObject_(sheet, map, rowId);
  if (shouldFinalizeRenewal_(row)) {
    finalizeRenewalToMain_(rowId, context || {});
  }
}

function shouldFinalizeRenewal_(row) {
  const status = String(row['Renewal Status'] || '').trim();
  return isGoogleDriveLink_(row['New Signed PDF URL']) || status === RENEWAL_FINAL_STATUSES.SIGNED || status === RENEWAL_FINAL_STATUSES.RENEWED;
}

function finalizeRenewalToMain_(renewalRowId, context) {
  const ss = SpreadsheetApp.getActive();
  const renewalSheet = ss.getSheetByName(RENEWALS_SHEET_NAME);
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const renewalMap = getHeaderMap_(renewalSheet);
  const mainMap = getHeaderMap_(mainSheet);
  const renewalRow = getRowObject_(renewalSheet, renewalMap, renewalRowId);
  const originalRowId = Number(renewalRow['Original Contract Row ID']);
  if (originalRowId < 2) throw new Error('Renewal is missing original contract row ID.');

  const editable = parseEditableFields_(getEditableDataValue_(renewalRow));
  const status = isGoogleDriveLink_(renewalRow['New Signed PDF URL']) ? RENEWAL_FINAL_STATUSES.SIGNED : String(renewalRow['Renewal Status'] || '').trim() || RENEWAL_FINAL_STATUSES.RENEWED;
  const targetRowId = findOrCreateMainRowForRenewal_(mainSheet, mainMap, renewalRow, editable);

  mainSheet.getRange(targetRowId, mainMap['Name']).setValue(editable.name || renewalRow['Name'] || '');
  mainSheet.getRange(targetRowId, mainMap['Email']).setValue(editable.email || renewalRow['Email'] || '');
  mainSheet.getRange(targetRowId, mainMap['Phone']).setValue(editable.phone || renewalRow['Phone'] || '');
  mainSheet.getRange(targetRowId, mainMap['Contract Link']).setValue(renewalRow['New Contract Link'] || '');
  if (mainMap['Zoho Request ID']) mainSheet.getRange(targetRowId, mainMap['Zoho Request ID']).setValue(renewalRow['New Zoho Request ID'] || '');
  if (mainMap['Zoho Status']) mainSheet.getRange(targetRowId, mainMap['Zoho Status']).setValue(renewalRow['New Zoho Status'] || status);
  if (mainMap['Zoho Error']) mainSheet.getRange(targetRowId, mainMap['Zoho Error']).setValue('');
  mainSheet.getRange(targetRowId, mainMap['Signed PDF URL']).setValue(renewalRow['New Signed PDF URL'] || '');
  mainSheet.getRange(targetRowId, mainMap['Version']).setValue(editable.version || renewalRow['New Version'] || '');
  mainSheet.getRange(targetRowId, mainMap['Contract Signed On']).setValue(parseDateValue_(renewalRow['New Contract Signed On']) || '');
  mainSheet.getRange(targetRowId, mainMap['Contract Start Date']).setValue(parseDateValue_(editable.contractStartDate) || '');
  mainSheet.getRange(targetRowId, mainMap['Contract Validity']).setValue(editable.contractValidity || '');
  mainSheet.getRange(targetRowId, mainMap['Contract Expiry Date']).setValue(parseDateValue_(renewalRow['New Contract Expiry Date'] || editable.contractExpiryDate) || '');
  mainSheet.getRange(targetRowId, mainMap['AI Scan Status']).setValue('Pending');
  mainSheet.getRange(targetRowId, mainMap['AI Scan Notes']).setValue('Pending AI scan after renewal finalization');
  mainSheet.getRange(targetRowId, mainMap['Renewal Status']).setValue(status);
  mainSheet.getRange(targetRowId, mainMap['Renewal Sheet Row ID']).setValue(String(renewalRowId));
  mainSheet.getRange(targetRowId, mainMap['Last Synced At']).setValue(new Date());
  setMainNotesForRenewal_(mainSheet, mainMap, targetRowId, renewalRow, context);

  renewalSheet.getRange(renewalRowId, renewalMap['Renewal Status']).setValue(RENEWAL_FINAL_STATUSES.RENEWED);
  clearRenewalError_(renewalSheet, renewalMap, renewalRowId);
  setRenewalNote_(renewalSheet, renewalMap, renewalRowId, 'Final contract synced to main Contracts sheet.');
  updateMainSheetRenewalState_(originalRowId, renewalSheet.getRange(renewalRowId, renewalMap['Renewal Status']).getValue(), renewalRowId);
  scanPendingContracts({ onlyRow: targetRowId, force: true });
  return { ok: true, mainRowId: targetRowId, renewalRowId: renewalRowId };
}

function findOrCreateMainRowForRenewal_(mainSheet, mainMap, renewalRow, editable) {
  const renewalRowId = String(renewalRow.__rowId);
  const existing = getRowsAsObjects_(mainSheet, mainMap).find((row) => String(row['Renewal Sheet Row ID'] || '') === renewalRowId);
  if (existing) return existing.__rowId;
  const originalRowId = Number(renewalRow['Original Contract Row ID']);
  const originalRow = getRowObject_(mainSheet, mainMap, originalRowId);
  if (!isGoogleDriveLink_(originalRow['Signed PDF URL'])) {
    return originalRowId;
  }

  const newRowId = mainSheet.getLastRow() + 1;
  const rowValues = blankRowForHeaders_(MAIN_COLUMNS);
  rowValues[mainMap['Name'] - 1] = editable.name || renewalRow['Name'] || '';
  rowValues[mainMap['Email'] - 1] = editable.email || renewalRow['Email'] || '';
  rowValues[mainMap['Phone'] - 1] = editable.phone || renewalRow['Phone'] || '';
  rowValues[mainMap['Renewal Sheet Row ID'] - 1] = renewalRow.__rowId;
  mainSheet.getRange(newRowId, 1, 1, rowValues.length).setValues([rowValues]);
  return newRowId;
}

function updateMainSheetRenewalState_(rowId, status, renewalRowId) {
  if (!rowId || rowId < 2) return;
  const sheet = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  sheet.getRange(rowId, map['Renewal Status']).setValue(status || '');
  if (renewalRowId) sheet.getRange(rowId, map['Renewal Sheet Row ID']).setValue(String(renewalRowId));
  sheet.getRange(rowId, map['Last Synced At']).setValue(new Date());
}

function setMainNotesForRenewal_(sheet, map, rowId, renewalRow, context) {
  const source = context && context.source ? String(context.source) : 'renewal';
  const note = 'Renewal synced from row ' + renewalRow.__rowId + ' via ' + source + '.';
  const current = String(sheet.getRange(rowId, map['Notes']).getValue() || '').trim();
  const stamped = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const next = current ? current + '\n' + '[' + stamped + '] ' + note : '[' + stamped + '] ' + note;
  sheet.getRange(rowId, map['Notes']).setValue(next);
}

function findActiveRenewalRowIdByOriginal_(originalRowId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RENEWALS_SHEET_NAME);
  const map = getHeaderMap_(sheet);
  const match = getRowsAsObjects_(sheet, map).find((row) => Number(row['Original Contract Row ID']) === Number(originalRowId) && !isRenewalFinalStatus_(row['Renewal Status']));
  return match ? match.__rowId : 0;
}

function parseEditableFields_(value) {
  const parsed = safeJsonParse_(value);
  if (parsed && typeof parsed === 'object') return parsed;
  return {};
}

function getEditableDataColumnIndex_(map) {
  for (var i = 0; i < EDITABLE_DATA_HEADERS.length; i++) {
    if (map[EDITABLE_DATA_HEADERS[i]]) return map[EDITABLE_DATA_HEADERS[i]];
  }
  throw new Error('Missing editable data column in Renewals sheet.');
}

function getEditableDataValue_(row) {
  for (var i = 0; i < EDITABLE_DATA_HEADERS.length; i++) {
    var header = EDITABLE_DATA_HEADERS[i];
    if (Object.prototype.hasOwnProperty.call(row, header) && row[header]) return row[header];
  }
  return '';
}

function validateRenewalRowId_(rowId) {
  if (!rowId || rowId < 2) throw new Error('Invalid renewal row ID.');
}

function isRenewalFinalStatus_(status) {
  const text = String(status || '').trim();
  return text === RENEWAL_FINAL_STATUSES.SIGNED || text === RENEWAL_FINAL_STATUSES.RENEWED || text === RENEWAL_FINAL_STATUSES.CANCELLED || text === RENEWAL_FINAL_STATUSES.NOT_RENEWING;
}

function setRenewalNote_(sheet, map, rowId, note) {
  if (!note) return;
  const current = String(sheet.getRange(rowId, map['Notes']).getValue() || '').trim();
  const stamped = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const next = current ? current + '\n' + '[' + stamped + '] ' + note : '[' + stamped + '] ' + note;
  sheet.getRange(rowId, map['Notes']).setValue(next);
}

function setRenewalError_(sheet, map, rowId, errorText) {
  sheet.getRange(rowId, map['Error']).setValue(errorText || 'Unknown error');
}

function clearRenewalError_(sheet, map, rowId) {
  sheet.getRange(rowId, map['Error']).setValue('');
}

function toContractRecord_(row) {
  const expiryDate = parseDateValue_(row['Contract Expiry Date']);
  const daysLeft = expiryDate ? calculateDaysLeft_(expiryDate) : valueToNumber_(row['Days Left']);
  return {
    rowId: row.__rowId,
    name: row['Name'] || '',
    email: row['Email'] || '',
    phone: row['Phone'] || '',
    contractLink: row['Contract Link'] || '',
    signedPdfUrl: row['Signed PDF URL'] || '',
    version: row['Version'] || '',
    contractSignedOn: formatDateValue_(row['Contract Signed On']),
    contractStartDate: formatDateValue_(row['Contract Start Date']),
    contractValidity: row['Contract Validity'] || '',
    contractExpiryDate: formatDateValue_(row['Contract Expiry Date']),
    daysLeft: daysLeft,
    aiScanStatus: row['AI Scan Status'] || '',
    aiScanNotes: row['AI Scan Notes'] || '',
    renewalStatus: row['Renewal Status'] || '',
    renewalSheetRowId: row['Renewal Sheet Row ID'] || '',
    lastSyncedAt: formatDateValue_(row['Last Synced At'], true),
    notes: row['Notes'] || '',
    statusTone: deriveStatusTone_(daysLeft)
  };
}

function getRowsAsObjects_(sheet, map) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values.map((rowValues, index) => {
    const row = { __rowId: index + 2 };
    Object.keys(map).forEach((header) => {
      row[header] = rowValues[map[header] - 1];
    });
    return row;
  });
}

function getRowObject_(sheet, map, rowId) {
  const values = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = { __rowId: rowId };
  Object.keys(map).forEach((header) => {
    row[header] = values[map[header] - 1];
  });
  return row;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    map[String(header).trim()] = index + 1;
  });
  return map;
}

function contractMatchesFilter_(row, filter, from, to) {
  const expiry = parseDateValue_(row['Contract Expiry Date']);
  if (!filter || filter === 'all') return true;
  if (!expiry) return filter === 'already_expired' ? false : true;

  const today = stripTime_(new Date());
  const range = buildDateRange_(filter, from, to, today);
  if (!range) return true;
  if (range.kind === 'expired') return expiry < today;
  return expiry >= range.start && expiry <= range.end;
}

function buildDateRange_(filter, from, to, today) {
  const year = today.getFullYear();
  const month = today.getMonth();
  if (filter === 'already_expired') return { kind: 'expired' };
  if (filter === 'this_month') return { start: new Date(year, month, 1), end: new Date(year, month + 1, 0) };
  if (filter === 'next_month') return { start: new Date(year, month + 1, 1), end: new Date(year, month + 2, 0) };
  if (filter === 'next_3_months') return { start: today, end: new Date(year, month + 3, 0) };
  if (filter === 'next_6_months') return { start: today, end: new Date(year, month + 6, 0) };
  if (filter === 'custom' && from && to) return { start: stripTime_(new Date(from)), end: stripTime_(new Date(to)) };
  return null;
}

function deriveStatusTone_(daysLeft) {
  if (daysLeft === null || daysLeft === '' || typeof daysLeft !== 'number' || isNaN(daysLeft)) return 'unknown';
  if (daysLeft < 0) return 'expired';
  if (daysLeft < 30) return 'urgent';
  if (daysLeft <= 90) return 'due-soon';
  return 'active';
}

function deriveRenewalStatus_(daysLeft, currentStatus) {
  const status = String(currentStatus || '').trim();
  if (status === 'Renewal Started' || status === 'Signed/Renewed' || status === 'On Hold' || status === 'Not Renewing') return status;
  if (daysLeft === null || daysLeft === '' || typeof daysLeft !== 'number' || isNaN(daysLeft)) return status || '';
  if (daysLeft < 0) return 'Expired';
  if (daysLeft <= 90) return 'Due Soon';
  return 'Not Due';
}

function writeAiFields_(sheet, map, rowId, values, force, notes) {
  const mismatches = [];
  setAiField_(sheet, map, rowId, 'Contract Signed On', values.contractSignedOn, force, mismatches, notes);
  setAiField_(sheet, map, rowId, 'Contract Start Date', values.contractStartDate, force, mismatches, notes);
  setAiField_(sheet, map, rowId, 'Contract Validity', values.contractValidity, force, mismatches, notes);
  setAiField_(sheet, map, rowId, 'Contract Expiry Date', values.contractExpiryDate, force, mismatches, notes);
  setAiField_(sheet, map, rowId, 'Days Left', values.daysLeft, force, mismatches, notes);
  return { mismatches: mismatches };
}

function setAiField_(sheet, map, rowId, header, nextValue, force, mismatches) {
  var column = map[header];
  if (!column) return;
  var currentValue = sheet.getRange(rowId, column).getValue();
  var normalizedCurrent = normalizeComparableValue_(currentValue);
  var normalizedNext = normalizeComparableValue_(nextValue);
  if (!force && normalizedCurrent && normalizedNext && normalizedCurrent !== normalizedNext) {
    mismatches.push(header + ' mismatch: sheet=' + normalizedCurrent + ' ai=' + normalizedNext + '. Kept manual value.');
    return;
  }
  if (!force && normalizedCurrent && !normalizedNext) return;
  sheet.getRange(rowId, column).setValue(nextValue || '');
}

function normalizeComparableValue_(value) {
  if (value === null || value === undefined || value === '') return '';
  var parsed = parseDateValue_(value);
  if (parsed) return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value).trim();
}

function deriveExpiryDateFromValidity_(startDate, validityText) {
  const text = String(validityText || '').trim().toLowerCase();
  if (!text || !startDate) return null;
  const yearMatch = text.match(/(\d+)\s*(year|years|yr|yrs)/i);
  if (yearMatch) return new Date(startDate.getFullYear() + Number(yearMatch[1]), startDate.getMonth(), startDate.getDate());
  const monthMatch = text.match(/(\d+)\s*(month|months)/i);
  if (monthMatch) return new Date(startDate.getFullYear(), startDate.getMonth() + Number(monthMatch[1]), startDate.getDate());
  const dayMatch = text.match(/(\d+)\s*(day|days)/i);
  if (dayMatch) return new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + Number(dayMatch[1]));
  return null;
}

function extractDriveFileId_(url) {
  const text = String(url || '');
  const patterns = [/\/file\/d\/([a-zA-Z0-9_-]+)/, /id=([a-zA-Z0-9_-]+)/, /\/d\/([a-zA-Z0-9_-]+)/];
  for (var i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match && match[1]) return match[1];
  }
  return '';
}

function isGoogleDriveLink_(value) {
  const text = String(value || '').trim();
  return !!text && /drive\.google\.com|docs\.google\.com/.test(text);
}

function incrementVersion_(version) {
  const match = String(version || '').match(/V(\d+)/i);
  if (!match) return 'V2';
  return 'V' + (Number(match[1]) + 1);
}

function buildRenewalId_(sheet) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const serial = String(Math.max(sheet.getLastRow(), 1)).padStart(4, '0');
  return 'REN-' + stamp + '-' + serial;
}

function blankRowForHeaders_(headers) {
  return headers.map(() => '');
}

function calculateDaysLeft_(expiryDate) {
  const today = stripTime_(new Date());
  const diffMs = stripTime_(expiryDate).getTime() - today.getTime();
  return Math.ceil(diffMs / 86400000);
}

function parseDateValue_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return stripTime_(value);
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) return null;
  return stripTime_(parsed);
}

function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDateValue_(value, withTime) {
  const date = parseDateValue_(value);
  if (!date) return value ? String(value) : '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), withTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
}

function valueToNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const num = Number(value);
  return isNaN(num) ? null : num;
}

function normalizeText_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function safeJsonParse_(value) {
  try {
    return JSON.parse(value);
  } catch (error) {
    return null;
  }
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
