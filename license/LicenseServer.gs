const ALLOWLIST_SHEET_NAME = 'Allowlist';

function doGet(event) {
  return handleLicenseRequest_(event);
}

function doPost(event) {
  return handleLicenseRequest_(event);
}

function handleLicenseRequest_(event) {
  const params = event && event.parameter ? event.parameter : {};
  const email = normalizeEmail_(params.email);
  const toolId = String(params.toolId || '').trim();
  const token = String(params.token || '').trim();
  const secret = PropertiesService.getScriptProperties().getProperty('LICENSE_SERVER_SECRET');

  if (!secret || token !== secret) {
    return json_({ ok: false, reason: 'INVALID_SERVER_TOKEN' });
  }
  if (!email) {
    return json_({ ok: false, reason: 'EMAIL_REQUIRED' });
  }

  const result = findAllowedUser_(email, toolId);
  return json_(result);
}

function findAllowedUser_(email, toolId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ALLOWLIST_SHEET_NAME);
  if (!sheet) {
    return { ok: false, reason: 'ALLOWLIST_SHEET_NOT_FOUND' };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return { ok: false, reason: 'EMAIL_NOT_REGISTERED' };
  }

  const headers = values[0].map(value => String(value || '').trim());
  const indexes = {};
  headers.forEach((header, index) => indexes[header] = index);

  const emailIndex = indexes.email;
  const statusIndex = indexes.status;
  const expiresAtIndex = indexes.expiresAt;
  const toolIdIndex = indexes.toolId;
  const displayNameIndex = indexes.displayName;

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const rowEmail = normalizeEmail_(row[emailIndex]);
    if (rowEmail !== email) {
      continue;
    }

    const rowToolId = toolIdIndex >= 0 ? String(row[toolIdIndex] || '').trim() : '';
    if (rowToolId && toolId && rowToolId !== toolId) {
      continue;
    }

    const status = statusIndex >= 0 ? String(row[statusIndex] || '').trim().toUpperCase() : 'ACTIVE';
    if (status && status !== 'ACTIVE') {
      return { ok: false, reason: 'USER_NOT_ACTIVE' };
    }

    const expiresAt = expiresAtIndex >= 0 ? row[expiresAtIndex] : '';
    if (expiresAt && new Date(expiresAt).getTime() < Date.now()) {
      return { ok: false, reason: 'LICENSE_EXPIRED' };
    }

    return {
      ok: true,
      email: email,
      displayName: displayNameIndex >= 0 ? String(row[displayNameIndex] || '') : '',
      toolId: rowToolId || toolId || '',
      expiresAt: expiresAt ? new Date(expiresAt).toISOString() : ''
    };
  }

  return { ok: false, reason: 'EMAIL_NOT_REGISTERED' };
}

function setupAllowlistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ALLOWLIST_SHEET_NAME) || ss.insertSheet(ALLOWLIST_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([[
    'email',
    'status',
    'expiresAt',
    'toolId',
    'displayName',
    'notes'
  ]]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 6);
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
