const WEBAPP_USERS_SHEET_NAME = 'Users';
const WEBAPP_TOOL_ID = 'ebay-shipping-policy-tool';
const WEBAPP_LOGIN_CODE_PREFIX = 'LOGIN_CODE_';
const WEBAPP_SESSION_TTL_SECONDS = 21600;
const WEBAPP_LOGIN_CODE_TTL_SECONDS = 600;

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('eBay Shipping Policy Tool')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupWebAppSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WEBAPP_USERS_SHEET_NAME) || ss.insertSheet(WEBAPP_USERS_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, getWebAppUserHeaders_().length).setValues([getWebAppUserHeaders_()]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, getWebAppUserHeaders_().length);
}

function webRequestLoginCode(email) {
  email = normalizeWebAppEmail_(email);
  if (!email) {
    throw new Error('メールアドレスを入力してください。');
  }

  const user = getWebAppUserByEmail_(email);
  if (!user || user.status !== 'ACTIVE') {
    throw new Error('このメールアドレスは利用許可されていません。');
  }
  if (user.expiresAt && new Date(user.expiresAt).getTime() < Date.now()) {
    throw new Error('利用期限が切れています。');
  }

  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put(WEBAPP_LOGIN_CODE_PREFIX + email, code, WEBAPP_LOGIN_CODE_TTL_SECONDS);
  MailApp.sendEmail({
    to: email,
    subject: 'eBay Shipping Policy Tool 確認コード',
    body: '確認コード: ' + code + '\n\n10分以内にWeb Appへ入力してください。'
  });
  return { ok: true };
}

function webVerifyLoginCode(email, code) {
  email = normalizeWebAppEmail_(email);
  code = String(code || '').trim();
  const cached = CacheService.getScriptCache().get(WEBAPP_LOGIN_CODE_PREFIX + email);
  if (!cached || cached !== code) {
    throw new Error('確認コードが違うか、有効期限が切れています。');
  }
  const token = createWebAppSessionToken_(email);
  return { ok: true, token: token, email: email };
}

function webGetDashboard(token) {
  const email = requireWebAppSession_(token);
  const user = getWebAppUserByEmail_(email);
  return {
    email: email,
    displayName: user.displayName || '',
    status: user.status || '',
    expiresAt: user.expiresAt || '',
    ebay: {
      clientIdSet: Boolean(user.ebayClientId),
      clientSecretSet: Boolean(user.ebayClientSecret),
      runameSet: Boolean(user.ebayRuname),
      refreshTokenSet: Boolean(user.ebayRefreshToken),
      marketplaceId: user.marketplaceId || 'EBAY_MOTORS',
      tradingSiteId: user.tradingSiteId || '100'
    }
  };
}

function webSaveEbayDeveloperConfig(token, config) {
  const email = requireWebAppSession_(token);
  const user = getWebAppUserByEmail_(email);
  config = config || {};
  updateWebAppUser_(email, {
    ebayClientId: String(config.ebayClientId || user.ebayClientId || '').trim(),
    ebayClientSecret: String(config.ebayClientSecret || user.ebayClientSecret || '').trim(),
    ebayRuname: String(config.ebayRuname || user.ebayRuname || '').trim(),
    marketplaceId: String(config.marketplaceId || 'EBAY_MOTORS').trim(),
    currency: String(config.currency || 'USD').trim(),
    contentLanguage: String(config.contentLanguage || 'en-US').trim(),
    tradingSiteId: String(config.tradingSiteId || '100').trim(),
    tradingApiVersion: String(config.tradingApiVersion || '1455').trim(),
    updatedAt: new Date()
  });
  return { ok: true };
}

function webGetEbayOAuthUrl(token) {
  const email = requireWebAppSession_(token);
  const user = getWebAppUserByEmail_(email);
  if (!user.ebayClientId || !user.ebayRuname) {
    throw new Error('Client ID と RuName を先に保存してください。');
  }
  const authBase = getWebAppWebAuthBase_(user);
  const url =
    authBase +
    '/oauth2/authorize?client_id=' + encodeURIComponent(user.ebayClientId) +
    '&redirect_uri=' + encodeURIComponent(user.ebayRuname) +
    '&response_type=code' +
    '&scope=' + encodeURIComponent(getWebAppEbayScopes_().join(' '));
  return { ok: true, url: url };
}

function webSaveEbayRefreshToken(token, redirectedUrlOrCode) {
  const email = requireWebAppSession_(token);
  const user = getWebAppUserByEmail_(email);
  if (!user.ebayClientId || !user.ebayClientSecret || !user.ebayRuname) {
    throw new Error('Client ID / Client Secret / RuName を先に保存してください。');
  }

  const authCode = extractWebAppAuthCode_(redirectedUrlOrCode);
  if (!authCode) {
    throw new Error('code= を含むURL、またはcode本体を入力してください。');
  }

  const credentials = Utilities.base64Encode(user.ebayClientId + ':' + user.ebayClientSecret);
  const response = UrlFetchApp.fetch(getWebAppAuthBase_(user) + '/identity/v1/oauth2/token', {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: {
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: user.ebayRuname
    },
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  const body = text ? JSON.parse(text) : {};
  if (code < 200 || code >= 300) {
    throw new Error('Refresh Token取得に失敗しました: HTTP ' + code + ' ' + text);
  }
  if (!body.refresh_token) {
    throw new Error('refresh_token が返りませんでした: ' + text);
  }

  updateWebAppUser_(email, {
    ebayRefreshToken: body.refresh_token,
    updatedAt: new Date()
  });
  return { ok: true };
}

function getWebAppUserHeaders_() {
  return [
    'email',
    'status',
    'expiresAt',
    'displayName',
    'ebayClientId',
    'ebayClientSecret',
    'ebayRuname',
    'ebayRefreshToken',
    'marketplaceId',
    'currency',
    'contentLanguage',
    'tradingSiteId',
    'tradingApiVersion',
    'updatedAt',
    'notes'
  ];
}

function getWebAppUsersSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEBAPP_USERS_SHEET_NAME);
  if (!sheet) {
    throw new Error('Usersシートがありません。setupWebAppSheets を実行してください。');
  }
  return sheet;
}

function getWebAppUserByEmail_(email) {
  email = normalizeWebAppEmail_(email);
  const sheet = getWebAppUsersSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return null;
  }
  const headers = values[0].map(value => String(value || '').trim());
  const emailIndex = headers.indexOf('email');
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    if (normalizeWebAppEmail_(row[emailIndex]) === email) {
      return rowToWebAppUser_(headers, row, rowIndex + 1);
    }
  }
  return null;
}

function updateWebAppUser_(email, fields) {
  email = normalizeWebAppEmail_(email);
  const sheet = getWebAppUsersSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value || '').trim());
  const emailIndex = headers.indexOf('email');
  let rowNumber = 0;
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    if (normalizeWebAppEmail_(values[rowIndex][emailIndex]) === email) {
      rowNumber = rowIndex + 1;
      break;
    }
  }
  if (!rowNumber) {
    throw new Error('Usersシートにメールアドレスが登録されていません: ' + email);
  }

  Object.keys(fields).forEach(key => {
    const columnIndex = headers.indexOf(key);
    if (columnIndex >= 0) {
      sheet.getRange(rowNumber, columnIndex + 1).setValue(fields[key]);
    }
  });
}

function rowToWebAppUser_(headers, row, rowNumber) {
  const user = { rowNumber: rowNumber };
  headers.forEach((header, index) => {
    user[header] = row[index];
  });
  user.email = normalizeWebAppEmail_(user.email);
  user.status = String(user.status || '').trim().toUpperCase();
  return user;
}

function createWebAppSessionToken_(email) {
  const expiresAt = Date.now() + WEBAPP_SESSION_TTL_SECONDS * 1000;
  const payload = email + '|' + expiresAt;
  return payload + '|' + signWebAppPayload_(payload);
}

function requireWebAppSession_(token) {
  const parts = String(token || '').split('|');
  if (parts.length !== 3) {
    throw new Error('ログインしてください。');
  }
  const payload = parts[0] + '|' + parts[1];
  const expected = signWebAppPayload_(payload);
  if (parts[2] !== expected) {
    throw new Error('ログイン情報が無効です。');
  }
  if (Number(parts[1]) < Date.now()) {
    throw new Error('ログイン期限が切れました。再ログインしてください。');
  }
  return normalizeWebAppEmail_(parts[0]);
}

function signWebAppPayload_(payload) {
  const secret = PropertiesService.getScriptProperties().getProperty('WEBAPP_SESSION_SECRET');
  if (!secret) {
    throw new Error('Script Propertiesに WEBAPP_SESSION_SECRET を設定してください。');
  }
  const signature = Utilities.computeHmacSha256Signature(payload, secret);
  return Utilities.base64EncodeWebSafe(signature);
}

function getWebAppEbayScopes_() {
  return [
    'https://api.ebay.com/oauth/api_scope',
    'https://api.ebay.com/oauth/api_scope/sell.inventory',
    'https://api.ebay.com/oauth/api_scope/sell.account'
  ];
}

function getWebAppAuthBase_(user) {
  return String(user.environment || '').toUpperCase() === 'SANDBOX'
    ? 'https://api.sandbox.ebay.com'
    : 'https://api.ebay.com';
}

function getWebAppWebAuthBase_(user) {
  return String(user.environment || '').toUpperCase() === 'SANDBOX'
    ? 'https://auth.sandbox.ebay.com'
    : 'https://auth.ebay.com';
}

function extractWebAppAuthCode_(value) {
  const text = String(value || '').trim();
  const match = text.match(/[?&]code=([^&]+)/);
  return decodeURIComponent(match ? match[1] : text);
}

function normalizeWebAppEmail_(value) {
  return String(value || '').trim().toLowerCase();
}
