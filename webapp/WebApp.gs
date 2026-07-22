const WEBAPP_USERS_SHEET_NAME = 'Users';
const WEBAPP_TOOL_ID = 'ebay-shipping-policy-tool';
const WEBAPP_LOGIN_CODE_PREFIX = 'LOGIN_CODE_';
const WEBAPP_SESSION_TTL_SECONDS = 604800; // ログイン有効期間: 7日間
const WEBAPP_LOGIN_CODE_TTL_SECONDS = 600;

// 診断用: 実行すると「今日あと何通メールを送れるか」をログに表示する
// 残り約100 = 無料アカウント / 残り約1500 = Google Workspace(有料枠)
function checkMailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();
  Logger.log('本日の残りメール送信可能数: ' + remaining + ' 通');
  Logger.log(remaining > 200 ? '→ Google Workspace(有料枠)の上限で動いています' : '→ 無料アカウントの上限(100通/日)で動いています');
}

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
  const ebayApp = getWebAppEbayAppConfig_(false);
  return {
    email: email,
    displayName: user.displayName || '',
    status: user.status || '',
    expiresAt: user.expiresAt || '',
    ebay: {
      appConfigured: Boolean(ebayApp.clientId && ebayApp.clientSecret && ebayApp.runame),
      refreshTokenSet: Boolean(user.ebayRefreshToken),
      marketplaceId: user.marketplaceId || 'EBAY_MOTORS',
      tradingSiteId: user.tradingSiteId || '100'
    }
  };
}

function webSaveEbayDeveloperConfig(token, config) {
  const email = requireWebAppSession_(token);
  config = config || {};
  updateWebAppUser_(email, {
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
  requireWebAppSession_(token);
  const ebayApp = getWebAppEbayAppConfig_();
  const authBase = getWebAppWebAuthBase_(ebayApp);
  const url =
    authBase +
    '/oauth2/authorize?client_id=' + encodeURIComponent(ebayApp.clientId) +
    '&redirect_uri=' + encodeURIComponent(ebayApp.runame) +
    '&response_type=code' +
    '&scope=' + encodeURIComponent(getWebAppEbayScopes_().join(' '));
  return { ok: true, url: url };
}

function webSaveEbayRefreshToken(token, redirectedUrlOrCode) {
  const email = requireWebAppSession_(token);
  const ebayApp = getWebAppEbayAppConfig_();

  const authCode = extractWebAppAuthCode_(redirectedUrlOrCode);
  if (!authCode) {
    throw new Error('code= を含むURL、またはcode本体を入力してください。');
  }

  const credentials = Utilities.base64Encode(ebayApp.clientId + ':' + ebayApp.clientSecret);
  const response = UrlFetchApp.fetch(getWebAppAuthBase_(ebayApp) + '/identity/v1/oauth2/token', {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: {
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: ebayApp.runame
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

function getWebAppEbayAppConfig_(required) {
  required = required !== false;
  const store = PropertiesService.getScriptProperties();
  const config = {
    environment: store.getProperty('ENVIRONMENT') || 'PRODUCTION',
    clientId: store.getProperty('EBAY_CLIENT_ID'),
    clientSecret: store.getProperty('EBAY_CLIENT_SECRET'),
    runame: store.getProperty('EBAY_RUNAME')
  };
  [
    'clientId',
    'clientSecret',
    'runame'
  ].forEach(key => {
    if (required && !config[key]) {
      throw new Error('Script Propertiesに ' + getWebAppEbayPropertyName_(key) + ' を設定してください。');
    }
  });
  return config;
}

function getWebAppEbayPropertyName_(key) {
  const names = {
    clientId: 'EBAY_CLIENT_ID',
    clientSecret: 'EBAY_CLIENT_SECRET',
    runame: 'EBAY_RUNAME'
  };
  return names[key] || key;
}

function getWebAppAuthBase_(config) {
  return String(config.environment || '').toUpperCase() === 'SANDBOX'
    ? 'https://api.sandbox.ebay.com'
    : 'https://api.ebay.com';
}

function getWebAppWebAuthBase_(config) {
  return String(config.environment || '').toUpperCase() === 'SANDBOX'
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

// ==== ここから本体機能のエンドポイント(EbayApi.gs を使用) ====

const WEB_BULK_CHUNK_TIME_LIMIT_MS = 40000;
const WEB_BULK_APPLY_MAX_PER_CALL = 10;

function webFetchPolicies(token) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  return waFetchAllPolicies_(props);
}

function webTradingPreview(token, form) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  const row = normalizeTradingForm_(form);
  const item = waGetTradingItem_(row.listingId, props);
  if (row.priceUSD === '' && item.price) {
    row.priceUSD = waRoundMoney_(item.price);
  }
  const requestXml = waBuildReviseFixedPriceItemXml_(row, item, props);
  return {
    ok: true,
    item: item,
    priceUSD: row.priceUSD,
    computedShipping: waGetShippingOverride_(row),
    computedAdditionalShipping: waGetAdditionalShippingCost_(waGetShippingOverride_(row), row),
    targetPolicyId: row.fulfillmentPolicyId || item.shippingProfileId,
    requestXml: requestXml
  };
}

function webTradingUpdate(token, form) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  const row = normalizeTradingForm_(form);
  const item = waGetTradingItem_(row.listingId, props);
  if (row.priceUSD === '' && item.price) {
    row.priceUSD = waRoundMoney_(item.price);
  }
  const requestXml = waBuildReviseFixedPriceItemXml_(row, item, props);
  waReviseFixedPriceItem_(requestXml, props);
  return {
    ok: true,
    itemId: item.itemId,
    computedShipping: waGetShippingOverride_(row),
    computedAdditionalShipping: waGetAdditionalShippingCost_(waGetShippingOverride_(row), row)
  };
}

function normalizeTradingForm_(form) {
  form = form || {};
  const listingId = String(form.listingId || '').trim();
  if (!listingId) {
    throw new Error('Item ID(出品番号)を入力してください。');
  }
  return {
    listingId: listingId,
    fulfillmentPolicyId: String(form.fulfillmentPolicyId || '').trim(),
    priceUSD: form.priceUSD === '' || typeof form.priceUSD === 'undefined' || form.priceUSD === null ? '' : form.priceUSD,
    dutyRate: form.dutyRate === '' || typeof form.dutyRate === 'undefined' || form.dutyRate === null ? '' : form.dutyRate,
    addFixedUSD: form.addFixedUSD === '' || typeof form.addFixedUSD === 'undefined' || form.addFixedUSD === null ? '' : form.addFixedUSD,
    additionalRatePercent: form.additionalRatePercent === '' || typeof form.additionalRatePercent === 'undefined' || form.additionalRatePercent === null ? '' : form.additionalRatePercent,
    overrideShippingCostUSD: form.overrideShippingCostUSD === '' || typeof form.overrideShippingCostUSD === 'undefined' || form.overrideShippingCostUSD === null ? '' : form.overrideShippingCostUSD
  };
}

function webInventoryPreview(token, form) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  const row = normalizeInventoryForm_(form);
  const offer = waGetExistingOfferForRow_(row, props);
  if (row.priceUSD === '') {
    row.priceUSD = waGetExistingOfferPrice_(offer);
  }
  const payload = waBuildExistingOfferUpdatePayload_(offer, row, props);
  return {
    ok: true,
    offerId: String(offer.offerId || ''),
    listingId: waGetListingIdFromOffer_(offer),
    sku: String(offer.sku || row.sku || ''),
    priceUSD: row.priceUSD,
    currentPolicyId: waGetFulfillmentPolicyIdFromOffer_(offer),
    targetPolicyId: row.fulfillmentPolicyId || waGetFulfillmentPolicyIdFromOffer_(offer),
    computedShipping: waGetShippingOverride_(row),
    computedAdditionalShipping: waGetAdditionalShippingCost_(waGetShippingOverride_(row), row),
    payloadPreview: JSON.stringify(payload, null, 2)
  };
}

function webInventoryUpdate(token, form) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  const row = normalizeInventoryForm_(form);
  const offer = waGetExistingOfferForRow_(row, props);
  const offerId = String(offer.offerId || '');
  if (!offerId) {
    throw new Error('offer IDを取得できませんでした。');
  }
  if (row.priceUSD === '') {
    row.priceUSD = waGetExistingOfferPrice_(offer);
  }
  const payload = waBuildExistingOfferUpdatePayload_(offer, row, props);
  waUpdateOffer_(offerId, payload, props);
  return {
    ok: true,
    offerId: offerId,
    computedShipping: waGetShippingOverride_(row)
  };
}

function normalizeInventoryForm_(form) {
  form = form || {};
  return {
    sku: String(form.sku || '').trim(),
    offerId: String(form.offerId || '').trim(),
    fulfillmentPolicyId: String(form.fulfillmentPolicyId || '').trim(),
    priceUSD: form.priceUSD === '' || typeof form.priceUSD === 'undefined' || form.priceUSD === null ? '' : form.priceUSD,
    dutyRate: form.dutyRate === '' || typeof form.dutyRate === 'undefined' || form.dutyRate === null ? '' : form.dutyRate,
    additionalRatePercent: form.additionalRatePercent === '' || typeof form.additionalRatePercent === 'undefined' || form.additionalRatePercent === null ? '' : form.additionalRatePercent,
    overrideShippingCostUSD: form.overrideShippingCostUSD === '' || typeof form.overrideShippingCostUSD === 'undefined' || form.overrideShippingCostUSD === null ? '' : form.overrideShippingCostUSD
  };
}

function webMigrateListing(token, listingId) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  listingId = String(listingId || '').trim();
  if (!listingId) {
    throw new Error('Item ID(出品番号)を入力してください。');
  }
  const response = waBulkMigrateListing_(listingId, props);
  const result = waGetMigrationResult_(response, listingId);
  if (result.errors && result.errors.length) {
    throw new Error('移行に失敗しました: ' + JSON.stringify(result.errors));
  }
  const items = result.inventoryItems || [];
  if (items.length === 0 || !items[0].offerId) {
    throw new Error('移行結果にoffer IDがありません: ' + JSON.stringify(result));
  }
  return {
    ok: true,
    offerId: String(items[0].offerId),
    sku: String(items[0].sku || '')
  };
}

// 一括変更: 出品を少しずつ検索する(ブラウザから繰り返し呼ばれる)
function webBulkSearchChunk(token, params) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  params = params || {};
  const fromPolicyId = String(params.fromPolicyId || '').trim();
  if (!fromPolicyId) {
    throw new Error('置換元の配送ポリシーIDを入力してください。');
  }

  const startedAt = Date.now();
  let pageNumber = Math.max(1, Number(params.pageNumber) || 1);
  let itemIndex = Math.max(0, Number(params.itemIndex) || 0);
  let totalPages = Number(params.totalPages) || 0;
  let checked = 0;
  const matches = [];
  let done = false;

  while (Date.now() - startedAt < WEB_BULK_CHUNK_TIME_LIMIT_MS) {
    // GetSellerList(終了日ウィンドウ方式)を使う: 25,000件上限がなく、出品国も取れる
    const page = waGetSellerListPage_(pageNumber, 200, props);
    totalPages = page.totalPages;
    if (page.items.length === 0) {
      done = true;
      break;
    }

    let finishedPage = true;
    for (let index = itemIndex; index < page.items.length; index++) {
      if (Date.now() - startedAt >= WEB_BULK_CHUNK_TIME_LIMIT_MS) {
        itemIndex = index;
        finishedPage = false;
        break;
      }
      const summary = page.items[index];
      checked++;
      // アメリカ(ebay.com)以外の出品はスキップ(各国出品が混ざって上限を圧迫していたため)
      if (summary.site && summary.site !== 'US') {
        continue;
      }
      const item = summary.shippingProfileId ? summary : waGetTradingItem_(summary.itemId, props);
      if (String(item.shippingProfileId) !== fromPolicyId) {
        continue;
      }
      matches.push({
        itemId: String(item.itemId),
        sku: String(item.sku || ''),
        title: String(item.title || ''),
        priceUSD: item.price ? waRoundMoney_(item.price) : ''
      });
    }

    if (!finishedPage) {
      break;
    }
    pageNumber++;
    itemIndex = 0;
    if (pageNumber > totalPages) {
      done = true;
      break;
    }
  }

  return {
    matches: matches,
    checked: checked,
    pageNumber: pageNumber,
    itemIndex: itemIndex,
    totalPages: totalPages,
    done: done
  };
}

// 一括変更: Item ID直接指定で出品情報を取得する(検索の25,000件上限で見つからない出品向け)
// eBayのポリシー画面「N listings」リンクから拾ったItem IDを貼り付けて使う
function webBulkLookupItems(token, request) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  request = request || {};
  const ids = (request.itemIds || []).slice(0, 10);
  const rows = ids.map(id => {
    const itemId = String(id || '').trim();
    try {
      const item = waGetTradingItem_(itemId, props);
      return {
        ok: true,
        itemId: String(item.itemId || itemId),
        sku: String(item.sku || ''),
        title: String(item.title || ''),
        priceUSD: item.price ? waRoundMoney_(item.price) : '',
        shippingProfileId: String(item.shippingProfileId || '')
      };
    } catch (err) {
      return { ok: false, itemId: itemId, error: String(err && err.message ? err.message : err) };
    }
  });
  return { rows: rows };
}

// 一括変更: 承認された出品をまとめて更新する(1回の呼び出しで最大10件)
function webBulkApply(token, request) {
  const email = requireWebAppSession_(token);
  const props = waGetUserEbayProps_(email);
  request = request || {};
  const toPolicyId = String(request.toPolicyId || '').trim();
  if (!toPolicyId) {
    throw new Error('変更先の配送ポリシーIDを入力してください。');
  }
  const items = (request.items || []).slice(0, WEB_BULK_APPLY_MAX_PER_CALL);
  const dutyRate = request.dutyRate === '' || typeof request.dutyRate === 'undefined' || request.dutyRate === null ? '' : request.dutyRate;
  const addFixedUSD = request.addFixedUSD === '' || typeof request.addFixedUSD === 'undefined' || request.addFixedUSD === null ? '' : request.addFixedUSD;
  const additionalRatePercent = request.additionalRatePercent === '' || typeof request.additionalRatePercent === 'undefined' || request.additionalRatePercent === null ? '' : request.additionalRatePercent;

  const results = items.map(entry => {
    const itemId = String(entry.itemId || '').trim();
    try {
      const row = {
        listingId: itemId,
        fulfillmentPolicyId: toPolicyId,
        priceUSD: entry.priceUSD === '' || typeof entry.priceUSD === 'undefined' || entry.priceUSD === null ? '' : entry.priceUSD,
        dutyRate: dutyRate,
        addFixedUSD: addFixedUSD,
        additionalRatePercent: additionalRatePercent,
        overrideShippingCostUSD: entry.overrideShippingCostUSD === '' || typeof entry.overrideShippingCostUSD === 'undefined' || entry.overrideShippingCostUSD === null ? '' : entry.overrideShippingCostUSD
      };
      const item = waGetTradingItem_(row.listingId, props);
      if (row.priceUSD === '' && item.price) {
        row.priceUSD = waRoundMoney_(item.price);
      }
      const requestXml = waBuildReviseFixedPriceItemXml_(row, item, props);
      waReviseFixedPriceItem_(requestXml, props);
      return {
        itemId: itemId,
        ok: true,
        computedShipping: waGetShippingOverride_(row),
        computedAdditionalShipping: waGetAdditionalShippingCost_(waGetShippingOverride_(row), row)
      };
    } catch (err) {
      return {
        itemId: itemId,
        ok: false,
        error: String(err && err.message ? err.message : err)
      };
    }
  });

  return { results: results };
}
