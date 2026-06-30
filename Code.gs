const SHEET_NAME = 'Listings';
const POLICIES_SHEET_NAME = 'Policies';
const BULK_POLICY_CHANGE_SHEET_NAME = 'BulkPolicyChange';
const DEFAULT_HEADERS = [
  'publish',
  'sku',
  'fulfillmentPolicyId',
  'title',
  'description',
  'categoryId',
  'priceUSD',
  'quantity',
  'condition',
  'imageUrls',
  'brand',
  'mpn',
  'aspectsJson',
  'weightLb',
  'weightOz',
  'lengthIn',
  'widthIn',
  'heightIn',
  'dutyRate',
  'overrideShippingCostUSD',
  'status',
  'offerId',
  'listingId',
  'lastError',
  'requestPreview'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('eBay出品')
    .addItem('シートを初期化', 'setupSheet')
    .addSeparator()
    .addItem('OAuth: 認証URLを表示', 'showEbayOAuthUrl')
    .addItem('OAuth: codeからRefresh Tokenを保存', 'saveRefreshTokenFromAuthCode')
    .addSeparator()
    .addItem('ポリシーID一覧を取得', 'fetchPolicyIds')
    .addSeparator()
    .addItem('一括置換: 対象出品を検索', 'prepareBulkFulfillmentPolicyChange')
    .addItem('一括置換: approve=TRUEを更新', 'applyBulkFulfillmentPolicyChange')
    .addSeparator()
    .addItem('既存出品: 選択行をドライラン', 'dryRunExistingSelectedRows')
    .addItem('既存出品: publish=TRUEをドライラン', 'dryRunExistingApprovedRows')
    .addSeparator()
    .addItem('既存出品: 選択行の送料を更新', 'updateExistingSelectedRows')
    .addItem('既存出品: publish=TRUEの送料を更新', 'updateExistingApprovedRows')
    .addToUi();
}

function showEbayOAuthUrl() {
  const props = getOAuthConfig_();
  const scope = getEbayScopes_().join(' ');
  const url =
    getWebAuthBase_(props) +
    '/oauth2/authorize?client_id=' + encodeURIComponent(props.EBAY_CLIENT_ID) +
    '&redirect_uri=' + encodeURIComponent(props.EBAY_RUNAME) +
    '&response_type=code' +
    '&scope=' + encodeURIComponent(scope);

  SpreadsheetApp.getUi().alert(
    'このURLをブラウザで開き、eBayにログインして許可してください。\n\n' +
    '許可後に移動したURL内の code= 以降をコピーします。\n\n' +
    url
  );
}

function saveRefreshTokenFromAuthCode() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'eBay OAuth code',
    '許可後URLの code= 以降を貼り付けてください。URL全体を貼っても処理できます。',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const authCode = extractAuthCode_(response.getResponseText());
  if (!authCode) {
    ui.alert('codeを読み取れませんでした。許可後URLの code= 以降を貼り付けてください。');
    return;
  }

  const props = getOAuthConfig_();
  const credentials = Utilities.base64Encode(props.EBAY_CLIENT_ID + ':' + props.EBAY_CLIENT_SECRET);
  const tokenResponse = UrlFetchApp.fetch(getAuthBase_(props) + '/identity/v1/oauth2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      Authorization: 'Basic ' + credentials
    },
    payload: {
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: props.EBAY_RUNAME
    },
    muteHttpExceptions: true
  });

  const body = parseEbayResponse_(tokenResponse, 'oauth authorization_code');
  if (!body.refresh_token) {
    throw new Error('OAuth responseにrefresh_tokenがありません: ' + JSON.stringify(body));
  }

  PropertiesService.getScriptProperties().setProperty('EBAY_REFRESH_TOKEN', body.refresh_token);
  CacheService.getScriptCache().remove('EBAY_ACCESS_TOKEN');
  ui.alert('EBAY_REFRESH_TOKENをスクリプトプロパティに保存しました。');
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, DEFAULT_HEADERS.length).setValues([DEFAULT_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange('A2:A').insertCheckboxes();
  sheet.getRange(2, 1, 1, DEFAULT_HEADERS.length).setValues([[
    false,
    'YOUR-SKU-001',
    '',
    '',
    '',
    '',
    208.99,
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    0.1,
    '',
    '',
    '',
    '',
    '',
    ''
  ]]);
  sheet.autoResizeColumns(1, DEFAULT_HEADERS.length);
  setupPolicySheet_();
}

function setupPolicySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POLICIES_SHEET_NAME) || ss.insertSheet(POLICIES_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, 4).setValues([['type', 'name', 'policyId', 'marketplaceId']]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
}

function fetchPolicyIds() {
  const props = getPolicyFetchConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(POLICIES_SHEET_NAME) || ss.insertSheet(POLICIES_SHEET_NAME);
  const rows = [['type', 'name', 'policyId', 'marketplaceId']];

  const fulfillment = getFulfillmentPolicies_(props).fulfillmentPolicies || [];
  fulfillment.forEach(policy => {
    rows.push(['FULFILLMENT', policy.name || '', policy.fulfillmentPolicyId || '', props.MARKETPLACE_ID]);
  });

  const payment = getPaymentPolicies_(props).paymentPolicies || [];
  payment.forEach(policy => {
    rows.push(['PAYMENT', policy.name || '', policy.paymentPolicyId || '', props.MARKETPLACE_ID]);
  });

  const returns = getReturnPolicies_(props).returnPolicies || [];
  returns.forEach(policy => {
    rows.push(['RETURN', policy.name || '', policy.returnPolicyId || '', props.MARKETPLACE_ID]);
  });

  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
  applyFulfillmentPolicyValidation_(fulfillment.length);
  SpreadsheetApp.getUi().alert('PoliciesシートにポリシーID一覧を取得しました。');
}

function prepareBulkFulfillmentPolicyChange() {
  const ui = SpreadsheetApp.getUi();
  const fromResponse = ui.prompt(
    '現在の配送ポリシーID',
    '置換元のfulfillmentPolicyIdを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (fromResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const toResponse = ui.prompt(
    '変更先の配送ポリシーID',
    '変更先のfulfillmentPolicyIdを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (toResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const fromPolicyId = fromResponse.getResponseText().trim();
  const toPolicyId = toResponse.getResponseText().trim();
  if (!fromPolicyId || !toPolicyId) {
    ui.alert('置換元と変更先のfulfillmentPolicyIdを両方入力してください。');
    return;
  }
  if (fromPolicyId === toPolicyId) {
    ui.alert('置換元と変更先が同じです。別のfulfillmentPolicyIdを指定してください。');
    return;
  }

  const props = getConfig_();
  const rows = [[
    'approve',
    'sku',
    'offerId',
    'listingId',
    'currentFulfillmentPolicyId',
    'targetFulfillmentPolicyId',
    'status',
    'lastError',
    'requestPreview'
  ]];

  const skus = getAllInventorySkus_(props);
  skus.forEach(sku => {
    const offers = (getOffersBySku_(sku, props).offers || []);
    offers.forEach(offer => {
      const currentPolicyId = getFulfillmentPolicyIdFromOffer_(offer);
      if (currentPolicyId !== fromPolicyId) {
        return;
      }
      const payload = buildExistingOfferUpdatePayload_(offer, {
        sku: offer.sku || sku,
        fulfillmentPolicyId: toPolicyId,
        priceUSD: 0,
        dutyRate: 0,
        overrideShippingCostUSD: getExistingShippingOverride_(offer)
      }, props);
      rows.push([
        false,
        offer.sku || sku,
        offer.offerId || '',
        getListingId_(offer),
        currentPolicyId,
        toPolicyId,
        'READY',
        '',
        JSON.stringify(payload, null, 2)
      ]);
    });
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BULK_POLICY_CHANGE_SHEET_NAME) || ss.insertSheet(BULK_POLICY_CHANGE_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  }
  sheet.autoResizeColumns(1, rows[0].length);
  ui.alert((rows.length - 1) + '件の候補をBulkPolicyChangeシートに作成しました。更新する行のapproveをTRUEにしてください。');
}

function applyBulkFulfillmentPolicyChange() {
  const props = getConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BULK_POLICY_CHANGE_SHEET_NAME);
  if (!sheet) {
    throw new Error(BULK_POLICY_CHANGE_SHEET_NAME + ' シートがありません。先に「一括置換: 対象出品を検索」を実行してください。');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert('更新対象がありません。');
    return;
  }

  const headers = values[0];
  const indexes = {};
  headers.forEach((header, index) => indexes[header] = index);
  let updated = 0;

  values.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;
    if (row[indexes.approve] !== true) {
      return;
    }

    try {
      const offerId = row[indexes.offerId];
      const targetPolicyId = row[indexes.targetFulfillmentPolicyId];
      if (!offerId || !targetPolicyId) {
        throw new Error('offerIdまたはtargetFulfillmentPolicyIdが空です。');
      }

      const offer = getOffer_(offerId, props);
      const payload = buildExistingOfferUpdatePayload_(offer, {
        sku: offer.sku || row[indexes.sku],
        fulfillmentPolicyId: targetPolicyId,
        priceUSD: 0,
        dutyRate: 0,
        overrideShippingCostUSD: getExistingShippingOverride_(offer)
      }, props);

      updateOffer_(offerId, payload, props);
      writeBulkCell_(sheet, rowNumber, indexes.status + 1, 'UPDATED');
      writeBulkCell_(sheet, rowNumber, indexes.lastError + 1, '');
      writeBulkCell_(sheet, rowNumber, indexes.requestPreview + 1, JSON.stringify(payload, null, 2));
      updated++;
    } catch (err) {
      writeBulkCell_(sheet, rowNumber, indexes.status + 1, 'ERROR');
      writeBulkCell_(sheet, rowNumber, indexes.lastError + 1, String(err && err.stack ? err.stack : err));
    }
  });

  SpreadsheetApp.getUi().alert(updated + '件を更新しました。');
}

function applyFulfillmentPolicyValidation_(fulfillmentCount) {
  if (!fulfillmentCount) {
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listingsSheet = ss.getSheetByName(SHEET_NAME);
  const policiesSheet = ss.getSheetByName(POLICIES_SHEET_NAME);
  if (!listingsSheet || !policiesSheet) {
    return;
  }

  const headers = listingsSheet.getRange(1, 1, 1, listingsSheet.getLastColumn()).getValues()[0];
  const policyColumn = headers.indexOf('fulfillmentPolicyId') + 1;
  if (policyColumn < 1) {
    return;
  }

  const policyIdRange = policiesSheet.getRange(2, 3, fulfillmentCount, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(policyIdRange, true)
    .setAllowInvalid(true)
    .build();
  listingsSheet.getRange(2, policyColumn, Math.max(1, listingsSheet.getMaxRows() - 1), 1).setDataValidation(rule);
}

function dryRunExistingSelectedRows() {
  processExistingRows_({ selectedOnly: true, updateEbay: false });
}

function dryRunExistingApprovedRows() {
  processExistingRows_({ selectedOnly: false, updateEbay: false });
}

function updateExistingSelectedRows() {
  processExistingRows_({ selectedOnly: true, updateEbay: true });
}

function updateExistingApprovedRows() {
  processExistingRows_({ selectedOnly: false, updateEbay: true });
}

function processExistingRows_(options) {
  const sheet = getSheet_();
  const rows = getRowsToProcess_(sheet, options.selectedOnly);
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('処理対象の行がありません。');
    return;
  }

  const props = getConfig_();
  rows.forEach(rowNumber => {
    try {
      const row = readRow_(sheet, rowNumber);
      const existingOffer = getExistingOfferForRow_(row, props);
      const offerId = row.offerId || existingOffer.offerId;
      if (!offerId) {
        throw new Error('既存OfferのofferIdを取得できませんでした。skuまたはofferIdを確認してください。');
      }

      const payload = buildExistingOfferUpdatePayload_(existingOffer, row, props);
      writeCell_(sheet, rowNumber, 'requestPreview', JSON.stringify(payload, null, 2));
      writeCell_(sheet, rowNumber, 'offerId', offerId);
      writeCell_(sheet, rowNumber, 'listingId', getListingId_(existingOffer));

      if (!options.updateEbay) {
        writeCell_(sheet, rowNumber, 'status', 'DRY_RUN_OK');
        writeCell_(sheet, rowNumber, 'lastError', '');
        return;
      }

      updateOffer_(offerId, payload, props);
      writeCell_(sheet, rowNumber, 'status', 'UPDATED');
      writeCell_(sheet, rowNumber, 'lastError', '');
    } catch (err) {
      writeCell_(sheet, rowNumber, 'status', 'ERROR');
      writeCell_(sheet, rowNumber, 'lastError', String(err && err.stack ? err.stack : err));
    }
  });
}

function getExistingOfferForRow_(row, props) {
  if (row.offerId) {
    return getOffer_(row.offerId, props);
  }
  if (!row.sku) {
    throw new Error('既存出品を探すには sku または offerId が必要です。');
  }

  const response = getOffersBySku_(row.sku, props);
  const offers = response.offers || [];
  if (offers.length === 0) {
    throw new Error('skuに紐づく既存Offerが見つかりません: ' + row.sku);
  }
  if (offers.length > 1) {
    throw new Error('同じskuで複数のOfferが見つかりました。offerId列に更新対象のofferIdを入れてください: ' + row.sku);
  }
  return offers[0];
}

function buildExistingOfferUpdatePayload_(offer, row, props) {
  const shippingOverride = getShippingOverride_(row);
  const payload = pickWritableOfferFields_(offer);
  payload.sku = payload.sku || row.sku;
  payload.marketplaceId = payload.marketplaceId || props.MARKETPLACE_ID;
  payload.format = payload.format || 'FIXED_PRICE';

  if (!payload.listingPolicies) {
    payload.listingPolicies = {};
  }
  payload.listingPolicies.fulfillmentPolicyId =
    row.fulfillmentPolicyId || payload.listingPolicies.fulfillmentPolicyId || props.FULFILLMENT_POLICY_ID;
  payload.listingPolicies.paymentPolicyId =
    payload.listingPolicies.paymentPolicyId || props.PAYMENT_POLICY_ID;
  payload.listingPolicies.returnPolicyId =
    payload.listingPolicies.returnPolicyId || props.RETURN_POLICY_ID;
  payload.listingPolicies.shippingCostOverrides = [{
    priority: 1,
    shippingServiceType: 'DOMESTIC',
    shippingCost: {
      value: shippingOverride,
      currency: props.CURRENCY
    }
  }];

  return removeEmptyValues_(payload);
}

function pickWritableOfferFields_(offer) {
  const payload = {};
  [
    'sku',
    'marketplaceId',
    'format',
    'availableQuantity',
    'categoryId',
    'listingDescription',
    'listingDuration',
    'listingPolicies',
    'merchantLocationKey',
    'pricingSummary',
    'quantityLimitPerBuyer',
    'lotSize',
    'tax',
    'storeCategoryNames',
    'hideBuyerDetails',
    'includeCatalogProductDetails'
  ].forEach(key => {
    if (typeof offer[key] !== 'undefined' && offer[key] !== null) {
      payload[key] = clone_(offer[key]);
    }
  });
  return payload;
}

function getShippingOverride_(row) {
  if (row.overrideShippingCostUSD !== '') {
    return asMoneyString_(row.overrideShippingCostUSD, 'overrideShippingCostUSD');
  }
  const price = asNumber_(row.priceUSD, 'priceUSD');
  const dutyRate = row.dutyRate === '' ? 0 : asNumber_(row.dutyRate, 'dutyRate');
  return roundMoney_(price * dutyRate);
}

function getOffer_(offerId, props) {
  return ebayFetch_('/sell/inventory/v1/offer/' + encodeURIComponent(offerId), {
    method: 'get'
  }, props);
}

function getOffersBySku_(sku, props) {
  const query =
    '?sku=' + encodeURIComponent(sku) +
    '&marketplace_id=' + encodeURIComponent(props.MARKETPLACE_ID) +
    '&format=FIXED_PRICE';
  return ebayFetch_('/sell/inventory/v1/offer' + query, {
    method: 'get'
  }, props);
}

function getFulfillmentPolicies_(props) {
  return ebayFetch_('/sell/account/v1/fulfillment_policy?marketplace_id=' + encodeURIComponent(props.MARKETPLACE_ID), {
    method: 'get'
  }, props);
}

function getPaymentPolicies_(props) {
  return ebayFetch_('/sell/account/v1/payment_policy?marketplace_id=' + encodeURIComponent(props.MARKETPLACE_ID), {
    method: 'get'
  }, props);
}

function getReturnPolicies_(props) {
  return ebayFetch_('/sell/account/v1/return_policy?marketplace_id=' + encodeURIComponent(props.MARKETPLACE_ID), {
    method: 'get'
  }, props);
}

function getAllInventorySkus_(props) {
  const skus = [];
  const limit = 200;
  let offset = 0;
  while (true) {
    const response = getInventoryItems_(limit, offset, props);
    const items = response.inventoryItems || [];
    items.forEach(item => {
      if (item.sku) {
        skus.push(item.sku);
      }
    });
    if (items.length < limit) {
      break;
    }
    offset += limit;
  }
  return skus;
}

function getInventoryItems_(limit, offset, props) {
  const query =
    '?limit=' + encodeURIComponent(limit) +
    '&offset=' + encodeURIComponent(offset);
  return ebayFetch_('/sell/inventory/v1/inventory_item' + query, {
    method: 'get'
  }, props);
}

function getFulfillmentPolicyIdFromOffer_(offer) {
  return offer && offer.listingPolicies ? String(offer.listingPolicies.fulfillmentPolicyId || '') : '';
}

function getExistingShippingOverride_(offer) {
  const overrides = offer && offer.listingPolicies ? offer.listingPolicies.shippingCostOverrides || [] : [];
  const domestic = overrides.filter(item => item.shippingServiceType === 'DOMESTIC')[0] || overrides[0];
  if (domestic && domestic.shippingCost && domestic.shippingCost.value !== '') {
    return domestic.shippingCost.value;
  }
  return '0.00';
}

function writeBulkCell_(sheet, rowNumber, columnNumber, value) {
  sheet.getRange(rowNumber, columnNumber).setValue(value);
}

function updateOffer_(offerId, payload, props) {
  return ebayFetch_('/sell/inventory/v1/offer/' + encodeURIComponent(offerId), {
    method: 'put',
    payload: payload
  }, props);
}

function ebayFetch_(path, request, props) {
  const options = {
    method: request.method,
    headers: {
      Authorization: 'Bearer ' + getValidAccessToken_(props),
      'Content-Language': props.CONTENT_LANGUAGE
    },
    muteHttpExceptions: true
  };
  if (request.payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(request.payload);
  }

  const url = getApiBase_(props) + path;
  const response = UrlFetchApp.fetch(url, options);
  return parseEbayResponse_(response, url);
}

function getValidAccessToken_(props) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('EBAY_ACCESS_TOKEN');
  if (cached) {
    return cached;
  }

  const credentials = Utilities.base64Encode(props.EBAY_CLIENT_ID + ':' + props.EBAY_CLIENT_SECRET);
  const response = UrlFetchApp.fetch(getAuthBase_(props) + '/identity/v1/oauth2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      Authorization: 'Basic ' + credentials
    },
    payload: {
      grant_type: 'refresh_token',
      refresh_token: props.EBAY_REFRESH_TOKEN,
      scope: getEbayScopes_().join(' ')
    },
    muteHttpExceptions: true
  });

  const body = parseEbayResponse_(response, 'oauth token');
  if (!body.access_token) {
    throw new Error('OAuth responseにaccess_tokenがありません: ' + JSON.stringify(body));
  }
  cache.put('EBAY_ACCESS_TOKEN', body.access_token, Math.max(60, Number(body.expires_in || 7200) - 300));
  return body.access_token;
}

function parseEbayResponse_(response, label) {
  const code = response.getResponseCode();
  const text = response.getContentText();
  const body = text ? JSON.parse(text) : {};
  if (code < 200 || code >= 300) {
    throw new Error(label + ' failed: HTTP ' + code + ' ' + text);
  }
  return body;
}

function getConfig_() {
  const store = PropertiesService.getScriptProperties();
  const props = {
    ENVIRONMENT: store.getProperty('ENVIRONMENT') || 'PRODUCTION',
    EBAY_CLIENT_ID: store.getProperty('EBAY_CLIENT_ID'),
    EBAY_CLIENT_SECRET: store.getProperty('EBAY_CLIENT_SECRET'),
    EBAY_REFRESH_TOKEN: store.getProperty('EBAY_REFRESH_TOKEN'),
    FULFILLMENT_POLICY_ID: store.getProperty('FULFILLMENT_POLICY_ID'),
    PAYMENT_POLICY_ID: store.getProperty('PAYMENT_POLICY_ID'),
    RETURN_POLICY_ID: store.getProperty('RETURN_POLICY_ID'),
    MARKETPLACE_ID: store.getProperty('MARKETPLACE_ID') || 'EBAY_US',
    CURRENCY: store.getProperty('CURRENCY') || 'USD',
    CONTENT_LANGUAGE: store.getProperty('CONTENT_LANGUAGE') || 'en-US'
  };

  [
    'EBAY_CLIENT_ID',
    'EBAY_CLIENT_SECRET',
    'EBAY_REFRESH_TOKEN'
  ].forEach(key => {
    if (!props[key]) {
      throw new Error('Script Propertiesに ' + key + ' を設定してください。');
    }
  });
  return props;
}

function getOAuthConfig_() {
  const store = PropertiesService.getScriptProperties();
  const props = {
    ENVIRONMENT: store.getProperty('ENVIRONMENT') || 'PRODUCTION',
    EBAY_CLIENT_ID: store.getProperty('EBAY_CLIENT_ID'),
    EBAY_CLIENT_SECRET: store.getProperty('EBAY_CLIENT_SECRET'),
    EBAY_RUNAME: store.getProperty('EBAY_RUNAME')
  };

  [
    'EBAY_CLIENT_ID',
    'EBAY_CLIENT_SECRET',
    'EBAY_RUNAME'
  ].forEach(key => {
    if (!props[key]) {
      throw new Error('Script Propertiesに ' + key + ' を設定してください。');
    }
  });
  return props;
}

function getPolicyFetchConfig_() {
  const store = PropertiesService.getScriptProperties();
  const props = {
    ENVIRONMENT: store.getProperty('ENVIRONMENT') || 'PRODUCTION',
    EBAY_CLIENT_ID: store.getProperty('EBAY_CLIENT_ID'),
    EBAY_CLIENT_SECRET: store.getProperty('EBAY_CLIENT_SECRET'),
    EBAY_REFRESH_TOKEN: store.getProperty('EBAY_REFRESH_TOKEN'),
    MARKETPLACE_ID: store.getProperty('MARKETPLACE_ID') || 'EBAY_US',
    CONTENT_LANGUAGE: store.getProperty('CONTENT_LANGUAGE') || 'en-US'
  };

  [
    'EBAY_CLIENT_ID',
    'EBAY_CLIENT_SECRET',
    'EBAY_REFRESH_TOKEN'
  ].forEach(key => {
    if (!props[key]) {
      throw new Error('Script Propertiesに ' + key + ' を設定してください。');
    }
  });
  return props;
}

function getApiBase_(props) {
  return props.ENVIRONMENT === 'SANDBOX' ? 'https://api.sandbox.ebay.com' : 'https://api.ebay.com';
}

function getAuthBase_(props) {
  return props.ENVIRONMENT === 'SANDBOX' ? 'https://api.sandbox.ebay.com' : 'https://api.ebay.com';
}

function getWebAuthBase_(props) {
  return props.ENVIRONMENT === 'SANDBOX' ? 'https://auth.sandbox.ebay.com' : 'https://auth.ebay.com';
}

function getEbayScopes_() {
  return [
    'https://api.ebay.com/oauth/api_scope/sell.inventory',
    'https://api.ebay.com/oauth/api_scope/sell.account.readonly'
  ];
}

function getRowsToProcess_(sheet, selectedOnly) {
  if (selectedOnly) {
    const range = sheet.getActiveRange();
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    const rows = [];
    for (let row = Math.max(2, start); row <= end; row++) {
      rows.push(row);
    }
    return rows;
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const publishIndex = headers.indexOf('publish');
  return values
    .map((row, index) => ({ row: row, rowNumber: index + 1 }))
    .filter(item => item.rowNumber > 1 && item.row[publishIndex] === true)
    .map(item => item.rowNumber);
}

function readRow_(sheet, rowNumber) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = {};
  headers.forEach((header, index) => {
    row[header] = values[index];
  });
  validateRequiredForExistingUpdate_(row);
  return row;
}

function validateRequiredForExistingUpdate_(row) {
  if (!row.sku && !row.offerId) {
    throw new Error('sku または offerId は必須です。');
  }
  if (row.overrideShippingCostUSD === '' && row.priceUSD === '') {
    throw new Error('overrideShippingCostUSD が空の場合は priceUSD が必須です。');
  }
}

function writeCell_(sheet, rowNumber, header, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(header);
  if (index === -1) {
    throw new Error('列が見つかりません: ' + header);
  }
  sheet.getRange(rowNumber, index + 1).setValue(value);
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(SHEET_NAME + ' シートがありません。先に「シートを初期化」を実行してください。');
  }
  return sheet;
}

function getListingId_(offer) {
  return offer && offer.listing && offer.listing.listingId ? offer.listing.listingId : '';
}

function clone_(value) {
  return JSON.parse(JSON.stringify(value));
}

function removeEmptyValues_(value) {
  if (Array.isArray(value)) {
    return value.map(removeEmptyValues_).filter(item => typeof item !== 'undefined');
  }
  if (value && typeof value === 'object') {
    const result = {};
    Object.keys(value).forEach(key => {
      const cleaned = removeEmptyValues_(value[key]);
      if (typeof cleaned !== 'undefined') {
        result[key] = cleaned;
      }
    });
    return result;
  }
  if (value === null || typeof value === 'undefined' || value === '') {
    return undefined;
  }
  return value;
}

function asNumber_(value, label) {
  const number = Number(value);
  if (!isFinite(number)) {
    throw new Error(label + ' は数値で入力してください: ' + value);
  }
  return number;
}

function asMoneyString_(value, label) {
  return roundMoney_(asNumber_(value, label));
}

function roundMoney_(value) {
  return (Math.round(Number(value) * 100) / 100).toFixed(2);
}

function extractAuthCode_(value) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }
  const match = text.match(/[?&]code=([^&]+)/);
  return decodeURIComponent(match ? match[1] : text);
}
