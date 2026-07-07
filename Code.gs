const SHEET_NAME = 'Listings';
const POLICIES_SHEET_NAME = 'Policies';
const BULK_POLICY_CHANGE_SHEET_NAME = 'BulkPolicyChange';
const BULK_TRADING_STATE_KEY = 'BULK_TRADING_STATE';
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
    .addItem('Trading API: 選択行をドライラン', 'dryRunTradingSelectedRows')
    .addItem('Trading API: 選択行を更新', 'updateTradingSelectedRows')
    .addSeparator()
    .addItem('Trading一括: 新規検索を開始', 'prepareBulkTradingPolicyChange')
    .addItem('Trading一括: 続きから検索', 'continueBulkTradingPolicySearch')
    .addItem('Trading一括: approve=TRUEを更新', 'applyBulkTradingPolicyChange')
    .addSeparator()
    .addItem('移行: 選択行をInventory APIへ移行', 'migrateSelectedRowsToInventory')
    .addItem('移行: publish=TRUEをInventory APIへ移行', 'migrateApprovedRowsToInventory')
    .addSeparator()
    .addItem('既存出品: 選択行の送料を更新', 'updateExistingSelectedRows')
    .addItem('既存出品: publish=TRUEの送料を更新', 'updateExistingApprovedRows')
    .addToUi();
}

function showEbayOAuthUrl() {
  try {
    const props = getOAuthConfig_();
    const scope = getEbayScopes_().join(' ');
    const url =
      getWebAuthBase_(props) +
      '/oauth2/authorize?client_id=' + encodeURIComponent(props.EBAY_CLIENT_ID) +
      '&redirect_uri=' + encodeURIComponent(props.EBAY_RUNAME) +
      '&response_type=code' +
      '&scope=' + encodeURIComponent(scope);

    writeOAuthUrlToSheet_(url, props);
    SpreadsheetApp.getUi().alert(
      'OAuthシートにeBay認証URLを書き出しました。\n\n' +
      'OAuthシートのB2セルを開いて、https://auth.ebay.com から始まるURLをブラウザで開いてください。\n\n' +
      '許可後に移動したURL内の code= 付きURL全体をコピーします。'
    );
  } catch (err) {
    writeOAuthErrorToSheet_(err);
    SpreadsheetApp.getUi().alert(
      'OAuthシートに設定エラーを書き出しました。\n\n' +
      'OAuthシートの内容を確認し、スクリプトプロパティを修正してください。'
    );
  }
}

function writeOAuthUrlToSheet_(url, props) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('OAuth') || ss.insertSheet('OAuth');
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['item', 'value']]);
  sheet.getRange(2, 1, 5, 2).setValues([
    ['authorizationUrl', url],
    ['urlMustStartWith', getWebAuthBase_(props) + '/oauth2/authorize'],
    ['environment', props.ENVIRONMENT],
    ['clientId', props.EBAY_CLIENT_ID],
    ['runame', props.EBAY_RUNAME]
  ]);
  sheet.getRange('B2').setFormula('=HYPERLINK("' + url.replace(/"/g, '""') + '","Open eBay authorization URL")');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
}

function writeOAuthErrorToSheet_(err) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('OAuth') || ss.insertSheet('OAuth');
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['item', 'value']]);
  sheet.getRange(2, 1, 8, 2).setValues([
    ['status', 'ERROR'],
    ['message', String(err && err.message ? err.message : err)],
    ['requiredProperty', 'EBAY_CLIENT_ID'],
    ['requiredProperty', 'EBAY_CLIENT_SECRET'],
    ['requiredProperty', 'EBAY_RUNAME'],
    ['optionalProperty', 'ENVIRONMENT=PRODUCTION'],
    ['note', 'Apps Script > プロジェクトの設定 > スクリプト プロパティを確認してください。'],
    ['runameExample', 'Makoto_Araki-MakotoAr-rakura-acrjgqi']
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
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

function prepareBulkTradingPolicyChange() {
  const ui = SpreadsheetApp.getUi();
  const fromResponse = ui.prompt(
    '現在の配送ポリシーID',
    '置換元のShippingProfileIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (fromResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const toResponse = ui.prompt(
    '変更先の配送ポリシーID',
    '変更先のShippingProfileIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (toResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const dutyResponse = ui.prompt(
    '関税率',
    '送料上書き額を計算する関税率を入力してください。例: 10%なら 0.1',
    ui.ButtonSet.OK_CANCEL
  );
  if (dutyResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const fromPolicyId = fromResponse.getResponseText().trim();
  const toPolicyId = toResponse.getResponseText().trim();
  const dutyRate = dutyResponse.getResponseText().trim();
  const limitResponse = ui.prompt(
    '最大検索件数',
    '今回チェックする最大出品数を入力してください。まずは 50 推奨です。',
    ui.ButtonSet.OK_CANCEL
  );
  if (limitResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const maxItems = asInteger_(limitResponse.getResponseText().trim() || 50, '最大検索件数');
  if (maxItems < 1) {
    ui.alert('最大検索件数は1以上で入力してください。');
    return;
  }
  if (!fromPolicyId || !toPolicyId || dutyRate === '') {
    ui.alert('置換元、変更先、関税率をすべて入力してください。');
    return;
  }
  if (fromPolicyId === toPolicyId) {
    ui.alert('置換元と変更先が同じです。別のShippingProfileIDを指定してください。');
    return;
  }
  asNumber_(dutyRate, 'dutyRate');

  initializeBulkTradingSheet_();
  saveBulkTradingState_({
    fromPolicyId: fromPolicyId,
    toPolicyId: toPolicyId,
    dutyRate: dutyRate,
    batchSize: maxItems,
    nextPageNumber: 1,
    nextItemIndex: 0,
    totalPages: '',
    checked: 0,
    matched: 0,
    done: false
  });
  continueBulkTradingPolicySearch();
}

function continueBulkTradingPolicySearch() {
  const state = getBulkTradingState_();
  if (!state) {
    SpreadsheetApp.getUi().alert('一括検索の状態がありません。先に「Trading一括: 新規検索を開始」を実行してください。');
    return;
  }
  if (state.done) {
    SpreadsheetApp.getUi().alert('検索は完了済みです。BulkPolicyChangeシートを確認してください。');
    return;
  }

  const props = getConfig_();
  const sheet = getBulkPolicyChangeSheet_();
  const entriesPerPage = 200;
  let checkedThisRun = 0;
  let matchedThisRun = 0;

  setBulkSearchStatus_('RUNNING', '検索中: page ' + state.nextPageNumber + ' / ' + (state.totalPages || '?') + ', 今回最大 ' + state.batchSize + ' 件');

  while (checkedThisRun < state.batchSize) {
    const page = getTradingActiveListPage_(state.nextPageNumber, entriesPerPage, props);
    state.totalPages = page.totalPages;
    if (page.items.length === 0) {
      state.done = true;
      break;
    }

    let finishedPage = true;
    for (let index = state.nextItemIndex || 0; index < page.items.length; index++) {
      if (checkedThisRun >= state.batchSize) {
        state.nextItemIndex = index;
        finishedPage = false;
        break;
      }
      const itemSummary = page.items[index];
      checkedThisRun++;
      state.checked++;
      const item = itemSummary.shippingProfileId ? itemSummary : getTradingItem_(itemSummary.itemId, props);
      if (String(item.shippingProfileId) !== String(state.fromPolicyId)) {
        continue;
      }

      const row = {
        listingId: item.itemId,
        sku: item.sku,
        priceUSD: item.price,
        dutyRate: state.dutyRate,
        fulfillmentPolicyId: state.toPolicyId,
        overrideShippingCostUSD: ''
      };
      const requestXml = buildReviseFixedPriceItemXml_(row, item, props);
      appendBulkTradingCandidate_(sheet, [
        false,
        item.itemId,
        item.sku,
        item.price,
        state.dutyRate,
        state.fromPolicyId,
        state.toPolicyId,
        '',
        'READY',
        '',
        requestXml
      ]);
      matchedThisRun++;
      state.matched++;
    }

    if (!finishedPage) {
      break;
    }
    state.nextPageNumber++;
    state.nextItemIndex = 0;
    if (state.nextPageNumber > state.totalPages) {
      state.done = true;
      break;
    }
  }

  saveBulkTradingState_(state);
  setBulkSearchStatus_(
    state.done ? 'DONE' : 'PAUSED',
    '今回チェック: ' + checkedThisRun + '件 / 今回候補: ' + matchedThisRun + '件 / 累計チェック: ' + state.checked + '件 / 累計候補: ' + state.matched + '件 / 次ページ: ' + state.nextPageNumber + ' / ' + (state.totalPages || '?')
  );
  SpreadsheetApp.getUi().alert(
    (state.done ? '検索完了。' : 'この回の検索が完了しました。続きは「Trading一括: 続きから検索」を実行してください。') +
    '\n今回チェック: ' + checkedThisRun +
    '\n今回候補: ' + matchedThisRun +
    '\n累計候補: ' + state.matched
  );
}

function initializeBulkTradingSheet_() {
  const sheet = getBulkPolicyChangeSheet_();
  const rows = [[
    'approve',
    'listingId',
    'sku',
    'priceUSD',
    'dutyRate',
    'currentFulfillmentPolicyId',
    'targetFulfillmentPolicyId',
    'overrideShippingCostUSD',
    'status',
    'lastError',
    'requestPreview'
  ]];
  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
  setBulkSearchStatus_('READY', '検索条件を保存しました。候補は2行目以降に追加されます。');
}

function appendBulkTradingCandidate_(sheet, row) {
  const lastRow = Math.max(2, sheet.getLastRow());
  const listingIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let rowNumber = 2;
  for (let index = listingIds.length - 1; index >= 0; index--) {
    if (listingIds[index][0] !== '') {
      rowNumber = index + 3;
      break;
    }
  }
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  sheet.getRange(rowNumber, 1).insertCheckboxes();
  SpreadsheetApp.flush();
}

function prepareBulkSheetWithStatus_(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BULK_POLICY_CHANGE_SHEET_NAME) || ss.insertSheet(BULK_POLICY_CHANGE_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 2, 2).setValues([
    ['status', 'message'],
    ['RUNNING', message]
  ]);
  sheet.autoResizeColumns(1, 2);
  SpreadsheetApp.flush();
}

function getBulkPolicyChangeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(BULK_POLICY_CHANGE_SHEET_NAME) || ss.insertSheet(BULK_POLICY_CHANGE_SHEET_NAME);
}

function saveBulkTradingState_(state) {
  PropertiesService.getScriptProperties().setProperty(BULK_TRADING_STATE_KEY, JSON.stringify(state));
}

function getBulkTradingState_() {
  const value = PropertiesService.getScriptProperties().getProperty(BULK_TRADING_STATE_KEY);
  return value ? JSON.parse(value) : null;
}

function setBulkSearchStatus_(status, message) {
  const sheet = getBulkPolicyChangeSheet_();
  const statusColumn = 13;
  sheet.getRange(1, statusColumn, 4, 2).setValues([
    ['bulkStatus', status],
    ['bulkMessage', message],
    ['updatedAt', new Date()],
    ['note', '候補行はA-K列です。approve=TRUEの行だけ更新されます。']
  ]);
  sheet.autoResizeColumns(statusColumn, 2);
  SpreadsheetApp.flush();
}

function applyBulkTradingPolicyChange() {
  const props = getConfig_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BULK_POLICY_CHANGE_SHEET_NAME);
  if (!sheet) {
    throw new Error(BULK_POLICY_CHANGE_SHEET_NAME + ' シートがありません。先に「Trading一括: 対象出品を検索」を実行してください。');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert('更新対象がありません。');
    return;
  }

  const headers = values[0];
  const indexes = {};
  headers.forEach((header, index) => indexes[header] = index);
  const ui = SpreadsheetApp.getUi();
  const limitResponse = ui.prompt(
    '最大更新件数',
    '今回更新する最大件数を入力してください。まずは 20〜50 推奨です。',
    ui.ButtonSet.OK_CANCEL
  );
  if (limitResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const maxUpdates = asInteger_(limitResponse.getResponseText().trim() || 20, '最大更新件数');
  if (maxUpdates < 1) {
    ui.alert('最大更新件数は1以上で入力してください。');
    return;
  }
  let updated = 0;
  let skipped = 0;

  values.slice(1).forEach((rowValues, index) => {
    if (updated >= maxUpdates) {
      return;
    }
    const rowNumber = index + 2;
    if (rowValues[indexes.approve] !== true) {
      return;
    }
    if (rowValues[indexes.status] === 'TRADING_UPDATED') {
      skipped++;
      return;
    }

    try {
      const row = {
        listingId: rowValues[indexes.listingId],
        sku: rowValues[indexes.sku],
        priceUSD: rowValues[indexes.priceUSD],
        dutyRate: rowValues[indexes.dutyRate],
        fulfillmentPolicyId: rowValues[indexes.targetFulfillmentPolicyId],
        overrideShippingCostUSD: rowValues[indexes.overrideShippingCostUSD]
      };
      if (!row.listingId || !row.fulfillmentPolicyId) {
        throw new Error('listingIdまたはtargetFulfillmentPolicyIdが空です。');
      }

      const item = getTradingItem_(row.listingId, props);
      if (row.priceUSD === '' && item.price) {
        row.priceUSD = roundMoney_(item.price);
        writeBulkCell_(sheet, rowNumber, indexes.priceUSD + 1, row.priceUSD);
      }
      const requestXml = buildReviseFixedPriceItemXml_(row, item, props);
      const response = reviseFixedPriceItem_(requestXml, props);
      writeBulkCell_(sheet, rowNumber, indexes.status + 1, 'TRADING_UPDATED');
      writeBulkCell_(sheet, rowNumber, indexes.lastError + 1, '');
      writeBulkCell_(sheet, rowNumber, indexes.requestPreview + 1, requestXml + '\n\n--- RESPONSE ---\n' + response);
      updated++;
    } catch (err) {
      writeBulkCell_(sheet, rowNumber, indexes.status + 1, 'ERROR');
      writeBulkCell_(sheet, rowNumber, indexes.lastError + 1, String(err && err.stack ? err.stack : err));
    }
  });

  SpreadsheetApp.getUi().alert(updated + '件をTrading APIで更新しました。更新済みスキップ: ' + skipped + '件。残りがあれば同じメニューを再実行してください。');
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

function dryRunTradingSelectedRows() {
  processTradingRows_({ selectedOnly: true, updateEbay: false });
}

function updateTradingSelectedRows() {
  processTradingRows_({ selectedOnly: true, updateEbay: true });
}

function processTradingRows_(options) {
  const sheet = getSheet_();
  const rows = getRowsToProcess_(sheet, options.selectedOnly);
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('処理対象の行がありません。');
    return;
  }

  const props = getConfig_();
  rows.forEach(rowNumber => {
    try {
      const row = readRowForTrading_(sheet, rowNumber);
      const item = getTradingItem_(row.listingId, props);
      fillRowFromTradingItem_(sheet, rowNumber, row, item);
      const requestXml = buildReviseFixedPriceItemXml_(row, item, props);
      writeCell_(sheet, rowNumber, 'requestPreview', requestXml);

      if (!options.updateEbay) {
        writeCell_(sheet, rowNumber, 'status', 'TRADING_DRY_RUN_OK');
        writeCell_(sheet, rowNumber, 'lastError', '');
        return;
      }

      const response = reviseFixedPriceItem_(requestXml, props);
      writeCell_(sheet, rowNumber, 'status', 'TRADING_UPDATED');
      writeCell_(sheet, rowNumber, 'lastError', '');
      writeCell_(sheet, rowNumber, 'requestPreview', requestXml + '\n\n--- RESPONSE ---\n' + response);
    } catch (err) {
      writeCell_(sheet, rowNumber, 'status', 'ERROR');
      writeCell_(sheet, rowNumber, 'lastError', String(err && err.stack ? err.stack : err));
    }
  });
}

function migrateSelectedRowsToInventory() {
  migrateRowsToInventory_({ selectedOnly: true });
}

function migrateApprovedRowsToInventory() {
  migrateRowsToInventory_({ selectedOnly: false });
}

function migrateRowsToInventory_(options) {
  const sheet = getSheet_();
  const rows = getRowsToProcess_(sheet, options.selectedOnly);
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('処理対象の行がありません。');
    return;
  }

  const props = getPolicyFetchConfig_();
  rows.forEach(rowNumber => {
    try {
      const row = readRowForMigration_(sheet, rowNumber);
      const response = bulkMigrateListing_(row.listingId, props);
      const result = getMigrationResult_(response, row.listingId);
      if (result.errors && result.errors.length) {
        throw new Error('Migration failed: ' + JSON.stringify(result.errors));
      }

      const inventoryItem = getFirstMigratedInventoryItem_(result);
      if (!inventoryItem.offerId) {
        throw new Error('移行レスポンスにofferIdがありません: ' + JSON.stringify(result));
      }

      writeCell_(sheet, rowNumber, 'offerId', inventoryItem.offerId);
      if (!row.sku && inventoryItem.sku) {
        writeCell_(sheet, rowNumber, 'sku', inventoryItem.sku);
      }
      writeCell_(sheet, rowNumber, 'status', 'MIGRATED');
      writeCell_(sheet, rowNumber, 'lastError', '');
      writeCell_(sheet, rowNumber, 'requestPreview', JSON.stringify(response, null, 2));
    } catch (err) {
      writeCell_(sheet, rowNumber, 'status', 'ERROR');
      writeCell_(sheet, rowNumber, 'lastError', String(err && err.stack ? err.stack : err));
    }
  });
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

      fillRowFromExistingOffer_(sheet, rowNumber, row, existingOffer);
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

function fillRowFromExistingOffer_(sheet, rowNumber, row, offer) {
  const existingPrice = getExistingOfferPrice_(offer);
  if (row.priceUSD === '' && existingPrice !== '') {
    row.priceUSD = existingPrice;
    writeCell_(sheet, rowNumber, 'priceUSD', existingPrice);
  }

  const existingPolicyId = getFulfillmentPolicyIdFromOffer_(offer);
  if (row.fulfillmentPolicyId === '' && existingPolicyId !== '') {
    row.fulfillmentPolicyId = existingPolicyId;
    writeCell_(sheet, rowNumber, 'fulfillmentPolicyId', existingPolicyId);
  }
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

function bulkMigrateListing_(listingId, props) {
  return ebayFetch_('/sell/inventory/v1/bulk_migrate_listing', {
    method: 'post',
    payload: {
      requests: [{
        listingId: String(listingId)
      }]
    }
  }, props);
}

function getMigrationResult_(response, listingId) {
  const responses = response.responses || [];
  if (responses.length === 0) {
    throw new Error('移行レスポンスが空です: ' + JSON.stringify(response));
  }
  const match = responses.filter(item => String(item.listingId) === String(listingId))[0];
  return match || responses[0];
}

function getFirstMigratedInventoryItem_(result) {
  const items = result.inventoryItems || [];
  if (items.length === 0) {
    throw new Error('移行レスポンスにinventoryItemsがありません: ' + JSON.stringify(result));
  }
  return items[0];
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

function getExistingOfferPrice_(offer) {
  const price = offer && offer.pricingSummary ? offer.pricingSummary.price : null;
  if (price && price.value !== '' && typeof price.value !== 'undefined') {
    return roundMoney_(price.value);
  }
  return '';
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

function getTradingItem_(listingId, props) {
  const xml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<Version>' + getTradingApiVersion_(props) + '</Version>' +
    '<DetailLevel>ReturnAll</DetailLevel>' +
    '<ItemID>' + escapeXml_(listingId) + '</ItemID>' +
    '</GetItemRequest>';
  const responseText = tradingFetch_('GetItem', xml, props);
  const doc = XmlService.parse(responseText);
  assertTradingAck_(doc, responseText);
  return parseTradingItem_(doc);
}

function reviseFixedPriceItem_(requestXml, props) {
  const responseText = tradingFetch_('ReviseFixedPriceItem', requestXml, props);
  const doc = XmlService.parse(responseText);
  assertTradingAck_(doc, responseText);
  return responseText;
}

function getAllTradingActiveItems_(props, maxItems) {
  const items = [];
  const entriesPerPage = 200;
  let pageNumber = 1;
  while (true) {
    const page = getTradingActiveListPage_(pageNumber, entriesPerPage, props);
    page.items.forEach(item => {
      if (!maxItems || items.length < maxItems) {
        items.push(item);
      }
    });
    if ((maxItems && items.length >= maxItems) || pageNumber >= page.totalPages || page.items.length === 0) {
      break;
    }
    pageNumber++;
  }
  return items;
}

function getTradingActiveListPage_(pageNumber, entriesPerPage, props) {
  const xml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<GetMyeBaySellingRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<Version>' + getTradingApiVersion_(props) + '</Version>' +
    '<DetailLevel>ReturnAll</DetailLevel>' +
    '<ActiveList>' +
    '<Include>true</Include>' +
    '<Pagination>' +
    '<EntriesPerPage>' + entriesPerPage + '</EntriesPerPage>' +
    '<PageNumber>' + pageNumber + '</PageNumber>' +
    '</Pagination>' +
    '</ActiveList>' +
    '</GetMyeBaySellingRequest>';
  const responseText = tradingFetch_('GetMyeBaySelling', xml, props);
  const doc = XmlService.parse(responseText);
  assertTradingAck_(doc, responseText);
  return parseTradingActiveListPage_(doc);
}

function tradingFetch_(callName, requestXml, props) {
  const url = getTradingApiBase_(props);
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'text/xml; charset=UTF-8',
    headers: {
      'X-EBAY-API-COMPATIBILITY-LEVEL': getTradingApiVersion_(props),
      'X-EBAY-API-CALL-NAME': callName,
      'X-EBAY-API-SITEID': getTradingSiteId_(props),
      'X-EBAY-API-IAF-TOKEN': getValidAccessToken_(props)
    },
    payload: requestXml,
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(callName + ' failed: HTTP ' + code + ' ' + text);
  }
  return text;
}

function parseTradingItem_(doc) {
  const root = doc.getRootElement();
  const ns = root.getNamespace();
  const item = child_(root, ns, 'Item');
  if (!item) {
    throw new Error('GetItem responseにItemがありません。');
  }

  const sellingStatus = child_(item, ns, 'SellingStatus');
  const currentPrice = sellingStatus ? child_(sellingStatus, ns, 'CurrentPrice') : null;
  const sellerProfiles = child_(item, ns, 'SellerProfiles');
  const paymentProfile = sellerProfiles ? child_(sellerProfiles, ns, 'SellerPaymentProfile') : null;
  const returnProfile = sellerProfiles ? child_(sellerProfiles, ns, 'SellerReturnProfile') : null;
  const shippingProfile = sellerProfiles ? child_(sellerProfiles, ns, 'SellerShippingProfile') : null;

  return {
    itemId: textChild_(item, ns, 'ItemID'),
    sku: textChild_(item, ns, 'SKU'),
    price: currentPrice ? currentPrice.getText() : '',
    paymentProfileId: paymentProfile ? textChild_(paymentProfile, ns, 'PaymentProfileID') : '',
    returnProfileId: returnProfile ? textChild_(returnProfile, ns, 'ReturnProfileID') : '',
    shippingProfileId: shippingProfile ? textChild_(shippingProfile, ns, 'ShippingProfileID') : ''
  };
}

function parseTradingActiveListPage_(doc) {
  const root = doc.getRootElement();
  const ns = root.getNamespace();
  const activeList = child_(root, ns, 'ActiveList');
  if (!activeList) {
    return { items: [], totalPages: 1 };
  }

  const pagination = child_(activeList, ns, 'PaginationResult');
  const totalPages = Number(pagination ? textChild_(pagination, ns, 'TotalNumberOfPages') || 1 : 1);
  const itemArray = child_(activeList, ns, 'ItemArray');
  const itemElements = itemArray ? itemArray.getChildren('Item', ns) : [];
  const items = itemElements.map(item => parseTradingItemElement_(item, ns));
  return { items: items, totalPages: totalPages || 1 };
}

function parseTradingItemElement_(item, ns) {
  const sellingStatus = child_(item, ns, 'SellingStatus');
  const currentPrice = sellingStatus ? child_(sellingStatus, ns, 'CurrentPrice') : null;
  const sellerProfiles = child_(item, ns, 'SellerProfiles');
  const paymentProfile = sellerProfiles ? child_(sellerProfiles, ns, 'SellerPaymentProfile') : null;
  const returnProfile = sellerProfiles ? child_(sellerProfiles, ns, 'SellerReturnProfile') : null;
  const shippingProfile = sellerProfiles ? child_(sellerProfiles, ns, 'SellerShippingProfile') : null;

  return {
    itemId: textChild_(item, ns, 'ItemID'),
    sku: textChild_(item, ns, 'SKU'),
    price: currentPrice ? currentPrice.getText() : '',
    paymentProfileId: paymentProfile ? textChild_(paymentProfile, ns, 'PaymentProfileID') : '',
    returnProfileId: returnProfile ? textChild_(returnProfile, ns, 'ReturnProfileID') : '',
    shippingProfileId: shippingProfile ? textChild_(shippingProfile, ns, 'ShippingProfileID') : ''
  };
}

function buildReviseFixedPriceItemXml_(row, item, props) {
  const paymentProfileId = item.paymentProfileId || props.PAYMENT_POLICY_ID;
  const returnProfileId = item.returnProfileId || props.RETURN_POLICY_ID;
  const shippingProfileId = row.fulfillmentPolicyId || item.shippingProfileId || props.FULFILLMENT_POLICY_ID;
  if (!paymentProfileId || !returnProfileId || !shippingProfileId) {
    throw new Error('Trading API更新には payment/return/shipping のBusiness Policy IDが必要です。GetItemまたはScript Propertiesを確認してください。');
  }

  const shippingOverride = getShippingOverride_(row);
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<ReviseFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<Version>' + getTradingApiVersion_(props) + '</Version>' +
    '<Item>' +
    '<ItemID>' + escapeXml_(row.listingId) + '</ItemID>' +
    '<SellerProfiles>' +
    '<SellerPaymentProfile><PaymentProfileID>' + escapeXml_(paymentProfileId) + '</PaymentProfileID></SellerPaymentProfile>' +
    '<SellerReturnProfile><ReturnProfileID>' + escapeXml_(returnProfileId) + '</ReturnProfileID></SellerReturnProfile>' +
    '<SellerShippingProfile><ShippingProfileID>' + escapeXml_(shippingProfileId) + '</ShippingProfileID></SellerShippingProfile>' +
    '</SellerProfiles>' +
    '<ShippingServiceCostOverrideList>' +
    '<ShippingServiceCostOverride>' +
    '<ShippingServiceType>Domestic</ShippingServiceType>' +
    '<ShippingServicePriority>1</ShippingServicePriority>' +
    '<ShippingServiceCost currencyID="' + escapeXml_(props.CURRENCY) + '">' + escapeXml_(shippingOverride) + '</ShippingServiceCost>' +
    '</ShippingServiceCostOverride>' +
    '</ShippingServiceCostOverrideList>' +
    '</Item>' +
    '</ReviseFixedPriceItemRequest>'
  );
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
    CONTENT_LANGUAGE: store.getProperty('CONTENT_LANGUAGE') || 'en-US',
    TRADING_SITE_ID: store.getProperty('TRADING_SITE_ID'),
    TRADING_API_VERSION: store.getProperty('TRADING_API_VERSION') || '1455'
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

function getTradingApiBase_(props) {
  return props.ENVIRONMENT === 'SANDBOX' ? 'https://api.sandbox.ebay.com/ws/api.dll' : 'https://api.ebay.com/ws/api.dll';
}

function getTradingApiVersion_(props) {
  return props.TRADING_API_VERSION || '1455';
}

function getTradingSiteId_(props) {
  if (props.TRADING_SITE_ID) {
    return props.TRADING_SITE_ID;
  }
  return props.MARKETPLACE_ID === 'EBAY_MOTORS' ? '100' : '0';
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

function readRowForMigration_(sheet, rowNumber) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = {};
  headers.forEach((header, index) => {
    row[header] = values[index];
  });
  if (!row.listingId) {
    throw new Error('Inventory APIへ移行するには listingId が必須です。');
  }
  return row;
}

function readRowForTrading_(sheet, rowNumber) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = {};
  headers.forEach((header, index) => {
    row[header] = values[index];
  });
  if (!row.listingId) {
    throw new Error('Trading API更新には listingId が必須です。');
  }
  return row;
}

function fillRowFromTradingItem_(sheet, rowNumber, row, item) {
  if (row.sku === '' && item.sku) {
    row.sku = item.sku;
    writeCell_(sheet, rowNumber, 'sku', item.sku);
  }
  if (row.priceUSD === '' && item.price) {
    row.priceUSD = roundMoney_(item.price);
    writeCell_(sheet, rowNumber, 'priceUSD', row.priceUSD);
  }
  if (row.fulfillmentPolicyId === '' && item.shippingProfileId) {
    row.fulfillmentPolicyId = item.shippingProfileId;
    writeCell_(sheet, rowNumber, 'fulfillmentPolicyId', item.shippingProfileId);
  }
}

function validateRequiredForExistingUpdate_(row) {
  if (!row.sku && !row.offerId) {
    throw new Error('sku または offerId は必須です。');
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

function asInteger_(value, label) {
  const number = asNumber_(value, label);
  if (Math.floor(number) !== number) {
    throw new Error(label + ' must be an integer: ' + value);
  }
  return number;
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

function assertTradingAck_(doc, responseText) {
  const root = doc.getRootElement();
  const ns = root.getNamespace();
  const ack = textChild_(root, ns, 'Ack');
  if (ack === 'Success' || ack === 'Warning') {
    return;
  }
  throw new Error('Trading API failed: ' + responseText);
}

function child_(element, ns, name) {
  return element ? element.getChild(name, ns) : null;
}

function textChild_(element, ns, name) {
  const child = child_(element, ns, name);
  return child ? child.getText() : '';
}

function escapeXml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
