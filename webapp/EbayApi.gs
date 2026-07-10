// 利用者(生徒)ごとのeBay API処理層。
// Code.gs(スプレッドシート版)の実績あるロジックを、
// 「スクリプトプロパティ固定」から「利用者ごとの設定(Usersシート)」で動くように移植したもの。
// 関数名の wa プレフィックスは WebApp 用の意味。

const WA_ACCESS_TOKEN_CACHE_PREFIX = 'WA_EBAY_AT_';

// ---- 利用者ごとの設定 ----

function waGetUserEbayProps_(email) {
  const user = getWebAppUserByEmail_(email);
  if (!user) {
    throw new Error('Usersシートに利用者が見つかりません: ' + email);
  }
  if (!user.ebayRefreshToken) {
    throw new Error('eBayが未接続です。先に「eBay接続」タブで認証を完了してください。');
  }
  const app = getWebAppEbayAppConfig_();
  return {
    ENVIRONMENT: app.environment,
    EBAY_CLIENT_ID: app.clientId,
    EBAY_CLIENT_SECRET: app.clientSecret,
    EBAY_REFRESH_TOKEN: user.ebayRefreshToken,
    MARKETPLACE_ID: String(user.marketplaceId || 'EBAY_US').trim() || 'EBAY_US',
    CURRENCY: String(user.currency || 'USD').trim() || 'USD',
    CONTENT_LANGUAGE: String(user.contentLanguage || 'en-US').trim() || 'en-US',
    TRADING_SITE_ID: String(user.tradingSiteId || '').trim(),
    TRADING_API_VERSION: String(user.tradingApiVersion || '1455').trim() || '1455',
    email: email
  };
}

// ---- アクセストークン(利用者別にキャッシュ) ----

function waGetValidAccessToken_(props) {
  const cache = CacheService.getScriptCache();
  const cacheKey = WA_ACCESS_TOKEN_CACHE_PREFIX + props.email;
  const cached = cache.get(cacheKey);
  if (cached) {
    return cached;
  }

  const credentials = Utilities.base64Encode(props.EBAY_CLIENT_ID + ':' + props.EBAY_CLIENT_SECRET);
  const response = UrlFetchApp.fetch(waGetAuthBase_(props) + '/identity/v1/oauth2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      Authorization: 'Basic ' + credentials
    },
    payload: {
      grant_type: 'refresh_token',
      refresh_token: props.EBAY_REFRESH_TOKEN,
      scope: getWebAppEbayScopes_().join(' ')
    },
    muteHttpExceptions: true
  });

  const body = waParseEbayResponse_(response, 'eBayアクセストークン取得');
  if (!body.access_token) {
    throw new Error('eBayの認証に失敗しました。「eBay接続」タブで再認証してください。');
  }
  cache.put(cacheKey, body.access_token, Math.max(60, Number(body.expires_in || 7200) - 300));
  return body.access_token;
}

// ---- REST API (Inventory / Account) ----

function waEbayFetch_(path, request, props) {
  const options = {
    method: request.method,
    headers: {
      Authorization: 'Bearer ' + waGetValidAccessToken_(props),
      'Content-Language': props.CONTENT_LANGUAGE
    },
    muteHttpExceptions: true
  };
  if (request.payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(request.payload);
  }
  const response = UrlFetchApp.fetch(waGetApiBase_(props) + path, options);
  return waParseEbayResponse_(response, path);
}

function waParseEbayResponse_(response, label) {
  const code = response.getResponseCode();
  const text = response.getContentText();
  let body = {};
  if (text) {
    try {
      body = JSON.parse(text);
    } catch (err) {
      body = {};
    }
  }
  if (code < 200 || code >= 300) {
    throw new Error(label + ' failed: HTTP ' + code + ' ' + text);
  }
  return body;
}

// ---- ポリシー一覧 ----

function waFetchAllPolicies_(props) {
  const marketplace = encodeURIComponent(props.MARKETPLACE_ID);
  const fulfillment = (waEbayFetch_('/sell/account/v1/fulfillment_policy?marketplace_id=' + marketplace, { method: 'get' }, props).fulfillmentPolicies || [])
    .map(policy => ({ type: 'FULFILLMENT', name: policy.name || '', policyId: String(policy.fulfillmentPolicyId || '') }));
  const payment = (waEbayFetch_('/sell/account/v1/payment_policy?marketplace_id=' + marketplace, { method: 'get' }, props).paymentPolicies || [])
    .map(policy => ({ type: 'PAYMENT', name: policy.name || '', policyId: String(policy.paymentPolicyId || '') }));
  const returns = (waEbayFetch_('/sell/account/v1/return_policy?marketplace_id=' + marketplace, { method: 'get' }, props).returnPolicies || [])
    .map(policy => ({ type: 'RETURN', name: policy.name || '', policyId: String(policy.returnPolicyId || '') }));
  return {
    fulfillment: fulfillment,
    payment: payment,
    returns: returns,
    marketplaceId: props.MARKETPLACE_ID
  };
}

// ---- Inventory API (SKU / offerId ベース) ----

function waGetOffer_(offerId, props) {
  return waEbayFetch_('/sell/inventory/v1/offer/' + encodeURIComponent(offerId), { method: 'get' }, props);
}

function waGetOffersBySku_(sku, props) {
  const query =
    '?sku=' + encodeURIComponent(sku) +
    '&marketplace_id=' + encodeURIComponent(props.MARKETPLACE_ID) +
    '&format=FIXED_PRICE';
  return waEbayFetch_('/sell/inventory/v1/offer' + query, { method: 'get' }, props);
}

function waUpdateOffer_(offerId, payload, props) {
  return waEbayFetch_('/sell/inventory/v1/offer/' + encodeURIComponent(offerId), {
    method: 'put',
    payload: payload
  }, props);
}

function waGetExistingOfferForRow_(row, props) {
  if (row.offerId) {
    return waGetOffer_(row.offerId, props);
  }
  if (!row.sku) {
    throw new Error('SKUまたはoffer IDを入力してください。');
  }
  const response = waGetOffersBySku_(row.sku, props);
  const offers = response.offers || [];
  if (offers.length === 0) {
    throw new Error('このSKUの出品が見つかりません: ' + row.sku);
  }
  if (offers.length > 1) {
    throw new Error('同じSKUで複数の出品が見つかりました。offer IDで指定してください: ' + row.sku);
  }
  return offers[0];
}

function waBuildExistingOfferUpdatePayload_(offer, row, props) {
  const shippingOverride = waGetShippingOverride_(row);
  const payload = waPickWritableOfferFields_(offer);
  payload.sku = payload.sku || row.sku;
  payload.marketplaceId = payload.marketplaceId || props.MARKETPLACE_ID;
  payload.format = payload.format || 'FIXED_PRICE';

  if (!payload.listingPolicies) {
    payload.listingPolicies = {};
  }
  payload.listingPolicies.fulfillmentPolicyId =
    row.fulfillmentPolicyId || payload.listingPolicies.fulfillmentPolicyId;
  payload.listingPolicies.shippingCostOverrides = [{
    priority: 1,
    shippingServiceType: 'DOMESTIC',
    shippingCost: {
      value: shippingOverride,
      currency: props.CURRENCY
    },
    additionalShippingCost: {
      value: waGetAdditionalShippingCost_(shippingOverride, row),
      currency: props.CURRENCY
    }
  }];

  return waRemoveEmptyValues_(payload);
}

function waPickWritableOfferFields_(offer) {
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
      payload[key] = JSON.parse(JSON.stringify(offer[key]));
    }
  });
  return payload;
}

function waGetExistingOfferPrice_(offer) {
  const price = offer && offer.pricingSummary ? offer.pricingSummary.price : null;
  if (price && price.value !== '' && typeof price.value !== 'undefined') {
    return waRoundMoney_(price.value);
  }
  return '';
}

function waGetFulfillmentPolicyIdFromOffer_(offer) {
  return offer && offer.listingPolicies ? String(offer.listingPolicies.fulfillmentPolicyId || '') : '';
}

function waGetListingIdFromOffer_(offer) {
  return offer && offer.listing && offer.listing.listingId ? String(offer.listing.listingId) : '';
}

// ---- Inventory APIへの移行 ----

function waBulkMigrateListing_(listingId, props) {
  return waEbayFetch_('/sell/inventory/v1/bulk_migrate_listing', {
    method: 'post',
    payload: {
      requests: [{ listingId: String(listingId) }]
    }
  }, props);
}

function waGetMigrationResult_(response, listingId) {
  const responses = response.responses || [];
  if (responses.length === 0) {
    throw new Error('移行の結果が空でした: ' + JSON.stringify(response));
  }
  const match = responses.filter(item => String(item.listingId) === String(listingId))[0];
  return match || responses[0];
}

// ---- Trading API (Item ID ベース) ----

function waTradingFetch_(callName, requestXml, props) {
  const response = UrlFetchApp.fetch(waGetTradingApiBase_(props), {
    method: 'post',
    contentType: 'text/xml; charset=UTF-8',
    headers: {
      'X-EBAY-API-COMPATIBILITY-LEVEL': props.TRADING_API_VERSION,
      'X-EBAY-API-CALL-NAME': callName,
      'X-EBAY-API-SITEID': waGetTradingSiteId_(props),
      'X-EBAY-API-IAF-TOKEN': waGetValidAccessToken_(props)
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

function waGetTradingItem_(listingId, props) {
  const xml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<Version>' + props.TRADING_API_VERSION + '</Version>' +
    '<DetailLevel>ReturnAll</DetailLevel>' +
    '<ItemID>' + waEscapeXml_(listingId) + '</ItemID>' +
    '</GetItemRequest>';
  const responseText = waTradingFetch_('GetItem', xml, props);
  const doc = XmlService.parse(responseText);
  waAssertTradingAck_(doc, responseText);
  const root = doc.getRootElement();
  const ns = root.getNamespace();
  const item = waChild_(root, ns, 'Item');
  if (!item) {
    throw new Error('出品情報を取得できませんでした: ' + listingId);
  }
  return waParseTradingItemElement_(item, ns);
}

function waGetTradingActiveListPage_(pageNumber, entriesPerPage, props) {
  const xml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<GetMyeBaySellingRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<Version>' + props.TRADING_API_VERSION + '</Version>' +
    '<DetailLevel>ReturnAll</DetailLevel>' +
    '<ActiveList>' +
    '<Include>true</Include>' +
    '<Pagination>' +
    '<EntriesPerPage>' + entriesPerPage + '</EntriesPerPage>' +
    '<PageNumber>' + pageNumber + '</PageNumber>' +
    '</Pagination>' +
    '</ActiveList>' +
    '</GetMyeBaySellingRequest>';
  const responseText = waTradingFetch_('GetMyeBaySelling', xml, props);
  const doc = XmlService.parse(responseText);
  waAssertTradingAck_(doc, responseText);

  const root = doc.getRootElement();
  const ns = root.getNamespace();
  const activeList = waChild_(root, ns, 'ActiveList');
  if (!activeList) {
    return { items: [], totalPages: 1 };
  }
  const pagination = waChild_(activeList, ns, 'PaginationResult');
  const totalPages = Number(pagination ? waTextChild_(pagination, ns, 'TotalNumberOfPages') || 1 : 1);
  const itemArray = waChild_(activeList, ns, 'ItemArray');
  const itemElements = itemArray ? itemArray.getChildren('Item', ns) : [];
  const items = itemElements.map(item => waParseTradingItemElement_(item, ns));
  return { items: items, totalPages: totalPages || 1 };
}

function waParseTradingItemElement_(item, ns) {
  const sellingStatus = waChild_(item, ns, 'SellingStatus');
  const currentPrice = sellingStatus ? waChild_(sellingStatus, ns, 'CurrentPrice') : null;
  const sellerProfiles = waChild_(item, ns, 'SellerProfiles');
  const paymentProfile = sellerProfiles ? waChild_(sellerProfiles, ns, 'SellerPaymentProfile') : null;
  const returnProfile = sellerProfiles ? waChild_(sellerProfiles, ns, 'SellerReturnProfile') : null;
  const shippingProfile = sellerProfiles ? waChild_(sellerProfiles, ns, 'SellerShippingProfile') : null;

  return {
    itemId: waTextChild_(item, ns, 'ItemID'),
    sku: waTextChild_(item, ns, 'SKU'),
    title: waTextChild_(item, ns, 'Title'),
    price: currentPrice ? currentPrice.getText() : '',
    paymentProfileId: paymentProfile ? waTextChild_(paymentProfile, ns, 'PaymentProfileID') : '',
    returnProfileId: returnProfile ? waTextChild_(returnProfile, ns, 'ReturnProfileID') : '',
    shippingProfileId: shippingProfile ? waTextChild_(shippingProfile, ns, 'ShippingProfileID') : ''
  };
}

function waBuildReviseFixedPriceItemXml_(row, item, props) {
  const paymentProfileId = item.paymentProfileId;
  const returnProfileId = item.returnProfileId;
  const shippingProfileId = row.fulfillmentPolicyId || item.shippingProfileId;
  if (!paymentProfileId || !returnProfileId || !shippingProfileId) {
    throw new Error('この出品はBusiness Policy(支払い/返品/配送ポリシー)が取得できないため更新できません。Item ID: ' + row.listingId);
  }

  const shippingOverride = waGetShippingOverride_(row);
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<ReviseFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">' +
    '<Version>' + props.TRADING_API_VERSION + '</Version>' +
    '<Item>' +
    '<ItemID>' + waEscapeXml_(row.listingId) + '</ItemID>' +
    '<SellerProfiles>' +
    '<SellerPaymentProfile><PaymentProfileID>' + waEscapeXml_(paymentProfileId) + '</PaymentProfileID></SellerPaymentProfile>' +
    '<SellerReturnProfile><ReturnProfileID>' + waEscapeXml_(returnProfileId) + '</ReturnProfileID></SellerReturnProfile>' +
    '<SellerShippingProfile><ShippingProfileID>' + waEscapeXml_(shippingProfileId) + '</ShippingProfileID></SellerShippingProfile>' +
    '</SellerProfiles>' +
    '<ShippingServiceCostOverrideList>' +
    '<ShippingServiceCostOverride>' +
    '<ShippingServiceType>Domestic</ShippingServiceType>' +
    '<ShippingServicePriority>1</ShippingServicePriority>' +
    '<ShippingServiceCost currencyID="' + waEscapeXml_(props.CURRENCY) + '">' + waEscapeXml_(shippingOverride) + '</ShippingServiceCost>' +
    '<ShippingServiceAdditionalCost currencyID="' + waEscapeXml_(props.CURRENCY) + '">' + waEscapeXml_(waGetAdditionalShippingCost_(shippingOverride, row)) + '</ShippingServiceAdditionalCost>' +
    '</ShippingServiceCostOverride>' +
    '</ShippingServiceCostOverrideList>' +
    '</Item>' +
    '</ReviseFixedPriceItemRequest>'
  );
}

function waReviseFixedPriceItem_(requestXml, props) {
  const responseText = waTradingFetch_('ReviseFixedPriceItem', requestXml, props);
  const doc = XmlService.parse(responseText);
  waAssertTradingAck_(doc, responseText);
  return responseText;
}

function waAssertTradingAck_(doc, responseText) {
  const root = doc.getRootElement();
  const ns = root.getNamespace();
  const ack = waTextChild_(root, ns, 'Ack');
  if (ack === 'Success' || ack === 'Warning') {
    return;
  }
  throw new Error('eBayがエラーを返しました: ' + waExtractTradingErrors_(root, ns, responseText));
}

function waExtractTradingErrors_(root, ns, fallbackText) {
  const errors = root.getChildren('Errors', ns) || [];
  const messages = errors
    .map(error => waTextChild_(error, ns, 'LongMessage') || waTextChild_(error, ns, 'ShortMessage'))
    .filter(message => message);
  return messages.length ? messages.join(' / ') : fallbackText;
}

// ---- 共通ヘルパー ----

// 送料 = 直接指定があればその金額。なければ 商品価格 × 関税率 + 固定追加額
function waGetShippingOverride_(row) {
  if (row.overrideShippingCostUSD !== '' && typeof row.overrideShippingCostUSD !== 'undefined' && row.overrideShippingCostUSD !== null) {
    return waRoundMoney_(waAsNumber_(row.overrideShippingCostUSD, '送料'));
  }
  const price = waAsNumber_(row.priceUSD, '商品価格');
  const dutyRate = (row.dutyRate === '' || typeof row.dutyRate === 'undefined' || row.dutyRate === null)
    ? 0
    : waAsNumber_(row.dutyRate, '関税率');
  const addFixed = (row.addFixedUSD === '' || typeof row.addFixedUSD === 'undefined' || row.addFixedUSD === null)
    ? 0
    : waAsNumber_(row.addFixedUSD, '固定追加額');
  return waRoundMoney_(price * dutyRate + addFixed);
}

// 同一商品2個目以降の追加送料 = 送料 × 割合(%)。空欄なら75%
function waGetAdditionalShippingCost_(shippingOverride, row) {
  const percent = (row.additionalRatePercent === '' || typeof row.additionalRatePercent === 'undefined' || row.additionalRatePercent === null)
    ? 75
    : waAsNumber_(row.additionalRatePercent, '追加送料の割合');
  if (percent < 0 || percent > 100) {
    throw new Error('追加送料の割合は0から100の数字で入力してください: ' + percent);
  }
  return waRoundMoney_(waAsNumber_(shippingOverride, '送料') * percent / 100);
}

function waAsNumber_(value, label) {
  const number = Number(value);
  if (!isFinite(number)) {
    throw new Error(label + ' は数値で入力してください: ' + value);
  }
  return number;
}

function waRoundMoney_(value) {
  return (Math.round(Number(value) * 100) / 100).toFixed(2);
}

function waRemoveEmptyValues_(value) {
  if (Array.isArray(value)) {
    return value.map(waRemoveEmptyValues_).filter(item => typeof item !== 'undefined');
  }
  if (value && typeof value === 'object') {
    const result = {};
    Object.keys(value).forEach(key => {
      const cleaned = waRemoveEmptyValues_(value[key]);
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

function waChild_(element, ns, name) {
  return element ? element.getChild(name, ns) : null;
}

function waTextChild_(element, ns, name) {
  const child = waChild_(element, ns, name);
  return child ? child.getText() : '';
}

function waEscapeXml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function waGetApiBase_(props) {
  return String(props.ENVIRONMENT || '').toUpperCase() === 'SANDBOX' ? 'https://api.sandbox.ebay.com' : 'https://api.ebay.com';
}

function waGetAuthBase_(props) {
  return waGetApiBase_(props);
}

function waGetTradingApiBase_(props) {
  return String(props.ENVIRONMENT || '').toUpperCase() === 'SANDBOX' ? 'https://api.sandbox.ebay.com/ws/api.dll' : 'https://api.ebay.com/ws/api.dll';
}

function waGetTradingSiteId_(props) {
  if (props.TRADING_SITE_ID) {
    return props.TRADING_SITE_ID;
  }
  return props.MARKETPLACE_ID === 'EBAY_MOTORS' ? '100' : '0';
}
