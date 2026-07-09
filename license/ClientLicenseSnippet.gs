const TOOL_ID = 'ebay-shipping-policy-tool';

function requireLicensedUser_() {
  const props = PropertiesService.getScriptProperties();
  const licenseUrl = props.getProperty('LICENSE_SERVER_URL');
  const licenseToken = props.getProperty('LICENSE_SERVER_TOKEN');
  const cachedEmail = props.getProperty('LICENSED_EMAIL');
  const cache = CacheService.getScriptCache();

  if (!licenseUrl || !licenseToken) {
    throw new Error('LICENSE_SERVER_URL と LICENSE_SERVER_TOKEN をスクリプトプロパティに設定してください。');
  }

  const email = normalizeLicenseEmail_(cachedEmail || Session.getActiveUser().getEmail());
  if (!email) {
    throw new Error('利用者メールアドレスを確認できません。先に saveLicensedEmail を実行してください。');
  }

  const cacheKey = 'LICENSE_OK_' + email;
  if (cache.get(cacheKey) === '1') {
    return email;
  }

  const query =
    '?email=' + encodeURIComponent(email) +
    '&toolId=' + encodeURIComponent(TOOL_ID) +
    '&token=' + encodeURIComponent(licenseToken);
  const response = UrlFetchApp.fetch(licenseUrl + query, {
    method: 'get',
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('ライセンス確認に失敗しました: HTTP ' + code + ' ' + text);
  }

  const result = JSON.parse(text);
  if (!result.ok) {
    throw new Error('このメールアドレスは利用許可されていません: ' + email + ' / reason=' + result.reason);
  }

  props.setProperty('LICENSED_EMAIL', email);
  cache.put(cacheKey, '1', 21600);
  return email;
}

function saveLicensedEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '利用者メールアドレス',
    '登録済みのメールアドレスを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  const email = normalizeLicenseEmail_(response.getResponseText());
  if (!email) {
    ui.alert('メールアドレスを入力してください。');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('LICENSED_EMAIL', email);
  CacheService.getScriptCache().remove('LICENSE_OK_' + email);
  requireLicensedUser_();
  ui.alert('利用者メールアドレスを確認しました: ' + email);
}

function normalizeLicenseEmail_(value) {
  return String(value || '').trim().toLowerCase();
}
