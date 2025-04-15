/**
 * @fileoverview Googleフォームからの申請に基づき、Grafana Cloudトークンを発行・通知するGASスクリプト。
 * @version 1.0.0
 */

// --- 定数定義 ---

// スプレッドシートの列設定（A列=1から始まるインデックス）
// !!! 重要: 実際のフォームとスプレッドシートに合わせて列番号を調整してください !!!
const COLUMN_EMAIL = 2; // 申請者のメールアドレスが自動収集される列 (フォームの設定によるが通常B列)
const COLUMN_EXPIRATION = 3; // フォームで有効期限を選択する質問の列 (例: C列)
const COLUMN_STATUS = 5; // 処理ステータスを書き込む列 (例: E列)
const COLUMN_TOKEN_NAME = 6; // 発行トークン名を書き込む列 (例: F列)
const COLUMN_EXPIRES_AT = 7; //  (例: G列)
const COLUMN_ERROR_DETAILS = 8; // エラー詳細を書き込む列 (例: H列)

// デフォルトのトークン有効期限（日数）
const DEFAULT_EXPIRATION_DAYS = 30;

// Grafana API 設定
const GRAFANA_CLOUD_API_ENDPOINT = 'https://www.grafana.com/api/v1/tokens';

// スクリプトプロパティのキー
const PROP_KEYS = {
  GRAFANA_CLOUD_API_KEY: 'GRAFANA_CLOUD_API_KEY',
  GRAFANA_CLOUD_ACCESS_POLICY_ID: 'GRAFANA_CLOUD_ACCESS_POLICY_ID',
  GRAFANA_CLOUD_REGION: 'GRAFANA_CLOUD_REGION',
  SUCCESS_EMAIL_FROM: 'SUCCESS_EMAIL_FROM',
  SUCCESS_EMAIL_NAME: 'SUCCESS_EMAIL_NAME',
  ADMIN_EMAIL: 'ADMIN_EMAIL', // エラー通知先管理者のメールアドレス
};

// --- メイン処理 ---

/**
 * Googleフォーム送信時に実行されるトリガー関数。
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e イベントオブジェクト
 */
function onFormSubmit(e) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = getScriptProperties(scriptProperties);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = e.range; // フォーム回答が書き込まれた範囲
  const row = range.getRow();

  Logger.log(`Processing form submission for row: ${row}`);
  Logger.log(`Event object: ${JSON.stringify(e, null, 2)}`);

  try {
    // --- 1. 申請者メールアドレス取得 (FR-02) ---
    let respondentEmail = '';
    // e.response が存在するかチェック
    if (e.response && typeof e.response.getRespondentEmail === 'function') {
        respondentEmail = e.response.getRespondentEmail();
    }
    // Fallback: フォーム設定によっては e.namedValues['メールアドレス'] などになる場合がある
    if (!respondentEmail && e.namedValues && e.namedValues['メールアドレス']) {
        respondentEmail = e.namedValues['メールアドレス'][0]; // 'メールアドレス'はフォームの質問名に合わせる
    }
    // Fallback: スプレッドシートから直接読み取る（要調整）
    if (!respondentEmail && sheet.getLastColumn() >= COLUMN_EMAIL) {
        respondentEmail = sheet.getRange(row, COLUMN_EMAIL).getValue();
    }

    if (!respondentEmail || !validateEmail(respondentEmail)) {
      throw new Error('有効な申請者メールアドレスを取得できませんでした。');
    }
    Logger.log(`Respondent Email: ${respondentEmail}`);

    // --- 2. 有効期限の取得と計算 (FR-02.1) ---
    let expirationString = '';
    // e.namedValues を使用してフォームの回答を取得（質問名を指定）
    if (sheet.getLastColumn() >= COLUMN_EXPIRATION) {
        // Fallback: スプレッドシートから直接読み取る
        expirationString = sheet.getRange(row, COLUMN_EXPIRATION).getValue();
    }
    Logger.log(`Expiration String from Form/Sheet: ${expirationString}`);

    const expiresAtISO = calculateExpirationDateISO(expirationString);
    Logger.log(`Calculated expiration (ISO): ${expiresAtISO}`);

    // --- 3. トークン名の生成 (FR-03) ---
    const tokenName = generateTokenName(respondentEmail);
    Logger.log(`Generated Token Name: ${tokenName}`);

    // --- 4. Grafana API 呼び出し (FR-03, FR-04) ---
    const grafanaResponse = createGrafanaToken(
      properties.apiKey,
      properties.region,
      properties.accessPolicyId,
      tokenName,
      expiresAtISO
    );

    // --- 5. 成功時処理 (FR-05) ---
    const tokenKey = grafanaResponse.token;
    const createdTokenName = grafanaResponse.name;
    const expiresAt = grafanaResponse.expiresAt;

    // 5a. 成功メール送信 (FR-05-2, FR-05-3)
    sendSuccessEmail(respondentEmail, tokenKey, createdTokenName, expiresAt, properties.emailFrom, properties.emailName);
    Logger.log(`Success email sent to ${respondentEmail}`);

    // 5b. スプレッドシート記録 (FR-07)
    updateSpreadsheet(sheet, row, '成功', createdTokenName, expiresAt, '');
    Logger.log(`Spreadsheet updated for row ${row} with status: 成功`);

    // 5c. 成功ログ (FR-05-1)
    Logger.log(`Token successfully created. Name: ${createdTokenName}`);

  } catch (error) {
    // --- 6. 失敗時処理 (FR-06) ---
    Logger.log(`Error occurred: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);

    let respondentEmailForError = '不明';
    try {
        // エラー発生前にメールアドレスが取得できていればそれを使う
        if (typeof respondentEmail !== 'undefined' && respondentEmail) {
            respondentEmailForError = respondentEmail;
        } else if (e.response && typeof e.response.getRespondentEmail === 'function') {
            respondentEmailForError = e.response.getRespondentEmail();
        } else if (e.namedValues && e.namedValues['メールアドレス']) {
             respondentEmailForError = e.namedValues['メールアドレス'][0];
        }
    } catch (e) {
        Logger.log('Could not retrieve respondent email even for error reporting.');
    }

    // 6a. エラーログ記録 (FR-06-1) - Logger.log で実施済み

    // 6b. スプレッドシート記録 (FR-07)
    try {
      updateSpreadsheet(sheet, row, '失敗', '', '', `エラー: ${error.message}`);
      Logger.log(`Spreadsheet updated for row ${row} with status: 失敗`);
    } catch (sheetError) {
      Logger.log(`Failed to update spreadsheet with error details: ${sheetError.message}`);
    }

    // 6c. 管理者へのエラー通知 (FR-06-2 - 推奨)
    if (properties.adminEmail) {
      try {
        sendErrorEmail(properties.adminEmail, error, respondentEmailForError, row);
        Logger.log(`Error notification sent to admin: ${properties.adminEmail}`);
      } catch (mailError) {
        Logger.log(`Failed to send error email to admin: ${mailError.message}`);
      }
    }
  } finally {
    Logger.log(`Processing finished for row: ${row}`);
    // 必要であればここでリソース解放などを行う
  }
}

// --- ヘルパー関数 ---

/**
 * スクリプトプロパティを取得・検証する (NFR-01, NFR-04)
 * @param {GoogleAppsScript.Properties.Properties} scriptProperties
 * @returns {{apiKey: string, accessPolicyId: string, adminEmail: string|null}}
 */
function getScriptProperties(scriptProperties) {
  const apiKey = scriptProperties.getProperty(PROP_KEYS.GRAFANA_CLOUD_API_KEY);
  const accessPolicyId = scriptProperties.getProperty(PROP_KEYS.GRAFANA_CLOUD_ACCESS_POLICY_ID);
  const region = scriptProperties.getProperty(PROP_KEYS.GRAFANA_CLOUD_REGION);
  const emailFrom = scriptProperties.getProperty(PROP_KEYS.SUCCESS_EMAIL_FROM);
  const emailName = scriptProperties.getProperty(PROP_KEYS.SUCCESS_EMAIL_NAME);
  const adminEmail = scriptProperties.getProperty(PROP_KEYS.ADMIN_EMAIL);

  if (!apiKey) {
    throw new Error(`スクリプトプロパティ '${PROP_KEYS.GRAFANA_CLOUD_API_KEY}' が設定されていません。`);
  }
  if (!accessPolicyId) {
    throw new Error(`スクリプトプロパティ '${PROP_KEYS.GRAFANA_CLOUD_ACCESS_POLICY_ID}' が設定されていません。`);
  }
  if (!region) {
    throw new Error(`スクリプトプロパティ '${PROP_KEYS.GRAFANA_CLOUD_REGION}' が設定されていません。`);
  }
  if (!adminEmail) {
    Logger.log(`スクリプトプロパティ '${PROP_KEYS.ADMIN_EMAIL}' が設定されていません。エラー発生時の管理者通知は行われません。`);
  }

  return { apiKey, accessPolicyId, region, emailFrom, emailName, adminEmail };
}

/**
 * 有効期限文字列（例: "90日"）を解析し、未来の日付のISO 8601形式文字列を返す (FR-02.1)
 * @param {string} expirationString フォームから取得した有効期限文字列
 * @returns {string} ISO 8601形式 (YYYY-MM-DDTHH:mm:ssZ) の有効期限日時文字列
 */
function calculateExpirationDateISO(expirationString) {
    let daysToAdd = DEFAULT_EXPIRATION_DAYS; // デフォルト値
    try {
        if (expirationString && typeof expirationString === 'string') {
            const match = expirationString.match(/(\d+)\s*日/); // "〇〇日" の形式を期待
            if (match && match[1]) {
                const parsedDays = parseInt(match[1], 10);
                if (!isNaN(parsedDays) && parsedDays > 0) {
                    daysToAdd = parsedDays;
                    Logger.log(`Parsed expiration days: ${daysToAdd} from string: "${expirationString}"`);
                } else {
                    Logger.log(`Invalid number format in expiration string: "${expirationString}". Using default: ${DEFAULT_EXPIRATION_DAYS} days.`);
                }
            } else {
                Logger.log(`Expiration string format not recognized: "${expirationString}". Using default: ${DEFAULT_EXPIRATION_DAYS} days.`);
            }
        } else {
             Logger.log(`Expiration string is empty or not a string. Using default: ${DEFAULT_EXPIRATION_DAYS} days.`);
        }
    } catch (e) {
        Logger.log(`Error parsing expiration string: "${expirationString}". Using default: ${DEFAULT_EXPIRATION_DAYS} days. Error: ${e.message}`);
    }


    const expirationDate = new Date();
    expirationDate.setDate(expirationDate.getDate() + daysToAdd);
    expirationDate.setHours(0, 0, 0, 0); // 日付の変わり目（UTC 00:00:00）に設定する場合など

    // UTC の ISO 8601 形式 (YYYY-MM-DDTHH:mm:ssZ) に変換
    return expirationDate.toISOString();
}


/**
 * メールアドレスのローカルパートからトークン名を生成する (FR-03)
 * @param {string} email メールアドレス
 * @returns {string} トークン名 (例: "user.name-timestamp")
 */
function generateTokenName(email) {
  const localPart = email.split('@')[0];
  // トークン名として使える文字種に制限がある場合、サニタイズが必要
  const sanitizedLocalPart = localPart.replace(/[^a-zA-Z0-9._-]/g, '_');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  return `${sanitizedLocalPart}-${timestamp}`;
}

/**
 * Grafana Cloud API を呼び出してトークンを作成する (FR-03, FR-04, NFR-02)
 * @param {string} apiKey Grafana APIキー
 * @param {string} region Region of the Access Policy
 * @param {string} accessPolicyId アクセスポリシーID (API仕様による)
 * @param {string} tokenName 生成するトークン名
 * @param {string} expiresAtISO ISO 8601形式の有効期限
 * @returns {object} APIレスポンスのJSONオブジェクト (例: { id: '...', name: '...', key: '...', ... })
 */
function createGrafanaToken(apiKey, region, accessPolicyId, tokenName, expiresAtISO) {
   const apiUrl = `${GRAFANA_CLOUD_API_ENDPOINT.split('?')[0]}?region=${region}`;
   const payload = {
     name: tokenName,
     expiresAt: expiresAtISO,
     accessPolicyId: accessPolicyId
   };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Accept': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // これでエラーレスポンスも取得できる (NFR-02)
  };

  Logger.log(`Calling Grafana API. URL: ${apiUrl}, Payload: ${JSON.stringify(payload)}`);

  let response;
  try {
    response = UrlFetchApp.fetch(apiUrl, options);
  } catch (e) {
    // ネットワークエラーなど fetch 自体の失敗 (NFR-02)
    Logger.log(`UrlFetchApp.fetch failed: ${e.message}`);
    throw new Error(`Grafana API への接続に失敗しました: ${e.message}`);
  }

  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  Logger.log(`Grafana API Response Code: ${responseCode}`);
  Logger.log(`Grafana API Response Body: ${responseBody}`);

  if (responseCode >= 200 && responseCode < 300) {
    // 成功 (FR-05)
    try {
      return JSON.parse(responseBody);
    } catch (e) {
      Logger.log(`Failed to parse Grafana API response JSON: ${e.message}`);
      throw new Error('Grafana APIからの成功応答の解析に失敗しました。');
    }
  } else {
    // 失敗 (FR-06)
    let errorMessage = `Grafana APIエラー (HTTP ${responseCode})`;
    try {
        // エラーレスポンスに詳細が含まれている場合がある
        const errorJson = JSON.parse(responseBody);
        if (errorJson.message) {
            errorMessage += `: ${errorJson.message}`;
        } else if (errorJson.error) {
             errorMessage += `: ${errorJson.error}`;
        } else {
            errorMessage += `. Response: ${responseBody}`;
        }
    } catch (e) {
        errorMessage += `. Response Body: ${responseBody}`;
    }
    Logger.log(`Grafana API call failed: ${errorMessage}`);
    throw new Error(errorMessage);
  }
}

/**
 * 成功通知メールを送信する (FR-05-2, FR-05-3)
 * @param {string} recipient 申請者のメールアドレス
 * @param {string} tokenKey 発行されたトークンキー
 * @param {string} tokenName 発行されたトークン名
 * @param {string} expiresAt
 * @param {string} emailFrom
 * @param {string} emailName
 */
function sendSuccessEmail(recipient, tokenKey, tokenName, expiresAt, emailFrom, emailName) {
  const emailNoReply = false;
  const subject = '【重要】Grafana Cloud トークン発行完了のお知らせ';
  const body = `
${recipient.split('@')[0]} 様

Grafana Cloud のトークン発行申請を受け付け、以下のトークンを発行しました。

トークン名: ${tokenName}
トークンキー: ${tokenKey}
有効期限: ${expiresAt}

--- 重要事項 ---
- このトークンキーは 機密性の高い情報です。パスワードと同様に扱ってください。
- このトークンキーは 組織外に共有しないでください。
- このトークンキーは管理者からも確認できません。紛失した場合は再発行が必要です。
- このトークンは指定された有効期限後に自動的に失効します。

不明な点があれば管理者までお問い合わせください。
`;

  const message = {
      to: recipient,
      subject: subject,
      body: body,
  };
  if (emailName) {
    message.name = emailName;
  }
  if (emailFrom) {
    message.from = emailFrom;
  }
  if (emailNoReply) {
    message.noReply = noReply;
  }
  try {
    MailApp.sendEmail(message);
  } catch (e) {
    Logger.log(`Failed to send success email to ${recipient}: ${e.message}`);
    // ここでエラーを投げると、API成功->メール失敗の場合に処理全体が失敗扱いになる。
    // 必要に応じて、メール送信失敗をスプレッドシートに記録するなどの処理を追加。
    throw new Error(`成功通知メールの送信に失敗しました: ${e.message}`);
  }
}

/**
 * 管理者へエラー通知メールを送信する (FR-06-2)
 * @param {string} adminEmail 管理者のメールアドレス
 * @param {Error} error 発生したエラーオブジェクト
 * @param {string} respondentEmail 申請者のメールアドレス（取得できていれば）
 * @param {number} row エラーが発生したスプレッドシートの行番号
 */
function sendErrorEmail(adminEmail, error, respondentEmail, row) {
  const subject = `【警告】Grafanaトークン発行処理でエラーが発生しました`;
  const body = `
Grafana Cloud トークン発行処理中にエラーが発生しました。

発生日時: ${new Date().toLocaleString('ja-JP')}
対象シート: ${SpreadsheetApp.getActiveSpreadsheet().getName()} (${SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()})
対象行: ${row}
申請者メールアドレス (推定): ${respondentEmail}

エラー詳細:
${error.message}

スタックトレース:
${error.stack}

GASのログやスプレッドシートの該当行を確認してください。
`;

  try {
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: body,
    });
  } catch (e) {
    // 管理者へのメール送信失敗はログに残すのみ
    Logger.log(`Failed to send error notification email to admin ${adminEmail}: ${e.message}`);
  }
}

/**
 * スプレッドシートの該当行に処理結果を書き込む (FR-07)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} row 書き込む行番号
 * @param {string} status 処理ステータス ('成功' または '失敗')
 * @param {string} tokenName 発行されたトークン名 (成功時)
 * @param {Date} expiresAt トークン有効期限 (成功時)
 * @param {string} errorDetails エラー詳細 (失敗時)
 */
function updateSpreadsheet(sheet, row, status, tokenName, expiresAt, errorDetails) {
  try {
    // 書き込む列が存在するか確認（エラー防止）
    if (sheet.getMaxColumns() >= Math.max(COLUMN_STATUS, COLUMN_TOKEN_NAME, COLUMN_EXPIRES_AT, COLUMN_ERROR_DETAILS)) {
        sheet.getRange(row, COLUMN_STATUS).setValue(status);
        sheet.getRange(row, COLUMN_TOKEN_NAME).setValue(tokenName);
        sheet.getRange(row, COLUMN_EXPIRES_AT).setValue(expiresAt);
        sheet.getRange(row, COLUMN_ERROR_DETAILS).setValue(errorDetails);
    } else {
        Logger.log(`Error: One or more target columns for spreadsheet update do not exist. Max columns: ${sheet.getMaxColumns()}`);
        // 最低限ステータスだけでも書き込む試み
        if (sheet.getMaxColumns() >= COLUMN_STATUS) {
             sheet.getRange(row, COLUMN_STATUS).setValue(`${status} (書き込み列不足)`);
        }
    }
  } catch (e) {
    Logger.log(`Failed to update spreadsheet row ${row}: ${e.message}`);
    // スプレッドシートへの書き込み失敗はログに残すのみ（ここでエラーを投げると無限ループの可能性）
  }
}

/**
 * メールアドレス形式を簡易的に検証する
 * @param {string} email 検証する文字列
 * @returns {boolean} 有効な形式であれば true
 */
function validateEmail(email) {
    if (!email || typeof email !== 'string') {
        return false;
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}
