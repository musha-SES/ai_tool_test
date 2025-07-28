/**
 * スプレッドシートを開いたときにカスタムメニューを追加する関数。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報AIツール')
    .addItem('日報質問送信（Chatwork）', 'sendDailyReportQuestions')
    .addItem('Chatwork日報取得・分析', 'processChatworkReplies')
    .addSeparator()
    .addItem('週次レポート生成', 'generateWeeklyReports')
    .addSeparator()
    .addItem('定期実行トリガーを設定', 'createDailyTriggers')
    .addItem('全てのトリガーを削除', 'deleteTriggers')
    .addSeparator()
    .addItem('1on1ヒアリング項目生成', 'generate1on1Topics')
    .addToUi();
}

// --- トリガー管理機能 ---

/**
 * スクリプトプロパティから定期実行の時刻を取得する。
 * @returns {{questionHour: number, questionMinute: number, replyHour: number, replyMinute: number}}
 */
function getScheduledTimes() {
  const properties = PropertiesService.getScriptProperties();

  const getPropertyAsInt = (key, defaultValue) => {
    const prop = properties.getProperty(key);
    const value = parseInt(prop, 10);
    if (isNaN(value)) {
      Logger.log(`スクリプトプロパティ「${key}」が未設定または不正です。デフォルト値 (${defaultValue}) を使用します。`);
      return defaultValue;
    }
    return value;
  };

  const questionHour = getPropertyAsInt('DAILY_QUESTION_TIME_HOUR', CONFIG.SCHEDULE_DEFAULTS.questionHour);
  const questionMinute = getPropertyAsInt('DAILY_QUESTION_TIME_MINUTE', CONFIG.SCHEDULE_DEFAULTS.questionMinute);
  const replyHour = getPropertyAsInt('DAILY_REPLY_COLLECT_TIME_HOUR', CONFIG.SCHEDULE_DEFAULTS.replyHour);
  const replyMinute = getPropertyAsInt('DAILY_REPLY_COLLECT_TIME_MINUTE', CONFIG.SCHEDULE_DEFAULTS.replyMinute);

  return { questionHour, questionMinute, replyHour, replyMinute };
}

/**
 * 毎日の定期実行トリガーを作成する。
 */
function createDailyTriggers() {
  deleteTriggers();
  const { questionHour, questionMinute, replyHour, replyMinute } = getScheduledTimes();

  ScriptApp.newTrigger('sendDailyReportQuestions')
    .timeBased()
    .everyDays(1)
    .atHour(questionHour)
    .nearMinute(questionMinute)
    .create();
  Logger.log(`日報質問送信トリガーを毎日 ${questionHour}時${questionMinute}分頃に設定しました。`);

  ScriptApp.newTrigger('processChatworkReplies')
    .timeBased()
    .everyDays(1)
    .atHour(replyHour)
    .nearMinute(replyMinute)
    .create();
  Logger.log(`Chatwork日報取得・分析トリガーを毎日 ${replyHour}時${replyMinute}分頃に設定しました。`);

  ScriptApp.newTrigger('generateWeeklyReports')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(CONFIG.TRIGGER_SETTINGS.WEEKLY_REPORT_HOUR)
    .create();
  Logger.log(`週次レポート生成トリガーを毎週月曜日${CONFIG.TRIGGER_SETTINGS.WEEKLY_REPORT_HOUR}時に設定しました。`);

  ScriptApp.newTrigger('cleanUpBotQuestionLog')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(CONFIG.TRIGGER_SETTINGS.CLEANUP_HOUR)
    .create();
  Logger.log(`BOT質問ログクリーンアップトリガーを毎週月曜日深夜${CONFIG.TRIGGER_SETTINGS.CLEANUP_HOUR}時に設定しました。`);

  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました。 質問送信: 毎日' + `${questionHour}時${questionMinute}分頃` + ' 返信収集: 毎日' + `${replyHour}時${replyMinute}分頃` + ' 週次レポート: 毎週月曜日10時' + ' ログクリーンアップ: 毎週月曜日深夜');
}

/**
 * BOT質問ログシートをクリーンアップする。
 */
function cleanUpBotQuestionLog() {
  const logSheetName = CONFIG.BOT_QUESTION_LOG_SHEET_NAME;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(logSheetName);
  if (!sheet) {
    Logger.log(`シート「${logSheetName}」が見つかりません。クリーンアップをスキップします。`);
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) {
    return;
  }

  const rowsToDelete = [];
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() - CONFIG.TRIGGER_SETTINGS.CLEANUP_LOG_OLDER_THAN_DAYS);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[3].toString();
    const timestamp = new Date(row[2]);

    if (status === CONFIG.STATUS_STRINGS.SUCCESS || ((status === CONFIG.STATUS_STRINGS.PENDING || status === CONFIG.STATUS_STRINGS.INVALID_FORMAT || status.startsWith(CONFIG.STATUS_STRINGS.ERROR_PREFIX)) && timestamp < targetDate)) {
      rowsToDelete.push(i + 1);
    }
  }

  rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => sheet.deleteRow(rowIndex));
  Logger.log(`BOT質問ログシートのクリーンアップが完了しました。${rowsToDelete.length}件のレコードを削除しました。`);
}

/**
 * このスクリプトで設定されたすべてのトリガーを削除する。
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log('すべてのトリガーを削除しました。');
}

// --- Chatwork連携機能 ---

/**
 * Chatwork APIキーをスクリプトプロパティから取得する。
 */
function getChatworkApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.CHATWORK_API_KEY);
  if (!apiKey) {
    throw new Error('Chatwork API Key is not set in Script Properties.');
  }
  return apiKey;
}

/**
 * 「Chatwork設定」シートから全社員の情報をフラットな配列として取得する。
 * @returns {Array<{employeeName: string, employeeRoomId: string, managerName: string, managerRoomId: string, weeklyReportMode: string, group: string, role: string}>}
 */
function getChatworkTargetRoomIds() {
  const sheetName = CONFIG.CHATWORK_SETTINGS_SHEET_NAME;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`エラー: シート「${sheetName}」が見つかりません。`);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) {
    Logger.log(`エラー: シート「${sheetName}」にデータがありません。`);
    return [];
  }

  const header = data[0];
  const headerMap = header.reduce((acc, col, index) => ({ ...acc, [col]: index }), {});

  const requiredColumns = CONFIG.CHATWORK_SETTINGS_HEADERS;
  for (const col of requiredColumns) {
    if (headerMap[col] === undefined) {
      Logger.log(`エラー: シート「${sheetName}」に必要な列「${col}」がありません。`);
      return [];
    }
  }

  const managers = {}; // グループごとのマネージャー情報と週次レポートモード
  const allEmployees = []; // 全社員のリスト

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const group = row[headerMap['グループ名']];
    const name = row[headerMap['氏名']];
    const roomId = row[headerMap['ルームID']] ? row[headerMap['ルームID']].toString() : '';
    const role = row[headerMap['役割']] ? row[headerMap['役割']].toLowerCase() : '';
    const weeklyReportMode = row[headerMap['週次レポートモード']] ? row[headerMap['週次レポートモード']].toLowerCase() : CONFIG.DEFAULT_WEEKLY_REPORT_MODE;

    if (!group || !name || !roomId || !role) {
      Logger.log(`警告: シート「${sheetName}」の${i + 1}行目に不足データがあります。スキップします。`);
      continue; // 空の行や不完全な行はスキップ
    }

    if (role === CONFIG.CHATWORK_ROLE_MANAGER) {
      managers[group] = { name, roomId, weeklyReportMode };
    }
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const group = row[headerMap['グループ名']];
    const name = row[headerMap['氏名']];
    const roomId = row[headerMap['ルームID']] ? row[headerMap['ルームID']].toString() : '';
    const role = row[headerMap['役割']] ? row[headerMap['役割']].toLowerCase() : '';

    if (!group || !name || !roomId || !role) {
      continue; // 空の行や不完全な行はスキップ（上記でログ済み）
    }

    const manager = managers[group];
    if (!manager) {
      Logger.log(`警告: 社員「${name}」のグループ「${group}」に対応するマネージャーが見つかりません。この社員はリストに含まれません。`);
      continue;
    }

    allEmployees.push({
      employeeName: name,
      employeeRoomId: roomId,
      managerName: manager.name,
      managerRoomId: manager.roomId,
      weeklyReportMode: manager.weeklyReportMode,
      group: group,
      role: role
    });
  }

  return allEmployees;
}

/**
 * 全ての部下に日報提出を促す質問をChatworkで送信する。
 */
function sendDailyReportQuestions() {
  const employeeList = getChatworkTargetRoomIds();
  const pendingQuestions = getPendingQuestionMessages();

  employeeList.forEach(employee => {
    const { employeeName, employeeRoomId, role } = employee;

    // マネージャーには日報質問を送信しない
    if (role === CONFIG.CHATWORK_ROLE_MANAGER) {
      Logger.log(`${employeeName}さん（ルームID: ${employeeRoomId}）はマネージャーのため、日報質問送信をスキップしました。`);
      return;
    }

    const existingPending = pendingQuestions.find(q =>
      q.roomId === employeeRoomId &&
      (q.status === CONFIG.STATUS_STRINGS.PENDING || q.status === CONFIG.STATUS_STRINGS.INVALID_FORMAT || q.status.startsWith(CONFIG.STATUS_STRINGS.ERROR_PREFIX))
    );

    if (existingPending) {
      Logger.log(`${employeeName}さん（ルームID: ${employeeRoomId}）には未処理の日報が存在するため、質問送信をスキップしました。ステータス: ${existingPending.status}`);
      return;
    }

    const message = CONFIG.DAILY_REPORT_QUESTION_MESSAGE_TEMPLATE
      .replace('{employeeRoomId}', employeeRoomId)
      .replace('{employeeName}', employeeName);

    try {
      const response = sendChatworkNotification(employeeRoomId, message);
      const messageId = response.message_id.toString();
      logQuestionMessageId(employeeRoomId, messageId, new Date(), CONFIG.STATUS_STRINGS.PENDING);
      Logger.log(`${employeeName}さん (Room ID: ${employeeRoomId}) への質問送信に成功しました。メッセージID: ${messageId}`);
    } catch (e) {
      Logger.log(`${employeeName}さん (Room ID: ${employeeRoomId}) への質問送信に失敗しました: ${e.message}`);
    }
  });
}

/**
 * Chatworkの返信を収集し、日報として解析・評価する。
 */
function processChatworkReplies() {
  const employeeList = getChatworkTargetRoomIds();
  const botAccountId = getBotChatworkAccountId();
  const pendingQuestions = getPendingQuestionMessages();

  employeeList.forEach(employee => {
    const { employeeName, employeeRoomId, managerName, managerRoomId } = employee;

    const questionsForThisRoom = pendingQuestions.filter(q =>
      q.roomId === employeeRoomId && q.status !== CONFIG.STATUS_STRINGS.SUCCESS
    );

    if (questionsForThisRoom.length === 0) {
      return;
    }
    
    const lockedQuestion = questionsForThisRoom.find(q => q.status === CONFIG.STATUS_STRINGS.INVALID_FORMAT || q.status.startsWith(CONFIG.STATUS_STRINGS.ERROR_PREFIX));
    if (lockedQuestion) {
        Logger.log(`${employeeName}さん (Room ID: ${employeeRoomId}) の日報はロックステータス（${lockedQuestion.status}）のためスキップしました。`);
        return;
    }

    try {
      const messages = getChatworkMessages(employeeRoomId, CONFIG.CHATWORK_FETCH_MESSAGE_COUNT);

      for (let i = messages.length - 1; i >= 0; i--) {
        const msg = messages[i];
        if (msg.account_id && msg.account_id.toString() === botAccountId) continue;

        const replyMatch = msg.body.match(CONFIG.CHATWORK_REPLY_REGEX);
        if (!replyMatch) continue;

        const repliedToMessageId = replyMatch[3];
        const matchedQuestion = questionsForThisRoom.find(q => q.messageId === repliedToMessageId);

        if (matchedQuestion) {
          if (msg.body.includes(CONFIG.REPORT_TAG)) {
            const reportData = parseReportFromMessage(employeeName, msg.body);
            const validationResult = validateDailyReport(reportData);

            if (validationResult.isValid) {
              try {
                assessAndNotify(reportData, managerName, managerRoomId, employee.employeeRoomId);
                updateQuestionStatus(employeeRoomId, matchedQuestion.messageId, CONFIG.STATUS_STRINGS.SUCCESS);
              } catch (e) {
                Logger.log(`日報処理中にエラーが発生しました: ${e.message}`);
                updateQuestionStatus(employeeRoomId, matchedQuestion.messageId, CONFIG.STATUS_STRINGS.ERROR_GENERAL, e.message);
              }
            } else {
              Logger.log(`${employeeName}さんの日報はフォーマット不正です。詳細: ${validationResult.message}`);
              updateQuestionStatus(employeeRoomId, matchedQuestion.messageId, CONFIG.STATUS_STRINGS.INVALID_FORMAT, validationResult.message);
            }
          }
          break; // このルームの処理を終了
        }
      }
    } catch (e) {
      Logger.log(`${employeeName}さん (Room ID: ${employeeRoomId}) のメッセージ処理中にエラー: ${e.message}`);
    }
  });
}

/**
 * Chatwork APIを呼び出してメッセージを取得する。
 */
function getChatworkMessages(roomId, count) {
  const apiKey = getChatworkApiKey();
  const url = `${CONFIG.CHATWORK_API_BASE_URL}/rooms/${roomId}/messages?force=1`;
  const options = {
    method: 'get',
    headers: { 'X-ChatWorkToken': apiKey },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  if (responseCode === 200) {
    return JSON.parse(responseText).slice(-count);
  }
  throw new Error(`Chatwork APIからのメッセージ取得に失敗。Status: ${responseCode}, Response: ${responseText}`);
}

/**
 * Chatworkメッセージから日報データを抽出する。
 */
function parseReportFromMessage(name, messageBody) {
  const workContentMatch = messageBody.match(CONFIG.PARSE_REPORT_REGEX.WORK_CONTENT);
  const moodMatch = messageBody.match(CONFIG.PARSE_REPORT_REGEX.MOOD);
  const problemsMatch = messageBody.match(CONFIG.PARSE_REPORT_REGEX.PROBLEMS);
  return {
    date: new Date().toLocaleString('ja-JP'),
    name: name,
    workContent: workContentMatch ? workContentMatch[1].trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    mood: moodMatch ? moodMatch[1].trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    problems: problemsMatch ? problemsMatch[1].trim() : CONFIG.DEFAULT_VALUES.NO_PROBLEM
  };
}

/**
 * 日報データのバリデーションを行う。
 */
function validateDailyReport(reportData) {
  if (!reportData.name || reportData.name.trim() === CONFIG.DEFAULT_VALUES.NOT_APPLICABLE) {
    return { isValid: false, message: '氏名が空です。' };
  }
  if (!reportData.workContent || reportData.workContent.trim() === CONFIG.DEFAULT_VALUES.NOT_APPLICABLE) {
    return { isValid: false, message: '業務内容が空です。' };
  }
  if (!reportData.mood || !CONFIG.MOOD_OPTIONS.includes(reportData.mood.trim())) {
    return { isValid: false, message: `気分に不正な値「${reportData.mood}」が入力されています。` };
  }
  return { isValid: true, message: '' };
}

// --- Gemini API & 通知機能 ---

/**
 * 日報データを評価し、必要に応じて通知する。
 * @param {Object} reportData 日報データ
 * @param {string} managerName マネージャー名
 * @param {string} managerRoomId マネージャーのルームID
 */
function assessAndNotify(reportData, managerName, managerRoomId, employeeRoomId) {
  const geminiPrompt = CONFIG.DAILY_REPORT_ASSESS_PROMPT_TEMPLATE
    .replace('{workContent}', reportData.workContent)
    .replace('{mood}', reportData.mood)
    .replace('{problems}', reportData.problems);

  let geminiStatus = CONFIG.DEFAULT_VALUES.UNKNOWN;
  let geminiReason = CONFIG.DEFAULT_VALUES.NO_RESPONSE;

  try {
    const geminiResponse = callGeminiApi(geminiPrompt);
    const rawGeminiText = geminiResponse.candidates[0].content.parts[0].text;
    const jsonMatch = rawGeminiText.match(/\{([\s\S]*?)\}/);
    if (jsonMatch && jsonMatch[0]) {
      const parsedResponse = JSON.parse(jsonMatch[0]);
      geminiStatus = parsedResponse.status;
      geminiReason = parsedResponse.reason;
    } else {
      throw new Error('Gemini API response did not contain a valid JSON object.');
    }
  } catch (error) {
    Logger.log('Gemini API呼び出しまたは応答解析に失敗: ' + error.message);
  }

  if (geminiStatus === CONFIG.STATUS_STRINGS.DANGER || 
      geminiStatus === CONFIG.STATUS_STRINGS.WARNING ||
      geminiStatus === CONFIG.STATUS_STRINGS.BAD ||
      (geminiStatus === CONFIG.STATUS_STRINGS.NORMAL && reportData.problems !== CONFIG.DEFAULT_VALUES.NO_PROBLEM)) {
    if (!managerRoomId) {
      Logger.log("マネージャーのChatworkルームIDが不明なため、通知をスキップします。");
    } else {
      let subject = CONFIG.DAILY_REPORT_ALERT_SUBJECT_TEMPLATE.replace('{employeeName}', reportData.name);
      if (geminiStatus === CONFIG.STATUS_STRINGS.NORMAL) {
        subject = `【情報】日報から困りごとの報告 - ${reportData.name}`;
      } else if (geminiStatus === CONFIG.STATUS_STRINGS.BAD) {
        subject = `【要確認】日報で「悪い」評価 - ${reportData.name}`;
      }

      const body = CONFIG.DAILY_REPORT_ALERT_BODY_TEMPLATE
        .replace('{subject}', subject)
        .replace('{employeeName}', reportData.name)
        .replace('{date}', reportData.date)
        .replace('{geminiStatus}', geminiStatus)
        .replace('{geminiReason}', geminiReason)
        .replace('{mood}', reportData.mood)
        .replace('{problems}', reportData.problems);
      try {
        sendChatworkNotification(managerRoomId, body);
        Logger.log('Chatworkへの注意通知が正常に送信されました。');
      } catch (e) {
        Logger.log('Chatworkへの注意通知の送信に失敗しました: ' + e.message);
      }
    }
  }
  
  try {
    logReportToSheet(reportData, geminiStatus, geminiReason, managerName, employeeRoomId);
  } catch (e) {
    Logger.log('スプレッドシートへの日報ログ記録に失敗しました: ' + e.message);
  }
}

/**
 * 日報データをスプレッドシートに記録する。
 * @param {Object} reportData 日報データ
 * @param {string} status AIによる評価状態
 * @param {string} reason AIによる評価理由
 * @param {string} managerName マネージャー名
 */
function logReportToSheet(reportData, status, reason, managerName, employeeRoomId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // A. メインの「日報ログ」シートへの書き込み
  const mainLogSheet = spreadsheet.getSheetByName(CONFIG.DAILY_REPORT_LOG_SHEET_NAME);
  if (!mainLogSheet) {
    throw new Error(`シート「${CONFIG.DAILY_REPORT_LOG_SHEET_NAME}」が見つかりません。`);
  }
  mainLogSheet.appendRow([
    new Date(),
    reportData.name,
    managerName || CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    reportData.date,
    reportData.workContent,
    reportData.mood,
    reportData.problems,
    status,
    reason
  ]);
  Logger.log(`日報データを「${CONFIG.DAILY_REPORT_LOG_SHEET_NAME}」シートに記録しました。`);

  // B. メンバーごとの「個別日報ログ」シートへの書き込み
  const memberSheetName = `${reportData.name}_${employeeRoomId}`;
  let memberSheet = spreadsheet.getSheetByName(memberSheetName);

  if (!memberSheet) {
    memberSheet = spreadsheet.insertSheet(memberSheetName);
    memberSheet.appendRow(CONFIG.MEMBER_DAILY_REPORT_LOG_HEADERS);

    // 列幅とヘッダー色の設定
    CONFIG.DAILY_REPORT_LOG_COLUMN_WIDTHS.forEach(col => {
      if (col.index <= CONFIG.MEMBER_DAILY_REPORT_LOG_HEADERS.length) {
        memberSheet.setColumnWidth(col.index, col.width);
      }
    });
    memberSheet.getRange(1, 1, 1, CONFIG.MEMBER_DAILY_REPORT_LOG_HEADERS.length).setBackground(CONFIG.DAILY_REPORT_LOG_HEADER_BG_COLOR);
    memberSheet.setFrozenRows(1); // 1行目を固定
  }

  memberSheet.appendRow([
    new Date(),
    reportData.name,
    managerName || CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    reportData.date,
    reportData.workContent,
    reportData.mood,
    reportData.problems
  ]);
  Logger.log(`日報データを「${memberSheetName}」シートに記録しました。`);
}

/**
 * Chatworkにメッセージを送信する汎用関数。
 */
function sendChatworkNotification(roomId, message, subject = null) {
  try {
    const apiKey = getChatworkApiKey();
    const url = `${CONFIG.CHATWORK_API_BASE_URL}/rooms/${roomId}/messages`;

    // Chatwork記法がエンコードされないように、特殊なエンコード処理を実装します。
    // 1. 全体をエンコードして安全性を確保
    let encodedMessage = encodeURIComponent(message);
    // 2. Chatwork記法で使われる角括弧 [], スラッシュ / のみデコードして戻す
    encodedMessage = encodedMessage.replace(/%5B/g, '[').replace(/%5D/g, ']').replace(/%2F/g, '/');

    let payloadString = `body=${encodedMessage}&self_unread=1`;
    if (subject) {
      payloadString += `&subject=${encodeURIComponent(subject)}`;
    }

    const options = {
      method: 'post',
      headers: { 'X-ChatWorkToken': apiKey },
      contentType: 'application/x-www-form-urlencoded',
      payload: payloadString,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      throw new Error(`Chatwork API送信失敗。Status: ${responseCode}, Response: ${responseText}`);
    }
    return JSON.parse(responseText);
  } catch (e) {
    Logger.log(`sendChatworkNotificationでエラー: ${e.message}`);
    throw e;
  }
}

/**
 * BOT自身のChatworkアカウントIDを取得する。
 */
function getBotChatworkAccountId() {
  const properties = PropertiesService.getScriptProperties();
  let botAccountId = properties.getProperty(CONFIG.CHATWORK_BOT_ACCOUNT_ID_KEY);
  if (botAccountId) return botAccountId;

  try {
    const apiKey = getChatworkApiKey();
    const url = `${CONFIG.CHATWORK_API_BASE_URL}/me`;
    const options = { method: 'get', headers: { 'X-ChatWorkToken': apiKey }, muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Chatwork API (me) 取得失敗: ${response.getContentText()}`);
    }
    const me = JSON.parse(response.getContentText());
    botAccountId = me.account_id.toString();
    properties.setProperty(CONFIG.CHATWORK_BOT_ACCOUNT_ID_KEY, botAccountId);
    return botAccountId;
  } catch (e) {
    Logger.log(`BOTアカウントIDの取得に失敗: ${e.message}`);
    throw e;
  }
}

/**
 * 送信した質問メッセージのIDをBOT質問ログシートに記録する。
 */
function logQuestionMessageId(roomId, messageId, timestamp, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.BOT_QUESTION_LOG_SHEET_NAME);
  if (!sheet) throw new Error(`シート「${CONFIG.BOT_QUESTION_LOG_SHEET_NAME}」が見つかりません。`);
  sheet.appendRow([roomId, messageId, timestamp, status, '']);
  Logger.log(`質問メッセージID ${messageId} をルーム ${roomId} に記録しました。`);
}

/**
 * BOT質問ログシートから未処理の質問メッセージIDを読み込む。
 */
function getPendingQuestionMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.BOT_QUESTION_LOG_SHEET_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  values.shift(); // ヘッダーを除去
  return values.map((row, index) => ({
    rowIndex: index + 2,
    roomId: row[0].toString(),
    messageId: row[1].toString(),
    timestamp: new Date(row[2]),
    status: row[3].toString(),
    errorDetail: row[4] ? row[4].toString() : ''
  }));
}

/**
 * BOT質問ログシートの質問メッセージのステータスを更新する。
 */
function updateQuestionStatus(roomId, messageId, newStatus, errorDetail = '') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.BOT_QUESTION_LOG_SHEET_NAME);
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === roomId && values[i][1].toString() === messageId) {
      sheet.getRange(i + 1, 4).setValue(newStatus);
      sheet.getRange(i + 1, 5).setValue(errorDetail);
      Logger.log(`質問メッセージID ${messageId} のステータスを ${newStatus} に更新しました。`);
      return;
    }
  }
}

/**
 * Gemini APIキーをスクリプトプロパティから取得する。
 */
function getGeminiApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.GEMINI_API_KEY);
  if (!apiKey) throw new Error('Gemini API Key is not set in Script Properties.');
  return apiKey;
}

/**
 * Gemini APIを呼び出し、応答をJSON形式で返す。
 */
function callGeminiApi(prompt) {
  try {
    const apiKey = getGeminiApiKey();
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL_NAME}:generateContent?key=${apiKey}`;
    const requestBody = { contents: [{ parts: [{ text: prompt }] }] };
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    if (responseCode === 200) {
      return JSON.parse(responseText);
    }
    throw new Error(`Gemini API呼び出し失敗。HTTP ${responseCode}, 応答: ${responseText}`);
  } catch (e) {
    throw e;
  }
}

// --- 1on1ヒアリング項目生成機能 ---

/**
 * 指定された社員の過去の日報データを取得し、要約する。
 * @param {string} employeeName 対象社員名
 * @returns {string} 要約された日報ログのテキスト
 */
function getDailyReportDataForEmployee(employeeName, employeeRoomId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DAILY_REPORT_LOG_SHEET_NAME);
  if (!sheet) {
    Logger.log(`シート「${CONFIG.DAILY_REPORT_LOG_SHEET_NAME}」が見つかりません。`);
    return '日報ログが見つかりませんでした。';
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return '過去の日報ログはありません。';
  }
  Logger.log(`「${CONFIG.DAILY_REPORT_LOG_SHEET_NAME}」シートから全 ${values.length - 1} 件のデータを読み込みました。`);

  const header = values[0];
  const dataRows = values.slice(1);

  const headerMap = header.reduce((acc, col, index) => ({ ...acc, [col]: index }), {});

  const nameCol = headerMap['氏名'];
  const dateCol = headerMap['日報日付'];
  const moodCol = headerMap['今日の気分'];
  const problemsCol = headerMap['困っていること'];
  const workContentCol = headerMap['今日の業務内容'];
  const aiStatusCol = headerMap['AI評価状態'];
  const aiReasonCol = headerMap['AI評価理由'];

  if (nameCol === undefined || dateCol === undefined || moodCol === undefined || problemsCol === undefined || workContentCol === undefined || aiStatusCol === undefined || aiReasonCol === undefined) {
    Logger.log('日報ログシートのヘッダーが不正です。');
    return '日報ログシートのヘッダーが不正なため、データを読み込めませんでした。';
  }

  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() - CONFIG.DAILY_REPORT_LOG_FETCH_DAYS_FOR_1ON1);

  const employeeReports = dataRows.filter(row => {
    const reportName = row[nameCol] ? row[nameCol].toString().trim() : '';
    const reportDate = row[dateCol] ? new Date(row[dateCol]) : null;
    return reportName.toLowerCase() === employeeName.toLowerCase() && reportDate && reportDate >= targetDate;
  });
  
  Logger.log(`${employeeName}さんの日報を ${employeeReports.length} 件抽出しました。`);

  if (employeeReports.length === 0) {
    return `過去${CONFIG.DAILY_REPORT_LOG_FETCH_DAYS_FOR_1ON1}日間の日報ログはありません。`;
  }

  // --- 日報ログの集約・要約 (ハイブリッド形式) ---
  const moodCounts = CONFIG.MOOD_OPTIONS.reduce((acc, mood) => ({ ...acc, [mood]: 0 }), {});
  const problemKeywords = CONFIG.PROBLEM_KEYWORDS.reduce((acc, keyword) => ({ ...acc, [keyword]: 0 }), {});
  const positiveKeywords = new Set();
  const representativeReports = [];
  const processedMonths = new Set();

  // AI評価が「危険」「少し悪い」の日報を優先的に収集
  const negativeReports = employeeReports.filter(report => {
    const aiStatus = report[aiStatusCol] ? report[aiStatusCol].toString().trim() : '';
    return aiStatus === CONFIG.STATUS_STRINGS.DANGER || aiStatus === CONFIG.STATUS_STRINGS.WARNING;
  });

  // その他の日報から、月ごとにランダムに選択
  const otherReports = employeeReports.filter(report => {
    const aiStatus = report[aiStatusCol] ? report[aiStatusCol].toString().trim() : '';
    return aiStatus !== CONFIG.STATUS_STRINGS.DANGER && aiStatus !== CONFIG.STATUS_STRINGS.WARNING;
  });

  // 代表的な日報の抜粋を収集 (最大件数はCONFIGで設定)
  const reportsToSample = [...negativeReports];
  
  if (reportsToSample.length < CONFIG.REPRESENTATIVE_REPORTS_COUNT) {
    const shuffledOtherReports = otherReports.sort(() => 0.5 - Math.random());
    for (const report of shuffledOtherReports) {
      const reportDate = new Date(report[dateCol]);
      const monthKey = `${reportDate.getFullYear()}-${reportDate.getMonth()}`;
      
      if (!processedMonths.has(monthKey) && reportsToSample.length < CONFIG.REPRESENTATIVE_REPORTS_COUNT) {
        reportsToSample.push(report);
        processedMonths.add(monthKey);
      }
      if (reportsToSample.length >= CONFIG.REPRESENTATIVE_REPORTS_COUNT) break;
    }
  }

  reportsToSample.slice(0, CONFIG.REPRESENTATIVE_REPORTS_COUNT).forEach(report => {
    const reportDate = report[dateCol] ? new Date(report[dateCol]).toLocaleDateString('ja-JP') : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE;
    const mood = report[moodCol] ? report[moodCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE;
    const problems = report[problemsCol] ? truncateText(report[problemsCol].toString().trim(), CONFIG.TRUNCATE_TEXT_MAX_LENGTH_EXCERPT) : CONFIG.DEFAULT_VALUES.NO_PROBLEM;
    const aiStatus = report[aiStatusCol] ? report[aiStatusCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE;
    representativeReports.push(`日付: ${reportDate}, 気分: ${mood}, 困り事: ${problems}, AI評価: ${aiStatus}`);
  });

  employeeReports.forEach(report => {
    const mood = report[moodCol] ? report[moodCol].toString().trim() : '';
    if (moodCounts.hasOwnProperty(mood)) {
      moodCounts[mood]++;
    }

    const problems = report[problemsCol] ? report[problemsCol].toString().trim() : '';
    for (const keyword in problemKeywords) {
      if (problems.includes(keyword)) {
        problemKeywords[keyword]++;
      }
    }

    const workContent = report[workContentCol] ? report[workContentCol].toString().trim() : '';
    if (mood === '良い' || mood === '非常に良い') {
      if (workContent.includes(CONFIG.POSITIVE_KEYWORDS.RELEASE)) positiveKeywords.add('新規機能リリース');
      if (CONFIG.POSITIVE_KEYWORDS.CUSTOMER_APPRECIATION_KEYWORDS.some(kw => workContent.includes(kw))) positiveKeywords.add('顧客高評価');
      if (workContent.includes(CONFIG.POSITIVE_KEYWORDS.IMPROVEMENT)) positiveKeywords.add('業務改善');
      if (workContent.includes(CONFIG.POSITIVE_KEYWORDS.ACHIEVEMENT)) positiveKeywords.add('目標達成');
    }
  });

  let quantitativeSummary = '気分推移: ';
  quantitativeSummary += Object.entries(moodCounts)
    .filter(([, count]) => count > 0)
    .map(([mood, count]) => `${mood}(${count}回)`) 
    .join(', ');

  const frequentProblemsList = Object.entries(problemKeywords)
    .filter(([, count]) => count > 0)
    .sort((a, b) => b[1] - a[1])
    .map(([keyword, count]) => `${keyword}(${count}回)`) 
    .join(', ');
  if (frequentProblemsList) {
    quantitativeSummary += `\n繰り返し課題: ${frequentProblemsList}`;
  }

  const positiveAspectsList = Array.from(positiveKeywords).join(', ');
  if (positiveAspectsList) {
    quantitativeSummary += `\nポジティブ事項: ${positiveAspectsList}`;
  }

  let excerptsSummary = '';
  if (representativeReports.length > 0) {
    excerptsSummary = `\n\n代表的な日報の抜粋 (最大${CONFIG.REPRESENTATIVE_REPORTS_COUNT}件):\n` + representativeReports.map(r => `- ${r}`).join('\n');
  }

  const finalSummary = `過去${CONFIG.DAILY_REPORT_LOG_FETCH_DAYS_FOR_1ON1}日間の日報サマリー (${employeeReports.length}件のデータ):\n${quantitativeSummary}${excerptsSummary}`;

  Logger.log(`${employeeName}さんの日報ログ要約が完了しました。`);
  return finalSummary;
}

/**
 * Googleドライブから指定された社員の自己評価シートを読み込み、データを抽出する。
 * @param {string} employeeName 対象社員名
 * @param {string} folderId 自己評価シートが保存されているGoogleドライブのフォルダID
 * @param {string} sheetName 自己評価シート内の読み込むシート名
 * @returns {Object|null} 抽出された自己評価データ、またはnull
 */
function getSelfEvaluationDataForEmployee(employeeName, folderId, sheetName) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    let selfEvalFile = null;

    const normalizedEmployeeName = employeeName.trim().toLowerCase();

    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().toLowerCase().includes(normalizedEmployeeName) && file.getName().includes(CONFIG.SELF_EVAL_FILENAME_KEYWORD)) {
        selfEvalFile = file;
        break;
      }
    }

    if (!selfEvalFile) {
      Logger.log(`自己評価シートが見つかりませんでした: ${employeeName}さんの「${CONFIG.SELF_EVAL_FILENAME_KEYWORD}」を含むGoogleスプレッドシートファイルがフォルダID ${folderId} 内に見つかりません。`);
      return null;
    }

    const spreadsheet = SpreadsheetApp.openById(selfEvalFile.getId());
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`エラー: スプレッドシートID「${selfEvalFile.getId()}」内にシート「${sheetName}」が見つかりません。`);
      return null;
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      Logger.log(`シート「${sheetName}」にデータがありません。`);
      return null;
    }
    Logger.log(`自己評価シート「${sheet.getParent().getName()}」の「${sheetName}」から ${values.length - 1} 件のデータを読み込みました。`);

    const header = values[0];
    const dataRows = values.slice(1);

    const headerMap = {};
    header.forEach((col, index) => {
      if (typeof col === 'string') {
        headerMap[col.trim()] = index;
      }
    });

    const extractedData = {
      employeeName: '',
      evaluationPeriod: '',
      questions: [],
      '来季目標': '',
      '目標達成のためにサポートしてほしい事': '',
      'サポート方針': '',
      '目標グレード': ''
    };

    const employeeNameHeaderIndex = headerMap[CONFIG.SELF_EVALUATION_DEFAULT_NAME_HEADER];
    if (employeeNameHeaderIndex === undefined) {
      Logger.log(`エラー: 自己評価シートに「${CONFIG.SELF_EVALUATION_DEFAULT_NAME_HEADER}」列が見つかりません。`);
      return null;
    }

    const employeeRow = dataRows.find(row => {
      const employeeNameInSheet = row[employeeNameHeaderIndex];
      return employeeNameInSheet && employeeNameInSheet.toString().trim().toLowerCase() === normalizedEmployeeName;
    });

    if (!employeeRow) {
      Logger.log(`シート「${sheetName}」に社員「${employeeName}」のデータが見つかりませんでした。`);
      return null;
    }
    Logger.log(`社員「${employeeName}」さんのデータを1件見つけました。`);

    extractedData.employeeName = employeeRow[employeeNameHeaderIndex] ? employeeRow[employeeNameHeaderIndex].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE;

    const evaluationPeriodIndex = headerMap[CONFIG.SELF_EVAL_PERIOD_HEADER];
    if (evaluationPeriodIndex !== undefined) {
      extractedData.evaluationPeriod = employeeRow[evaluationPeriodIndex] ? employeeRow[evaluationPeriodIndex].toString().trim() : '';
    }

    CONFIG.SELF_EVAL_SINGLE_ITEM_HEADERS.forEach(itemHeader => {
      const colIndex = headerMap[itemHeader];
      if (colIndex !== undefined) {
        const value = employeeRow[colIndex];
        extractedData[itemHeader] = value ? value.toString().trim() : '';
      } else {
        extractedData[itemHeader] = '';
      }
    });

    const questionNames = new Set();

    header.forEach(h => {
      if (typeof h === 'string') {
        for (const suffix of CONFIG.SELF_EVAL_QUESTION_SUFFIXES) {
          if (h.endsWith(suffix)) {
            questionNames.add(h.replace(suffix, '').trim());
            break;
          }
        }
      }
    });

    questionNames.forEach(qName => {
      const questionData = {
        questionName: qName,
        questionContent: '',
        selfComment: '',
        managerComment: '',
        selfEvaluation: '',
        managerEvaluation: ''
      };

      CONFIG.SELF_EVAL_QUESTION_SUFFIXES.forEach(suffix => {
        const fullHeader = `${qName}${suffix}`;
        const colIndex = headerMap[fullHeader];
        if (colIndex !== undefined) {
          const value = employeeRow[colIndex];
          if (suffix === '_設問内容') questionData.questionContent = value ? value.toString().trim() : '';
          if (suffix === '_本人コメント') questionData.selfComment = value ? value.toString().trim() : '';
          if (suffix === '_マネージャコメント') questionData.managerComment = value ? value.toString().trim() : '';
          if (suffix === '_自己評価') questionData.selfEvaluation = value ? value.toString().trim() : '';
          if (suffix === '_マネージャ評価') questionData.managerEvaluation = value ? value.toString().trim() : '';
        }
      });
      extractedData.questions.push(questionData);
    });

    Logger.log(`自己評価シート「${selfEvalFile.getName()}」からデータを抽出しました。`);
    return extractedData;

  } catch (e) {
    Logger.log(`自己評価シートの読み込みまたは解析中にエラーが発生しました: ${e.message}`);
    return null;
  }
}

/**
 * テキストを指定した最大長に短縮し、中央を省略する。
 * @param {string} text 対象のテキスト
 * @param {number} maxLength 最大長
 * @returns {string} 短縮されたテキスト
 */
function truncateText(text, maxLength = CONFIG.TRUNCATE_TEXT_MAX_LENGTH) {
  if (text.length <= maxLength) {
    return text;
  }
  const half = Math.floor((maxLength - 5) / 2); // 5は " ... " の分
  return text.substring(0, half) + ' ... ' + text.substring(text.length - half);
}

/**
 * 1on1ヒアリング項目を生成する。
 */
function generate1on1Topics() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const targetEmployeeName = properties.getProperty(CONFIG.TARGET_EMPLOYEE_NAME_FOR_1ON1_KEY);
  const selfEvaluationFolderId = properties.getProperty(CONFIG.SELF_EVALUATION_FOLDER_ID_KEY);
  const selfEvaluationInputSheetName = properties.getProperty(CONFIG.SELF_EVALUATION_INPUT_SHEET_NAME_KEY);

  if (!targetEmployeeName) {
    const errorMessage = `スクリプトプロパティ「${CONFIG.TARGET_EMPLOYEE_NAME_FOR_1ON1_KEY}」が設定されていません。処理をスキップします。`;
    Logger.log(errorMessage);
    ui.alert('設定エラー', errorMessage + '「プロジェクトの設定」 > 「スクリプトプロパティ」を確認してください。', ui.ButtonSet.OK);
    return;
  }

  if (!selfEvaluationFolderId) {
    const errorMessage = `スクリプトプロパティ「${CONFIG.SELF_EVALUATION_FOLDER_ID_KEY}」が設定されていません。処理をスキップします。自己評価シートの読み込みにはこの設定が必要です。`;
    Logger.log(errorMessage);
    ui.alert('設定エラー', errorMessage + '「プロジェクトの設定」 > 「スクリプトプロパティ」を確認してください。', ui.ButtonSet.OK);
    return;
  }

  if (!selfEvaluationInputSheetName) {
    const errorMessage = `スクリプトプロパティ「${CONFIG.SELF_EVALUATION_INPUT_SHEET_NAME_KEY}」が設定されていません。処理をスキップします。自己評価シートの読み込みにはこの設定が必要です。`;
    Logger.log(errorMessage);
    ui.alert('設定エラー', errorMessage + '「プロジェクトの設定」 > 「スクリプトプロパティ」を確認してください。', ui.ButtonSet.OK);
    return;
  }

  const employeeList = getChatworkTargetRoomIds();
  const targetEmployee = employeeList.find(emp => emp.employeeName === targetEmployeeName);

  if (!targetEmployee) {
      const errorMessage = `「${CONFIG.CHATWORK_SETTINGS_SHEET_NAME}」シートに「${targetEmployeeName}」さんが見つかりません。`;
      Logger.log(errorMessage);
      ui.alert('設定エラー', errorMessage, ui.ButtonSet.OK);
      return;
  }
  
  const { managerRoomId } = targetEmployee;

  if (!managerRoomId) {
      Logger.log(`マネージャーのChatworkルームIDが設定されていません（${targetEmployeeName}さん）。`);
      ui.alert('エラー', 'マネージャーのルームIDが設定されていません。', ui.ButtonSet.OK);
      return;
  }

  Logger.log(`TARGET_EMPLOYEE_NAME_FOR_1ON1で指定された${targetEmployeeName}さんの1on1ヒアリング項目生成を開始します。`);

  // 日報ログの要約を取得
  const dailyReportSummary = getDailyReportDataForEmployee(targetEmployeeName, targetEmployee.employeeRoomId);

  // 自己評価シートのデータを取得
  const selfEvalData = getSelfEvaluationDataForEmployee(targetEmployeeName, selfEvaluationFolderId, selfEvaluationInputSheetName);

  if (!selfEvalData || Object.keys(selfEvalData).length === 0) {
    const infoMessage = `「${targetEmployeeName}」さんの自己評価シートから有効なデータが抽出できませんでした。シート名「${selfEvaluationInputSheetName}」またはシートのフォーマットを確認してください。`;
    Logger.log(infoMessage);
    ui.alert('情報', infoMessage, ui.ButtonSet.OK);
    return;
  }

  // Geminiプロンプトの構築
  let promptSelfEvalData = '';

  if (selfEvalData.evaluationPeriod) {
    promptSelfEvalData += `評価期間: ${selfEvalData.evaluationPeriod}\n`;
  }

  selfEvalData.questions.forEach((q, index) => {
    promptSelfEvalData += `${index + 1}. ${q.questionContent}`;
    if (q.selfComment) promptSelfEvalData += ` (本人コメント): ${q.selfComment}`;
    if (q.selfEvaluation) promptSelfEvalData += ` / 自己評価: ${q.selfEvaluation}`;
    promptSelfEvalData += `\n`;
  });

  if (selfEvalData['来季目標']) promptSelfEvalData += `来季目標: ${selfEvalData['来季目標']}\n`;
  if (selfEvalData['目標達成のためにサポートしてほしい事']) promptSelfEvalData += `目標達成のためにサポートしてほしい事: ${selfEvalData['目標達成のためにサポートしてほしい事']}\n`;
  if (selfEvalData['サポート方針']) promptSelfEvalData += `サポート方針: ${selfEvalData['サポート方針']}\n`;
  if (selfEvalData['目標グレード']) promptSelfEvalData += `目標グレード: ${selfEvalData['目標グレード']}\n`;

  const geminiPrompt = CONFIG.ONE_ON_ONE_PROMPT_TEMPLATE
    .replace(/{anonymousName}/g, targetEmployeeName)
    .replace('{dailyReportSummary}', dailyReportSummary)
    .replace('{promptSelfEvalData}', promptSelfEvalData);

  try {
    const geminiResponse = callGeminiApi(geminiPrompt);
    const hearingTopicsRaw = geminiResponse.candidates[0].content.parts[0].text;
    Logger.log('--- Geminiからの生応答 ---\n' + hearingTopicsRaw);

    // --- Chatwork通知メッセージの整形ロジック（修正版） ---
    let formattedTopics = '';
    const topics = hearingTopicsRaw.split(/\n\s*\n/); // 質問・根拠のペアで分割

    let questionCount = 0;
    topics.forEach(topic => {
      const questionMatch = topic.match(/\*\d+\.\s*([\s\S]*?)(?=\n\*具体的な質問と根拠\*)/);
      const rationaleMatch = topic.match(/\*具体的な質問と根拠\*([\s\S]*)/);

      if (questionMatch && questionMatch[1] && rationaleMatch && rationaleMatch[1]) {
        questionCount++;
        const question = questionMatch[1].replace(/\*/g, '').trim();
        const rationale = rationaleMatch[1].replace(/\*/g, '').trim();

        formattedTopics += `${questionCount}. ${question}\n`;
        formattedTopics += `   根拠: ${rationale}\n\n`;
      }
    });

    if (questionCount === 0) {
      Logger.log('Gemini応答の解析に失敗しました。正規表現が期待通りにマッチしませんでした。フォールバックとして生テキストを送信します。');
      formattedTopics = hearingTopicsRaw.replace(/\*/g, ''); // Markdown記号を除去
    }
    // --- 整形ロジックここまで ---

    const subject = CONFIG.ONE_ON_ONE_SUBJECT_TEMPLATE.replace('{employeeName}', targetEmployeeName);
    const body = CONFIG.ONE_ON_ONE_BODY_TEMPLATE
      .replace('{subject}', subject)
      .replace('{formattedTopics}', formattedTopics.trim());

    sendChatworkNotification(managerRoomId, body);
    Logger.log(`${targetEmployeeName}さんの1on1ヒアリング項目をChatworkに通知しました。`);
    ui.alert('完了', `${targetEmployeeName}さんの1on1ヒアリング項目生成と通知が完了しました。`, ui.ButtonSet.OK);
  } catch (error) {
    const errorMessage = `1on1ヒアリング項目生成中にエラーが発生しました: ${error.message}`;
    Logger.log(errorMessage);
    ui.alert('エラー', errorMessage, ui.ButtonSet.OK);
  }
}

// --- 週次レポート生成機能 ---

/**
 * 週次レポートを生成し、マネージャーに通知する。
 * 毎週月曜日の朝に実行されることを想定。
 */
function generateWeeklyReports() {
  const ui = SpreadsheetApp.getUi();
  Logger.log('週次レポートの生成を開始します。');

  try {
    const employeeList = getChatworkTargetRoomIds();
    const managersToProcess = {}; // Key: managerRoomId, Value: { managerName, managerRoomId, weeklyReportMode, groupName, employees: [] }

    // マネージャーごとに社員をグループ化
    employeeList.forEach(employee => {
      const { employeeName, employeeRoomId, managerName, managerRoomId, weeklyReportMode, group, role } = employee;

      if (!managerRoomId) {
        Logger.log(`社員「${employeeName}」のマネージャーのChatworkルームIDが不明なため、処理をスキップします。`);
        return;
      }

      if (!managersToProcess[managerRoomId]) {
        managersToProcess[managerRoomId] = {
          name: managerName,
          roomId: managerRoomId,
          weeklyReportMode: weeklyReportMode,
          groupName: group,
          employees: []
        };
      }
      managersToProcess[managerRoomId].employees.push({ employeeName, employeeRoomId, role });
    });

    // 各マネージャー/グループに対してレポートを生成
    for (const managerRoomId in managersToProcess) {
      const managerData = managersToProcess[managerRoomId];
      const today = new Date();
      const subjectDate = `${today.getFullYear()}/${today.getMonth() + 1}/${today.getDate()}`;

      if (managerData.weeklyReportMode === CONFIG.WEEKLY_REPORT_MODES.INDIVIDUAL) {
        // Individual report mode (週次メンバーコンディションレポート)
        Logger.log(`マネージャー「${managerData.name}」さん向けに個別週次レポートを生成します。`);
        const allIndividualReportsForManager = [];

        managerData.employees.forEach(employee => {
          if (employee.role === CONFIG.CHATWORK_ROLE_EMPLOYEE) {
            const rawReports = getWeeklyRawReports(employee.employeeName, employee.employeeRoomId);
            if (rawReports.length > 0) {
              try {
                const { reportText, subject } = generateIndividualReportWithGemini(rawReports, employee.employeeName);
                allIndividualReportsForManager.push({ reportText, subject });
              } catch (e) {
                Logger.log(`「${employee.employeeName}」さんの個別レポート生成中にエラーが発生しました: ${e.message}`);
              }
            } else {
              Logger.log(`社員「${employee.employeeName}」の過去1週間の日報データが見つかりませんでした。`);
            }
          }
        });

        if (allIndividualReportsForManager.length > 0) {
          let combinedReportBody = '';
          allIndividualReportsForManager.forEach((report, index) => {
            combinedReportBody += report.reportText;
            if (index < allIndividualReportsForManager.length - 1) {
              combinedReportBody += '\n[hr]\n';
            }
          });

          const subject = `【週次メンバーコンディションレポート】${managerData.name}さん向け - ${subjectDate}週`;
          const finalCombinedReportBody = `${subject}\n[hr]\n\n▼ 各メンバーのコンディションサマリー\n\n${combinedReportBody}\n\n詳細については、日報ログスプレッドシートをご確認ください。`;
          sendChatworkNotification(managerData.roomId, finalCombinedReportBody);
          Logger.log(`マネージャー「${managerData.name}」さんへ週次メンバーコンディションレポートを通知しました。`);
        }

      } else if (managerData.weeklyReportMode === CONFIG.WEEKLY_REPORT_MODES.TEAM) {
        // Team summary mode (週次チームコンディションサマリー)
        Logger.log(`グループ「${managerData.groupName}」向けに週次チームコンディションサマリーを生成します。`);
        const teamEmployees = managerData.employees.filter(emp => emp.role === CONFIG.CHATWORK_ROLE_EMPLOYEE);
        if (teamEmployees.length > 0) {
          const summaryData = getWeeklyTeamReportSummaryForGroup(teamEmployees.map(e => e.employeeName));
          if (summaryData) {
            const teamSummaryReport = generateTeamSummaryReportWithGemini(summaryData, managerData.groupName);
            const subject = CONFIG.TEAM_SUMMARY_SUBJECT_TEMPLATE
              .replace('{groupName}', managerData.groupName)
              .replace('{date}', subjectDate);
            const body = CONFIG.TEAM_SUMMARY_BODY_TEMPLATE
              .replace('{subject}', subject)
              .replace('{reportBody}', teamSummaryReport);
            sendChatworkNotification(managerData.roomId, body);
            Logger.log(`マネージャー「${managerData.name}」さんへグループ「${managerData.groupName}」の週次チームコンディションサマリーを通知しました。`);
          } else {
            Logger.log(`グループ「${managerData.groupName}」の過去1週間の日報データが見つかりませんでした。`);
          }
        }
      } else {
        Logger.log(`グループ「${managerData.groupName}」にレポート対象の社員が見つかりませんでした。`);
      }
    }

    Logger.log('週次レポートの生成と通知が正常に完了しました。');
    // 手動実行時のみアラート
    if (typeof ScriptApp === 'undefined' || !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'generateWeeklyReports')) {
       ui.alert('完了', '週次レポートの生成と通知が完了しました。', ui.ButtonSet.OK);
    }

  } catch (e) {
    Logger.log(`週次レポート生成中にエラーが発生しました: ${e.message}`);
    // 手動実行時のみアラート
     if (typeof ScriptApp === 'undefined' || !ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'generateWeeklyReports')) {
       ui.alert('エラー', `処理中にエラーが発生しました: ${e.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 指定された社員の過去1週間分の日報ログの生データを取得する。
 * @param {string} employeeName 対象社員名
 * @returns {Array<Object>} 過去1週間分の日報データ（オブジェクトの配列）
 */
function getWeeklyRawReports(employeeName, employeeRoomId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DAILY_REPORT_LOG_SHEET_NAME);
  if (!sheet) {
    Logger.log(`シート「${CONFIG.DAILY_REPORT_LOG_SHEET_NAME}」が見つかりません。`);
    return [];
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  const header = values[0];
  const dataRows = values.slice(1);
  const headerMap = header.reduce((acc, col, index) => ({ ...acc, [col]: index }), {});

  const nameCol = headerMap['氏名'];
  const dateCol = headerMap['タイムスタンプ']; // Use timestamp for filtering
  const moodCol = headerMap['今日の気分'];
  const problemsCol = headerMap['困っていること'];
  const workContentCol = headerMap['今日の業務内容'];
  const aiStatusCol = headerMap['AI評価状態'];
  const aiReasonCol = headerMap['AI評価理由'];

  if (nameCol === undefined || dateCol === undefined || moodCol === undefined || problemsCol === undefined || workContentCol === undefined || aiStatusCol === undefined || aiReasonCol === undefined) {
    Logger.log('日報ログシートのヘッダーが不正です。必要な列が見つかりません。');
    return [];
  }

  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - CONFIG.WEEKLY_REPORT_FETCH_DAYS);

  const employeeWeeklyReports = dataRows.filter(row => {
    const reportName = row[nameCol] ? row[nameCol].toString().trim() : '';
    const reportDate = row[dateCol] ? new Date(row[dateCol]) : null;
    return reportName.toLowerCase() === employeeName.toLowerCase() && reportDate && reportDate >= oneWeekAgo;
  }).map(row => ({
    date: row[dateCol] ? new Date(row[dateCol]).toLocaleDateString('ja-JP') : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    workContent: row[workContentCol] ? row[workContentCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    mood: row[moodCol] ? row[moodCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    problems: row[problemsCol] ? row[problemsCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    aiStatus: row[aiStatusCol] ? row[aiStatusCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE,
    aiReason: row[aiReasonCol] ? row[aiReasonCol].toString().trim() : CONFIG.DEFAULT_VALUES.NOT_APPLICABLE
  }));

  return employeeWeeklyReports;
}

/**
 * 指定された社員リストの過去1週間分の日報ログを集計・分析する
 * @param {Array<string>} employeeNames 対象社員名の配列
 * @returns {Object|null} 集計データ
 */
function getWeeklyTeamReportSummaryForGroup(employees) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DAILY_REPORT_LOG_SHEET_NAME);
  if (!sheet) {
    Logger.log(`シート「${CONFIG.DAILY_REPORT_LOG_SHEET_NAME}」が見つかりません。`);
    return null;
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    Logger.log('日報ログがありません。');
    return null;
  }

  const header = values[0];
  const dataRows = values.slice(1);
  const headerMap = header.reduce((acc, col, index) => ({ ...acc, [col]: index }), {});

  const nameCol = headerMap['氏名'];
  const dateCol = headerMap['タイムスタンプ']; // Use timestamp for filtering
  const moodCol = headerMap['今日の気分'];
  const problemsCol = headerMap['困っていること'];
  const workContentCol = headerMap['今日の業務内容'];

  if (nameCol === undefined || dateCol === undefined || moodCol === undefined || problemsCol === undefined || workContentCol === undefined) {
    Logger.log('日報ログシートのヘッダーが不正です。必要な列が見つかりません。');
    return null;
  }

  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - CONFIG.WEEKLY_REPORT_FETCH_DAYS);

  const employeeNames = employees.map(e => e.employeeName);
  const weeklyReports = dataRows.filter(row => {
    const reportName = row[nameCol] ? row[nameCol].toString().trim() : '';
    const reportDate = row[dateCol] ? new Date(row[dateCol]) : null;
    return employeeNames.includes(reportName) && reportDate && reportDate >= oneWeekAgo;
  });

  if (weeklyReports.length === 0) {
    Logger.log('過去1週間の日報データがありません。');
    return null;
  }

  const moodCounts = {};
  const problemKeywords = {};
  const positiveKeywords = new Set();

  weeklyReports.forEach(report => {
    // 気分集計
    const mood = report[moodCol] ? report[moodCol].toString().trim() : CONFIG.DEFAULT_VALUES.UNKNOWN;
    moodCounts[mood] = (moodCounts[mood] || 0) + 1;

    // 課題キーワード集計
    const problems = report[problemsCol] ? report[problemsCol].toString().trim() : '';
    if (problems && problems !== CONFIG.DEFAULT_VALUES.NO_PROBLEM) {
        for (const keyword of CONFIG.PROBLEM_KEYWORDS) {
            if (problems.includes(keyword)) {
                problemKeywords[keyword] = (problemKeywords[keyword] || 0) + 1;
            }
        }
    }

    // ポジティブな動き
    const workContent = report[workContentCol] ? report[workContentCol].toString().trim() : '';
    const moodForPositive = report[moodCol] ? report[moodCol].toString().trim() : '';
    if (moodForPositive === '良い' || moodForPositive === '非常に良い') {
      if (workContent.includes('完了') || workContent.includes('達成')) positiveKeywords.add('タスク完了/目標達成');
      if (workContent.includes('感謝') || workContent.includes('助かった')) positiveKeywords.add('チームワーク/協力');
      if (workContent.includes('改善') || workContent.includes('効率化')) positiveKeywords.add('業務改善/効率化');
    }
  });

  const commonProblems = Object.entries(problemKeywords)
    .filter(([, count]) => count > CONFIG.COMMON_PROBLEMS_MIN_COUNT)
    .sort((a, b) => b[1] - a[1])
    .slice(0, CONFIG.COMMON_PROBLEMS_MAX_DISPLAY)
    .map(([keyword]) => keyword);

  return {
    totalReports: weeklyReports.length,
    moodCounts: moodCounts,
    commonProblems: commonProblems,
    positiveTrends: Array.from(positiveKeywords)
  };
}

/**
 * 個別の日報データからGeminiプロンプトを生成し、APIを呼び出してコンディションレポートを取得する
 * @param {Array<Object>} rawReports 対象メンバーの過去1週間分の日報データ
 * @param {string} employeeName 対象メンバー名
 * @returns {{reportText: string, subject: string}} Geminiが生成したコンディションレポートと件名
 */
function generateIndividualReportWithGemini(rawReports, employeeName) {
  const reportEntries = rawReports.map(r => 
    `日付: ${r.date}, 業務内容: ${r.workContent}, 気分: ${r.mood}, 困り事: ${r.problems}, AI評価: ${r.aiStatus || CONFIG.DEFAULT_VALUES.NOT_APPLICABLE}`
  ).join('\n');

  const individualReportPrompt = CONFIG.WEEKLY_REPORT_PROMPT_TEMPLATE
    .replace(/{reportType}/g, CONFIG.ANONYMOUS_EMPLOYEE_NAME_SHORT)
    .replace(/{subjectName}/g, employeeName)
    .replace(/{dataSummary}/g, reportEntries)
    .replace(/{specificEmphasis}/g, '')
    .replace(/{reportBody}/g, '[ここにレポート本文を記述。物語形式で、メンバーの「空気感」を伝えるように記述してください。以下の要素を盛り込むことを推奨します。]\n- 全体的な傾向や雰囲気\n- 特に注目すべきポジティブな点（具体的な業務内容や成果に触れる）\n- 潜在的な課題や懸念点（背景や影響にも触れる）\n- 変化の兆候や、今後注目すべき点\n\n[レポート本文の具体的な記述例]\n今週の[社員名/チーム名]は、[全体的な傾向、例：概ね良好なコンディションを維持しました]。[ポジティブな点、例：特に〇〇プロジェクトでの新規機能リリースに貢献し、高い達成感が見られました]。一方で、[懸念点、例：週後半にはテスト環境の不安定さに起因する軽微なストレスも報告されており、注意が必要です]。全体として、[今後の展望や注目点、例：来週の進捗に期待しつつ、〇〇の状況を注視していきます]。');

  try {
    Logger.log(`Gemini APIに${employeeName}さんの個別レポート生成をリクエストします。`);
    const geminiResponse = callGeminiApi(individualReportPrompt);
    const reportText = geminiResponse.candidates[0].content.parts[0].text;
    Logger.log(`Gemini APIから${employeeName}さんの個別レポートを取得しました。`);

    const today = new Date();
    const subjectDate = `${today.getFullYear()}/${today.getMonth() + 1}/${today.getDate()}`;
    const subject = CONFIG.INDIVIDUAL_REPORT_SUBJECT_TEMPLATE
      .replace('{anonymousName}', employeeName)
      .replace('{date}', subjectDate);

    return { reportText, subject }; // Return an object
  } catch (e) {
    Logger.log(`${employeeName}さんの個別レポート生成中にGemini API呼び出しでエラーが発生しました: ${e.message}`);
    throw new Error('Gemini APIからの個別レポート生成に失敗しました。');
  }
}

/**
 * 集計データからGeminiプロンプトを生成し、APIを呼び出してサマリーレポートを取得する
 * @param {Object} summaryData 
 * @param {string} groupName
 * @returns {string} Geminiが生成したサマリーレポート
 */
function generateTeamSummaryReportWithGemini(summaryData, groupName) {
  const moodSummary = Object.entries(summaryData.moodCounts).map(([mood, count]) => `${mood}(${count}回)`).join(', ');
  const problemsSummary = summaryData.commonProblems.length > 0 ? summaryData.commonProblems.join(', ') : '特筆すべき共通課題は報告されていません。';
  const positivesSummary = summaryData.positiveTrends.length > 0 ? summaryData.positiveTrends.join(', ') : '特筆すべきポジティブな動きは報告されていません。';

  const dataSummary = `総日報数: ${summaryData.totalReports}件\n気分報告の傾向: ${moodSummary}\n共通の課題: ${problemsSummary}\nポジティブな動き: ${positivesSummary}`;

  const teamSummaryPrompt = CONFIG.WEEKLY_REPORT_PROMPT_TEMPLATE
    .replace(/{reportType}/g, 'チーム')
    .replace(/{subjectName}/g, groupName)
    .replace(/{dataSummary}/g, dataSummary)

  try {
    Logger.log('Gemini APIにサマリー生成をリクエストします。');
    const geminiResponse = callGeminiApi(teamSummaryPrompt);
    const reportText = geminiResponse.candidates[0].content.parts[0].text;
    Logger.log('Gemini APIからサマリーレポートを取得しました。');
    return reportText;
  } catch (e) {
    Logger.log(`Gemini API呼び出し中にエラーが発生しました: ${e.message}`);
    throw new Error('Gemini APIからのレポート生成に失敗しました。');
  }
}