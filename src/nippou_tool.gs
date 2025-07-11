const CONFIG = {
  // シート名
  BOT_QUESTION_LOG_SHEET_NAME: 'BOT質問ログ',
  DAILY_REPORT_LOG_SHEET_NAME: '日報ログ',
  CHATWORK_SETTINGS_SHEET_NAME: 'Chatwork設定',

  // 日報データ取得範囲
  DAILY_REPORT_LOG_FETCH_DAYS: 365,

  // truncateTextの最大長
  TRUNCATE_TEXT_MAX_LENGTH: 100,

  // validateDailyReportの気分選択肢
  MOOD_OPTIONS: ['良い', '普通', '少し悪い', '悪い'],

  // getDailyReportDataForEmployee内の問題キーワード
  PROBLEM_KEYWORDS: ['疲労', '残業', '人間関係', '遅延', 'プレッシャー', 'モチベーション', 'コミュニケーション', 'スキル', '不明点', '認識齟齬'],

  // getDailyReportDataForEmployee内の代表日報抜粋数
  REPRESENTATIVE_REPORTS_COUNT: 5,

  // Chatwork設定シートの役割名
  CHATWORK_ROLE_MANAGER: 'manager',
  CHATWORK_ROLE_EMPLOYEE: 'employee',

  // Chatwork設定シートのヘッダー名
  CHATWORK_SETTINGS_HEADERS: ['グループ名', '氏名', 'ルームID', '役割'],

  // 日報ログシートのヘッダー名
  DAILY_REPORT_LOG_HEADERS: ['タイムスタンプ', '氏名', 'マネージャー', '日報日付', '今日の業務内容', '今日の気分', '困っていること', 'AI評価状態', 'AI評価理由'],

  // BOTアカウントIDのスクリプトプロパティキー
  CHATWORK_BOT_ACCOUNT_ID_KEY: 'CHATWORK_BOT_ACCOUNT_ID',

  // 1on1対象社員名のスクリプトプロパティキー
  TARGET_EMPLOYEE_NAME_FOR_1ON1_KEY: 'TARGET_EMPLOYEE_NAME_FOR_1ON1',

  // 自己評価フォルダIDのスクリプトプロパティキー
  SELF_EVALUATION_FOLDER_ID_KEY: 'SELF_EVALUATION_FOLDER_ID',

  // 自己評価入力シート名のスクリプトプロパティキー
  SELF_EVALUATION_INPUT_SHEET_NAME_KEY: 'SELF_EVALUATION_INPUT_SHEET_NAME',

  // 自己評価シートの氏名列ヘッダー名（デフォルト）
  SELF_EVALUATION_DEFAULT_NAME_HEADER: '氏名',

  // Gemini APIモデル名
  GEMINI_MODEL_NAME: 'gemini-2.0-flash',

  // 定期実行のデフォルト時刻
  SCHEDULE_DEFAULTS: {
    questionHour: 9,
    questionMinute: 0,
    replyHour: 18,
    replyMinute: 0
  }
};

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する関数。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報AIツール')
    .addItem('日報質問送信（Chatwork）', 'sendDailyReportQuestions')
    .addItem('Chatwork日報取得・分析', 'processChatworkReplies')
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
  const defaults = CONFIG.SCHEDULE_DEFAULTS;

  const getPropertyAsInt = (key, defaultValue) => {
    const prop = properties.getProperty(key);
    const value = parseInt(prop, 10);
    if (isNaN(value)) {
      Logger.log(`スクリプトプロパティ「${key}」が未設定または不正です。デフォルト値 (${defaultValue}) を使用します。`);
      return defaultValue;
    }
    return value;
  };

  const questionHour = getPropertyAsInt('DAILY_QUESTION_TIME_HOUR', defaults.questionHour);
  const questionMinute = getPropertyAsInt('DAILY_QUESTION_TIME_MINUTE', defaults.questionMinute);
  const replyHour = getPropertyAsInt('DAILY_REPLY_COLLECT_TIME_HOUR', defaults.replyHour);
  const replyMinute = getPropertyAsInt('DAILY_REPLY_COLLECT_TIME_MINUTE', defaults.replyMinute);

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

  ScriptApp.newTrigger('cleanUpBotQuestionLog')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();
  Logger.log(`BOT質問ログクリーンアップトリガーを毎週月曜日深夜2時に設定しました。`);

  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました。 質問送信: 毎日' + `${questionHour}時${questionMinute}分頃` + ' 返信収集: 毎日' + `${replyHour}時${replyMinute}分頃` + ' ログクリーンアップ: 毎週月曜日深夜');
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
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[3].toString();
    const timestamp = new Date(row[2]);

    if (status === '返信済み_処理成功' || ((status === '未返信' || status === '返信済み_フォーマット不正' || status.startsWith('エラー発生')) && timestamp < sevenDaysAgo)) {
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
  const apiKey = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  if (!apiKey) {
    throw new Error('Chatwork API Key is not set in Script Properties.');
  }
  return apiKey;
}

/**
 * 「Chatwork設定」シートから全社員の情報をフラットな配列として取得する。
 * @returns {Array<{employeeName: string, employeeRoomId: string, managerName: string, managerRoomId: string}>}
 */
function getChatworkTargetRoomIds() {
  const sheetName = CONFIG.CHATWORK_SETTINGS_SHEET_NAME;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`エラー: シート「${sheetName}」が見つかりません。`);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const header = data.shift(); // ヘッダー行を除去
  const headerMap = header.reduce((acc, col, index) => ({ ...acc, [col]: index }), {});

  const requiredColumns = CONFIG.CHATWORK_SETTINGS_HEADERS;
  for (const col of requiredColumns) {
    if (headerMap[col] === undefined) {
      Logger.log(`エラー: シート「${sheetName}」に必要な列「${col}」がありません。`);
      return [];
    }
  }

  const managers = {};
  const employees = [];

  data.forEach(row => {
    const group = row[headerMap[CONFIG.CHATWORK_SETTINGS_HEADERS[0]]];
    const name = row[headerMap[CONFIG.CHATWORK_SETTINGS_HEADERS[1]]];
    const roomId = row[headerMap[CONFIG.CHATWORK_SETTINGS_HEADERS[2]]].toString();
    const role = row[headerMap[CONFIG.CHATWORK_SETTINGS_HEADERS[3]]];

    if (!group || !name || !roomId || !role) return; // 空の行はスキップ

    if (role.toLowerCase() === CONFIG.CHATWORK_ROLE_MANAGER) {
      managers[group] = { name, roomId };
    } else if (role.toLowerCase() === CONFIG.CHATWORK_ROLE_EMPLOYEE) {
      employees.push({ group, name, roomId });
    }
  });

  const flatEmployeeList = employees.map(emp => {
    const manager = managers[emp.group];
    if (!manager) {
      Logger.log(`警告: 社員「${emp.name}」のグループ「${emp.group}」に対応するマネージャーが見つかりません。`);
      return null;
    }
    return {
      employeeName: emp.name,
      employeeRoomId: emp.roomId,
      managerName: manager.name,
      managerRoomId: manager.roomId
    };
  }).filter(Boolean); // nullを除外

  return flatEmployeeList;
}

/**
 * 全ての部下に日報提出を促す質問をChatworkで送信する。
 */
function sendDailyReportQuestions() {
  const employeeList = getChatworkTargetRoomIds();
  const pendingQuestions = getPendingQuestionMessages();

  employeeList.forEach(employee => {
    const { employeeName, employeeRoomId } = employee;

    const existingPending = pendingQuestions.find(q =>
      q.roomId === employeeRoomId &&
      (q.status === '未返信' || q.status === '返信済み_フォーマット不正' || q.status.startsWith('エラー発生'))
    );

    if (existingPending) {
      Logger.log(`${employeeName}さん（ルームID: ${employeeRoomId}）には未処理の日報が存在するため、質問送信をスキップしました。ステータス: ${existingPending.status}`);
      return;
    }

    const message = `[To:${employeeRoomId}] ${employeeName}さん\nおはようございます！\n本日の日報を以下のフォーマットでご返信ください。\n\n#日報\n業務内容：\n気分：(良い/普通/少し悪い/悪い)\n困っていること：`;
    try {
      const response = sendChatworkNotification(employeeRoomId, message);
      const messageId = response.message_id.toString();
      logQuestionMessageId(employeeRoomId, messageId, new Date(), '未返信');
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
      q.roomId === employeeRoomId && q.status !== '返信済み_処理成功'
    );

    if (questionsForThisRoom.length === 0) {
      return;
    }
    
    const lockedQuestion = questionsForThisRoom.find(q => q.status === '返信済み_フォーマット不正' || q.status.startsWith('エラー発生'));
    if (lockedQuestion) {
        Logger.log(`${employeeName}さん (Room ID: ${employeeRoomId}) の日報はロックステータス（${lockedQuestion.status}）のためスキップしました。`);
        return;
    }

    try {
      const messages = getChatworkMessages(employeeRoomId, 50);
      let processedThisRoom = false;

      for (let i = messages.length - 1; i >= 0; i--) {
        const msg = messages[i];
        if (msg.account_id && msg.account_id.toString() === botAccountId) continue;

        const replyMatch = msg.body.match(/[[]rp aid=(\d+) to=(\d+)-(\d+)[]]/);
        if (!replyMatch) continue;

        const repliedToMessageId = replyMatch[3];
        const matchedQuestion = questionsForThisRoom.find(q => q.messageId === repliedToMessageId);

        if (matchedQuestion) {
          if (msg.body.includes('#日報')) {
            const reportData = parseReportFromMessage(employeeName, msg.body);
            const validationResult = validateDailyReport(reportData);

            if (validationResult.isValid) {
              try {
                assessAndNotify(reportData, managerName, managerRoomId);
                updateQuestionStatus(employeeRoomId, matchedQuestion.messageId, '返信済み_処理成功');
              } catch (e) {
                Logger.log(`日報処理中にエラーが発生しました: ${e.message}`);
                updateQuestionStatus(employeeRoomId, matchedQuestion.messageId, 'エラー発生_その他', e.message);
              }
            } else {
              Logger.log(`${employeeName}さんの日報はフォーマット不正です。詳細: ${validationResult.message}`);
              updateQuestionStatus(employeeRoomId, matchedQuestion.messageId, '返信済み_フォーマット不正', validationResult.message);
            }
          }
          processedThisRoom = true;
          break;
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
  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages?force=1`;
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
  const workContentMatch = messageBody.match(/業務内容：\s*([\s\S]*?)(?=\n*気分：|\n*困っていること：|$)/);
  const moodMatch = messageBody.match(/気分：\s*([\s\S]*?)(?=\n*困っていること：|$)/);
  const problemsMatch = messageBody.match(/困っていること：\s*([\s\S]*)/);
  return {
    date: new Date().toLocaleString('ja-JP'),
    name: name,
    workContent: workContentMatch ? workContentMatch[1].trim() : 'N/A',
    mood: moodMatch ? moodMatch[1].trim() : 'N/A',
    problems: problemsMatch ? problemsMatch[1].trim() : '特になし'
  };
}

/**
 * 日報データのバリデーションを行う。
 */
function validateDailyReport(reportData) {
  const moods = CONFIG.MOOD_OPTIONS;
  if (!reportData.name || reportData.name.trim() === 'N/A') {
    return { isValid: false, message: '氏名が空です。' };
  }
  if (!reportData.workContent || reportData.workContent.trim() === 'N/A') {
    return { isValid: false, message: '業務内容が空です。' };
  }
  if (!reportData.mood || !moods.includes(reportData.mood.trim())) {
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
function assessAndNotify(reportData, managerName, managerRoomId) {
  const geminiPrompt = `以下の日報の内容を分析し、提出者の現在の心理状態や業務の調子について、5段階（非常に良い、良い、普通、少し悪い、危険）で評価してください。\n**特に「今日の気分」が悪い場合や、「困っていること」にネガティブな兆候が見られる場合は「危険」と判断し、その理由も簡潔に述べてください。**\n氏名は匿名化し、「提出者」として言及してください。\n\n業務内容：${reportData.workContent}\n気分：${reportData.mood}\n困っていること：${reportData.problems}\n\n結果はJSON形式で返してください。例: { "status": "危険", "reason": "具体例：今日の気分が悪いと申告しており、困っている内容にXXとあるため。" }`;
  let geminiStatus = '不明';
  let geminiReason = 'Gemini APIからの応答がありませんでした。';

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

  if (geminiStatus === '危険' || geminiStatus === '少し悪い') {
    if (!managerRoomId) {
      Logger.log("マネージャーのChatworkルームIDが不明なため、通知をスキップします。");
    } else {
      const subject = `【注意】日報から社員の調子に懸念 - ${reportData.name}`;
      const body = `[info][title]${subject}[/title]提出者：${reportData.name}\n日付：${reportData.date}\nGemini AIによる評価：${geminiStatus}\n理由：${geminiReason}\n[hr]▼ 日報抜粋\n今日の気分：${reportData.mood}\n困っていること：${reportData.problems}\n[hr]詳細については、スプレッドシートをご確認ください。[/info]`;
      try {
        sendChatworkNotification(managerRoomId, body);
        Logger.log('Chatworkへの注意通知が正常に送信されました。');
      } catch (e) {
        Logger.log('Chatworkへの注意通知の送信に失敗しました: ' + e.message);
      }
    }
  }
  
  try {
    logReportToSheet(reportData, geminiStatus, geminiReason, managerName);
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
function logReportToSheet(reportData, status, reason, managerName) {
  const logSheetName = CONFIG.DAILY_REPORT_LOG_SHEET_NAME;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(logSheetName);
  if (!sheet) {
    throw new Error(`シート「${logSheetName}」が見つかりません。`);
  }
  sheet.appendRow([
    new Date(),           // タイムスタンプ
    managerName || 'N/A', // マネージャー名
    reportData.name,      // 氏名
    reportData.date,      // 日報日付
    reportData.workContent,
    reportData.mood,
    reportData.problems,
    status,
    reason
  ]);
  Logger.log(`日報データを「${logSheetName}」シートに記録しました。`);
}

/**
 * Chatworkにメッセージを送信する汎用関数。
 */
function sendChatworkNotification(roomId, message) {
  try {
    const apiKey = getChatworkApiKey();
    const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;

    // Chatwork記法がエンコードされないように、特殊なエンコード処理を実装します。
    // 1. 全体をエンコードして安全性を確保
    let encodedMessage = encodeURIComponent(message);
    // 2. Chatwork記法で使われる角括弧 [], スラッシュ / のみデコードして戻す
    encodedMessage = encodedMessage.replace(/%5B/g, '[').replace(/%5D/g, ']').replace(/%2F/g, '/');

    const payloadString = `body=${encodedMessage}&self_unread=1`;

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
    const url = 'https://api.chatwork.com/v2/me';
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
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
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
function getDailyReportDataForEmployee(employeeName) {
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
  targetDate.setDate(targetDate.getDate() - CONFIG.DAILY_REPORT_LOG_FETCH_DAYS);

  const employeeReports = dataRows.filter(row => {
    const reportName = row[nameCol] ? row[nameCol].toString().trim() : '';
    const reportDate = row[dateCol] ? new Date(row[dateCol]) : null;
    return reportName.toLowerCase() === employeeName.toLowerCase() && reportDate && reportDate >= targetDate;
  });
  
  Logger.log(`${employeeName}さんの日報を ${employeeReports.length} 件抽出しました。`);

  if (employeeReports.length === 0) {
    return `過去${CONFIG.DAILY_REPORT_LOG_FETCH_DAYS}日間の日報ログはありません。`;
  }

  // --- 日報ログの集約・要約 (ハイブリッド形式) ---
  const moodCounts = { '非常に良い': 0, '良い': 0, '普通': 0, '少し悪い': 0, '悪い': 0 };
  const problemKeywords = CONFIG.PROBLEM_KEYWORDS.reduce((acc, keyword) => ({ ...acc, [keyword]: 0 }), {});
  const positiveKeywords = new Set();
  const representativeReports = [];
  const processedMonths = new Set();

  // AI評価が「危険」「少し悪い」の日報を優先的に収集
  const negativeReports = employeeReports.filter(report => {
    const aiStatus = report[aiStatusCol] ? report[aiStatusCol].toString().trim() : '';
    return aiStatus === '危険' || aiStatus === '少し悪い';
  });

  // その他の日報から、月ごとにランダムに選択
  const otherReports = employeeReports.filter(report => {
    const aiStatus = report[aiStatusCol] ? report[aiStatusCol].toString().trim() : '';
    return aiStatus !== '危険' && aiStatus !== '少し悪い';
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
    const reportDate = report[dateCol] ? new Date(report[dateCol]).toLocaleDateString('ja-JP') : 'N/A';
    const mood = report[moodCol] ? report[moodCol].toString().trim() : 'N/A';
    const problems = report[problemsCol] ? truncateText(report[problemsCol].toString().trim(), 50) : '特になし';
    const aiStatus = report[aiStatusCol] ? report[aiStatusCol].toString().trim() : 'N/A';
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
      if (workContent.includes('リリース')) positiveKeywords.add('新規機能リリース');
      if (workContent.includes('顧客') && (workContent.includes('高評価') || workContent.includes('感謝'))) positiveKeywords.add('顧客高評価');
      if (workContent.includes('改善')) positiveKeywords.add('業務改善');
      if (workContent.includes('達成')) positiveKeywords.add('目標達成');
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

  const finalSummary = `過去${CONFIG.DAILY_REPORT_LOG_FETCH_DAYS}日間の日報サマリー (${employeeReports.length}件のデータ):\n${quantitativeSummary}${excerptsSummary}`;

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
      if (file.getName().toLowerCase().includes(normalizedEmployeeName) && file.getName().includes("自己評価")) {
        selfEvalFile = file;
        break;
      }
    }

    if (!selfEvalFile) {
      Logger.log(`自己評価シートが見つかりませんでした: ${employeeName}さんの「自己評価」を含むGoogleスプレッドシートファイルがフォルダID ${folderId} 内に見つかりません。`);
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

    extractedData.employeeName = employeeRow[employeeNameHeaderIndex] ? employeeRow[employeeNameHeaderIndex].toString().trim() : 'N/A';

    const evaluationPeriodIndex = headerMap['評価期間'];
    if (evaluationPeriodIndex !== undefined) {
      extractedData.evaluationPeriod = employeeRow[evaluationPeriodIndex] ? employeeRow[evaluationPeriodIndex].toString().trim() : '';
    }

    const SINGLE_ITEM_HEADERS = ['来季目標', '目標達成のためにサポートしてほしい事', 'サポート方針', '目標グレード'];
    SINGLE_ITEM_HEADERS.forEach(itemHeader => {
      const colIndex = headerMap[itemHeader];
      if (colIndex !== undefined) {
        const value = employeeRow[colIndex];
        extractedData[itemHeader] = value ? value.toString().trim() : '';
      } else {
        extractedData[itemHeader] = '';
      }
    });

    const questionSuffixes = ['_設問内容', '_本人コメント', '_マネージャコメント', '_自己評価', '_マネージャ評価'];
    const questionNames = new Set();

    header.forEach(h => {
      if (typeof h === 'string') {
        for (const suffix of questionSuffixes) {
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

      questionSuffixes.forEach(suffix => {
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
  const dailyReportSummary = getDailyReportDataForEmployee(targetEmployeeName);

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

  const anonymousName = `対象者`;

  const geminiPrompt = `以下の${anonymousName}さんの過去1年間の日報サマリーと自己評価シートのデータを総合的に分析し、次回の1on1面談でマネージャーが${anonymousName}さんにヒアリングすべき具体的な質問やテーマを5つ提案してください。質問は部下の心情に寄り添い、具体的な行動を促す形式にしてください。\n\n**過去1年間の日報サマリー：**\n${dailyReportSummary}\n\n**自己評価シートデータ：**\n${promptSelfEvalData}\n\n---\n\n**【重要】回答フォーマットの厳守**\n以下のフォーマットに厳密に従って、5つのヒアリング項目を提案してください。各項目は、質問と根拠を明確に分けて記述してください。\n\n*1. [ここに1つ目の質問内容を記述]\n*具体的な質問と根拠*[ここに1つ目の質問の根拠を記述]\n\n*2. [ここに2つ目の質問内容を記述]\n*具体的な質問と根拠*[ここに2つ目の質問の根拠を記述]\n\n*3. [ここに3つ目の質問内容を記述]\n*具体的な質問と根拠*[ここに3つ目の質問の根拠を記述]\n\n*4. [ここに4つ目の質問内容を記述]\n*具体的な質問と根拠*[ここに4つ目の質問の根拠を記述]\n\n*5. [ここに5つ目の質問内容を記述]\n*具体的な質問と根拠*[ここに5つ目の質問の根拠を記述]`;

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

    const subject = `【1on1ヒアリング項目提案】${targetEmployeeName}さん向け`;
    const body = `${subject}\n[hr]\n\n▼ 提案されたヒアリング項目\n\n${formattedTopics.trim()}\n\n[hr]\nこの提案は、日報ログと自己評価シートの分析に基づいています。詳細なデータはスプレッドシートをご確認ください`;

    sendChatworkNotification(managerRoomId, body);
    Logger.log(`${targetEmployeeName}さんの1on1ヒアリング項目をChatworkに通知しました。`);
    ui.alert('完了', `${targetEmployeeName}さんの1on1ヒアリング項目生成と通知が完了しました。`, ui.ButtonSet.OK);
  } catch (error) {
    const errorMessage = `1on1ヒアリング項目生成中にエラーが発生しました: ${error.message}`;
    Logger.log(errorMessage);
    ui.alert('エラー', errorMessage, ui.ButtonSet.OK);
  }
}