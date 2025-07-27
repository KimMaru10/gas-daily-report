function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📌 日報ツール")
    .addItem("✏️ 誤字を修正", "reviseDailyReportText")
    .addItem("📧 日報を送信", "sendTodaysDailyReport")
    .addToUi();
}

function getOpenAIApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
}

// ChatGPTに誤字修正処理の関数
function correctTextWithChatGPT(text) {
  const OPENAI_API_KEY = getOpenAIApiKey();

  if (!OPENAI_API_KEY) {
    Logger.log("エラー: OPENAI_API_KEY がスクリプトプロパティに設定されていません。");
    SpreadsheetApp.getUi().alert("エラー: OpenAI APIキーが設定されていません。開発者に連絡してください。");
    return text;
  }

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-3.5-turbo", // 使用するモデル (gpt-4o, gpt-4 なども可能ですが、料金が変わります)
    messages: [
      {
        role: "system",
        content: "以下の文章の誤字脱字、言い回しの改善をお願いします。内容は変えずに自然な日本語にしてください。"
      },
      {
        role: "user",
        content: text
      }
    ],
    max_tokens: 1000,
    temperature: 0.7
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + OPENAI_API_KEY // 認証ヘッダー
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      const json = JSON.parse(responseText);
      Logger.log(`ChatGPT API成功レスポンス（JSON）: ${JSON.stringify(json, null, 2)}`);

      const result = json?.choices?.[0]?.message?.content;
      
      if (!result) {
        Logger.log(`ChatGPT APIは成功したが、修正テキストが結果に含まれていません。APIレスポンス: ${responseText}`);
        SpreadsheetApp.getUi().alert("修正処理は完了しましたが、ChatGPTからの修正結果が空でした。元のテキストを維持します。");
      }
      return result || text; 
    } else {
      Logger.log(`ChatGPT APIエラー - ステータスコード: ${responseCode}`);
      Logger.log(`ChatGPT APIエラー - レスポンス本文: ${responseText}`);
      SpreadsheetApp.getUi().alert(`ChatGPT APIからの応答エラーです。ステータスコード: ${responseCode}。詳細についてはログを確認してください。`);
      return text;
    }
  } catch (e) {
    Logger.log(`API呼び出し中に予期せぬ例外が発生しました: ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`ChatGPT API呼び出し中に予期せぬエラーが発生しました。詳細についてはログを確認してください。`);
    return text;
  }
}

/**
 * 指定した文字数で文字列に改行を挿入する関数。
 * @param {string} text 元のテキスト
 * @param {number} lineLength 1行の最大文字数
 * @returns {string} 改行が挿入されたテキスト
 */
function insertNewlines(text, lineLength) {
  if (!text || text.length <= lineLength) {
    return text;
  }

  let result = '';
  let currentLine = '';
  const lines = text.split('\n'); // 既存の改行で分割

  lines.forEach(line => {
    // 各行をさらに指定の文字数で分割
    for (let i = 0; i < line.length; i++) {
      currentLine += line[i];
      if (currentLine.length >= lineLength && line[i] !== ' ') { 
        result += currentLine + '\n';
        currentLine = '';
      }
    }
    if (currentLine.length > 0) {
      result += currentLine; 
      currentLine = '';
    }
    result += '\n'; 
  });

  return result.trim(); 
}




function reviseDailyReportText() {
  Logger.log("--- reviseDailyReportText 関数開始 ---");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("日報");
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  Logger.log("getDataRange().getValues() 開始");
  const data = sheet.getDataRange().getValues();
  Logger.log("getDataRange().getValues() 終了");
  const rows = data.slice(1); // ヘッダー除く

  const dateCol = 0;
  const reflectionCol = 4;
  const fiveActCol = 5;

  let foundRowIndex = -1;
  let todaysRow;

  Logger.log("forEach ループ開始");
  rows.forEach((row, index) => {
    const rowDate = Utilities.formatDate(new Date(row[dateCol]), "Asia/Tokyo", "yyyy/MM/dd");
    if (rowDate === today && foundRowIndex === -1) {
      todaysRow = row;
      foundRowIndex = index + 1; // header行を除くため+1
    }
  });
  Logger.log("forEach ループ終了");

  if (!todaysRow) {
    SpreadsheetApp.getUi().alert("今日の日報が見つかりません。");
    Logger.log("今日の日報が見つかりませんでした。");
    return;
  }

  Logger.log(`今日の日報の行インデックス（0から始まる配列内）: ${foundRowIndex}`);
  Logger.log(`スプレッドシートの書き込み対象行番号（1から始まる）: ${foundRowIndex + 1}`);

  Logger.log("ChatGPT API呼び出し開始 (感想)");
  let revisedReflection = correctTextWithChatGPT(todaysRow[reflectionCol]); // let に変更
  Logger.log("ChatGPT API呼び出し終了 (感想)");

  Logger.log("ChatGPT API呼び出し開始 (行動5者)");
  let revisedFiveAct = correctTextWithChatGPT(todaysRow[fiveActCol]);     // let に変更
  Logger.log("ChatGPT API呼び出し終了 (行動5者)");

  // --- ここから改行処理を追加 ---
  const MAX_LINE_LENGTH = 26; // 1行の最大文字数
  revisedReflection = insertNewlines(revisedReflection, MAX_LINE_LENGTH);
  revisedFiveAct = insertNewlines(revisedFiveAct, MAX_LINE_LENGTH);
  // --- 改行処理の追加ここまで ---

  // シートに書き戻す
  const targetRangeReflection = sheet.getRange(foundRowIndex + 1, reflectionCol + 1);
  const targetRangeFiveAct = sheet.getRange(foundRowIndex + 1, fiveActCol + 1);

  Logger.log(`感想を書き込むセル: ${targetRangeReflection.getA1Notation()} に "${revisedReflection}"`);
  Logger.log(`行動５者を書き込むセル: ${targetRangeFiveAct.getA1Notation()} に "${revisedFiveAct}"`);

  Logger.log("setValue 開始");
  targetRangeReflection.setValue(revisedReflection);
  targetRangeFiveAct.setValue(revisedFiveAct);
  Logger.log("setValue 終了");

  SpreadsheetApp.getUi().alert("感想と行動５者を修正しました。");
  Logger.log("--- reviseDailyReportText 関数終了 ---");
}

// 毎日9時に本日の日報欄を作る関数
function insertDailyReportTemplate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("日報");
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");

  // テンプレート行（名前以外は空欄でOK）
  const newRow = [
    formattedDate,         // 日付
    "",                    // 名前
    "",                    // TO
    "",                    // CC
    "",                    // 今日の感想
    "",                    // 行動5者
    "",                    // 業務内容
    "",                    // 明日の予定
    ""                     // TODO
  ];

  sheet.appendRow(newRow);
}

//日報を送る関数
function sendTodaysDailyReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("日報");
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // ヘッダー除く
  const headers = data[0];

  const dateCol = 0; // A列（0番目）
  const sentCol = 9; // J列（9番目）

  let foundRowIndex = -1;
  let todaysRow;

  rows.forEach((row, index) => {
    const rowDate = Utilities.formatDate(new Date(row[dateCol]), "Asia/Tokyo", "yyyy/MM/dd");
    const sentFlag = row[sentCol];
    if (rowDate === today && !sentFlag && foundRowIndex === -1) {
      todaysRow = row;
      foundRowIndex = index + 1; // headerを除いているので+1
    }
  });

  if (foundRowIndex === -1 || !todaysRow) {
    SpreadsheetApp.getUi().alert("今日の日報が未入力、またはすでに送信されています。");
    return;
  }

  const [date, name, to, cc, reflection, fiveAct, summary, schedule, todo] = todaysRow;
  const formattedDate = formatDateWithJapaneseWeekday(date);
  const subject = `【業務報告書】${formattedDate}　${name}`;

  const body = `
関係者の皆様

　お疲れ様です。
　２５新卒の${name}です。
　
　本日の日報を提出いたします。
　ご確認よろしくお願いします。

────────────────────────────────────
●今日の感想や気づいた点
────────────────────────────────────
${reflection}

●今日の行動５者
────────────────────────────────────
${fiveAct}

●今日の業務内容のまとめ
────────────────────────────────────
${summary}

●明日のスケジュール／ＴＯＤＯ
────────────────────────────────────
スケジュール：
${schedule}

ＴＯＤＯ：
${todo}

────────────────────────────────────
`;

  GmailApp.sendEmail(to, subject, body, { cc: cc });

  // 「送信済み」とマーク
  sheet.getRange(foundRowIndex + 1, sentCol + 1).setValue("送信済み");

  SpreadsheetApp.getUi().alert("日報メールを送信しました！");
}

//曜日を漢字に変換する関数
function formatDateWithJapaneseWeekday(date) {
  const youbiList = ['日', '月', '火', '水', '木', '金', '土'];
  const d = new Date(date);
  const year = d.getFullYear();
  const month = ("0" + (d.getMonth() + 1)).slice(-2);
  const day = ("0" + d.getDate()).slice(-2);
  const youbi = youbiList[d.getDay()];
  return `${year}年　${month}月${day}日（${youbi}）`;
}