/**
 * email_classifier.gs
 * 「要処理」ラベルの未読メールをClaude APIで分類・要約し、
 * スプレッドシートへの記録とSlack通知を行う。
 */

// シート名・ラベル名・API設定の定数
var LABEL_TODO      = '要処理';
var LABEL_DONE      = '処理済み';
var SHEET_MAIL_LOG  = 'メールログ';
var SHEET_ERROR_LOG = 'エラーログ';
var CLAUDE_MODEL    = 'claude-haiku-4-5-20251001';
var CLAUDE_API_URL  = 'https://api.anthropic.com/v1/messages';

// ────────────────────────────────────────────
// メイン処理
// ────────────────────────────────────────────
function classifyEmails() {
  var props         = PropertiesService.getScriptProperties();
  var claudeApiKey  = props.getProperty('CLAUDE_API_KEY');
  var slackWebhook  = props.getProperty('SLACK_WEBHOOK_URL');
  var ss            = SpreadsheetApp.getActiveSpreadsheet();

  // Gmail ラベルを取得（「処理済み」は存在しなければ作成）
  var todoLabel = GmailApp.getUserLabelByName(LABEL_TODO);
  var doneLabel = GmailApp.getUserLabelByName(LABEL_DONE) || GmailApp.createLabel(LABEL_DONE);

  if (!todoLabel) {
    logError(ss, '「要処理」ラベルが存在しません。Gmailで作成してください。');
    return;
  }

  var threads = todoLabel.getThreads();

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(message) {
      // 未読メールのみ処理
      if (!message.isUnread()) return;

      var subject = message.getSubject();

      try {
        var date   = message.getDate();
        var sender = message.getFrom();
        // 本文が長い場合はAPIコスト削減のため先頭2000文字に絞る
        var body   = message.getPlainBody().substring(0, 2000);

        // ── 分類・要約 ────────────────────────────
        var result = classifyWithClaude(claudeApiKey, subject, body);

        // ── スプレッドシートに記録 ────────────────
        logToMailSheet(ss, date, sender, subject, result.category, result.summary);

        // ── Slack 通知 ────────────────────────────
        notifySlack(slackWebhook, subject, result.category, result.summary);

        // ── ラベル付け替えと既読化 ────────────────
        thread.addLabel(doneLabel);
        thread.removeLabel(todoLabel);
        message.markRead();

      } catch (e) {
        logError(ss, 'メール処理中エラー（件名: ' + subject + '）: ' + e.message);
      }
    });
  });
}

// ────────────────────────────────────────────
// Claude API でメールを分類・要約する
// ────────────────────────────────────────────
function classifyWithClaude(apiKey, subject, body) {
  var prompt = [
    '以下のメールを分類し、要約してください。',
    '',
    '【分類カテゴリ】次の4つから最も適切なものを1つ選んでください：',
    '- クレーム：苦情・不満・クレームに関する内容',
    '- 質問：問い合わせ・確認・質問に関する内容',
    '- 注文：商品・サービスの注文・申込みに関する内容',
    '- その他：上記に当てはまらない内容',
    '',
    '【出力形式】必ずJSON形式のみで返してください。前後に説明文を加えないこと。',
    '{"category": "クレーム|質問|注文|その他", "summary": "50文字以内の要約"}',
    '',
    '【件名】' + subject,
    '【本文】' + body
  ].join('\n');

  var response = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 256,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var statusCode = response.getResponseCode();
  if (statusCode !== 200) {
    throw new Error('Claude API エラー (HTTP ' + statusCode + '): ' + response.getContentText());
  }

  var responseJson = JSON.parse(response.getContentText());
  var text = responseJson.content[0].text.trim();

  // レスポンスにマークダウンコードブロックが含まれる場合を考慮して抽出
  var match = text.match(/\{[\s\S]*\}/);
  if (!match) {
    throw new Error('Claude APIレスポンスからJSONを抽出できませんでした: ' + text);
  }

  return JSON.parse(match[0]);
}

// ────────────────────────────────────────────
// 「メールログ」シートに記録する
// ────────────────────────────────────────────
function logToMailSheet(ss, date, sender, subject, category, summary) {
  var sheet = ss.getSheetByName(SHEET_MAIL_LOG);

  // シートが存在しなければ作成してヘッダーを追加
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MAIL_LOG);
    sheet.getRange(1, 1, 1, 5).setValues([['受信日時', '送信者', '件名', '分類', '要約']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  sheet.appendRow([date, sender, subject, category, summary]);
}

// ────────────────────────────────────────────
// Slack Incoming Webhook で通知する
// ────────────────────────────────────────────
function notifySlack(webhookUrl, subject, category, summary) {
  // カテゴリ別に絵文字を付けて視認性を高める
  var emoji = { 'クレーム': ':rotating_light:', '質問': ':question:', '注文': ':shopping_trolley:', 'その他': ':mailbox:' };
  var icon  = emoji[category] || ':mailbox:';

  var text = [
    icon + ' *新着メール通知*',
    '件名：' + subject,
    '分類：' + category,
    '要約：' + summary
  ].join('\n');

  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  });
}

// ────────────────────────────────────────────
// エラーを「エラーログ」シートに記録する
// ────────────────────────────────────────────
function logError(ss, message) {
  var sheet = ss.getSheetByName(SHEET_ERROR_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ERROR_LOG);
    sheet.getRange(1, 1, 1, 2).setValues([['発生日時', 'エラー内容']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  sheet.appendRow([new Date(), message]);
  Logger.log('ERROR: ' + message);
}

// ────────────────────────────────────────────
// 5分おきの自動トリガーを登録する
// ※ GASエディタからこの関数を一度だけ手動実行してください。
//   以降は5分おきに classifyEmails() が自動で起動します。
// ────────────────────────────────────────────
function setFiveMinuteTrigger() {
  // 重複登録を防ぐため既存の同名トリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'classifyEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 5分ごとに実行するトリガーを作成
  ScriptApp.newTrigger('classifyEmails')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('5分おきのトリガーを登録しました。');
}
