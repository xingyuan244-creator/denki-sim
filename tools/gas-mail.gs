/**
 * かんたん電気代シミュレーション ─ フォーム受信 → メール転送 ＋ Slack通知
 * =====================================================================
 * 【設置手順】
 *  1. 通知を受け取りたいGoogleアカウントでログインした状態で https://script.google.com を開く
 *  2. 「新しいプロジェクト」→ このファイルの内容をすべて貼り付けて保存
 *  3. 右上「デプロイ」→「新しいデプロイ」→ 種類の選択で「ウェブアプリ」
 *       説明          : 電気代シミュレーション フォーム
 *       次のユーザーとして実行 : 自分（ログイン中のアカウント）
 *       アクセスできるユーザー : 全員
 *  4. 「デプロイ」→ 初回は権限の承認を求められるので許可する
 *       （「このアプリは確認されていません」と出たら「詳細」→「安全ではないページに移動」）
 *  5. 表示された「ウェブアプリのURL」をコピー
 *  6. denki-sim.html の SUBMIT_ENDPOINT に貼り付ける
 *       var SUBMIT_ENDPOINT = "https://script.google.com/macros/s/........./exec";
 *
 * 【すでに稼働中のプロジェクトを更新する場合】
 *  ・コードを貼り替えたら、必ず「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
 *    バージョン「新バージョン」→「デプロイ」で再デプロイしてください。
 *    保存しただけでは公開URLに反映されません（URLは変わりません）。
 *  ・NOTIFY_TO を空のまま貼り替えると通知が届かなくなります。必ず記入してください。
 *
 * 【Slack通知を使う場合】
 *  1. https://api.slack.com/apps →「Create New App」→「From scratch」
 *  2. アプリ名（例：電気代シミュレーション通知）とワークスペースを選ぶ
 *  3. 左メニュー「Incoming Webhooks」→ スイッチを「On」
 *  4. 「Add New Webhook to Workspace」→ 投稿先チャンネルを選び「許可する」
 *  5. 発行された Webhook URL を下の SLACK_WEBHOOK_URL に貼る → 保存 → 再デプロイ
 *  ・空のままなら Slack通知は行われず、メールのみ送信されます
 *  ・ワークスペースの設定によっては、アプリ追加に管理者の承認が必要です
 *
 * 【スプレッドシートにも記録したい場合】
 *  ・新規スプレッドシートを作り、URL の /d/ と /edit の間のIDを SHEET_ID に設定
 *
 * 【注意】
 *  ・Gmail の1日あたりの送信上限は無料アカウントで100通程度です
 * =====================================================================
 */

var NOTIFY_TO         = "";   // ← 通知先メールアドレス（公開リポジトリのため実アドレスは記載していません）
var SLACK_WEBHOOK_URL = "";   // ← Slack通知が不要なら空のまま
var SHEET_ID          = "";   // ← スプレッドシート記録が不要なら空のまま

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ---- メール送信（主たる通知。ここは必ず実行する） ----
    sendMail_(data);

    // ---- Slack通知（失敗してもメールの成否には影響させない） ----
    if (SLACK_WEBHOOK_URL) {
      try { postToSlack_(data); }
      catch (slackErr) { console.error("Slack通知に失敗: " + slackErr); }
    }

    // ---- スプレッドシート記録（任意） ----
    if (SHEET_ID) {
      try { appendToSheet_(data); }
      catch (sheetErr) { console.error("シート記録に失敗: " + sheetErr); }
    }

    return json_({ ok: true });
  } catch (err) {
    MailApp.sendEmail(NOTIFY_TO, "【電気代シミュレーション】受信エラー",
      String(err) + "\n\n" + (e && e.postData ? e.postData.contents : ""));
    return json_({ ok: false, error: String(err) });
  }
}

/** ブラウザで直接URLを開いたときの死活確認用 */
function doGet() {
  return json_({ ok: true, message: "endpoint is alive" });
}

/* =========================================================================
   メール
   ========================================================================= */
function sendMail_(data) {
  var lines = [];
  Object.keys(data).forEach(function (k) {
    if (k.indexOf("―") === 0) { lines.push("", k, ""); }   // 区切り行
    else { lines.push(k + "： " + data[k]); }
  });

  var siteName = data["サイト名"] || "電気代シミュレーション";
  var company  = data["会社名／屋号"] || "（未入力）";
  var options  = {
    to: NOTIFY_TO,
    subject: "【" + siteName + "】お問い合わせ：" + company,
    body: lines.join("\n")
  };
  var replyTo = data["メールアドレス"];
  if (replyTo && replyTo.indexOf("@") > 0) { options.replyTo = replyTo; }
  MailApp.sendEmail(options);
}

/* =========================================================================
   Slack
   -------------------------------------------------------------------------
   チャンネルでは一覧性が命なので、本文をそのまま流さず
   「誰から・どのくらいの規模か・どこから来たか」が一目で分かる形に整える。
   ========================================================================= */
function postToSlack_(data) {
  var company = data["会社名／屋号"] || "（未入力）";
  var blocks  = [];

  blocks.push({
    type: "header",
    text: { type: "plain_text", text: "⚡ 新しいお問い合わせ：" + trim_(company, 120), emoji: true }
  });

  // 返信に必要な情報を先頭にまとめる
  var head = [];
  pushField_(head, "メールアドレス", data["メールアドレス"]);
  pushField_(head, "ご担当者名",     data["ご担当者名"] || data["お名前"]);
  pushField_(head, "電話番号",       data["電話番号"]);
  pushField_(head, "業種",           data["業種"]);
  if (head.length) { blocks.push({ type: "section", fields: head }); }

  // 金額まわり（シミュレーションからの送信のみ）
  var money = [];
  pushField_(money, "直近1ヶ月の電気代", data["直近1ヶ月の電気代"]);
  pushField_(money, "想定 年間削減額",   data["想定 年間削減額"]);
  pushField_(money, "見直しの優先度",     data["見直しの優先度"]);
  pushField_(money, "契約数",            data["契約数"]);
  if (money.length) { blocks.push({ type: "divider" }, { type: "section", fields: money }); }

  // お問い合わせフォームからの本文
  if (data["本文"]) {
    blocks.push({ type: "divider" });
    blocks.push({
      type: "section",
      text: { type: "mrkdwn", text: "*お問い合わせ内容*\n>>> " + trim_(String(data["本文"]), 2800) }
    });
  }

  // 補足情報
  var ctx = [];
  if (data["エリア"])   { ctx.push("エリア: " + data["エリア"]); }
  if (data["事業形態"]) { ctx.push(data["事業形態"]); }
  if (data["主な設備"]) { ctx.push("設備: " + data["主な設備"]); }
  if (data["流入経路"]) { ctx.push("流入経路: *" + data["流入経路"] + "*"); }
  if (data["送信日時"]) { ctx.push(data["送信日時"]); }
  if (ctx.length) {
    blocks.push({ type: "context", elements: [{ type: "mrkdwn", text: trim_(ctx.join("  ｜  "), 2800) }] });
  }

  var payload = {
    text: "新しいお問い合わせ：" + trim_(company, 120),   // 通知バナーに出るプレーンテキスト
    blocks: blocks
  };

  var res = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error("Slack応答 " + res.getResponseCode() + ": " + res.getContentText());
  }
}

/** Slackのfieldsは1ブロック10件までなので、超える分は積まない */
function pushField_(arr, label, value) {
  if (!value || arr.length >= 10) { return; }
  arr.push({ type: "mrkdwn", text: "*" + label + "*\n" + trim_(String(value), 1800) });
}

function trim_(s, max) {
  return s.length > max ? s.slice(0, max - 1) + "…" : s;
}

/* =========================================================================
   スプレッドシート
   ========================================================================= */
function appendToSheet_(data) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  var keys = Object.keys(data).filter(function (k) { return k.indexOf("―") !== 0; });
  if (sh.getLastRow() === 0) { sh.appendRow(keys); }
  sh.appendRow(keys.map(function (k) { return data[k]; }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================================
   動作テスト
   -------------------------------------------------------------------------
   エディタ上でこの関数を実行すると、テストのメールとSlack通知が届きます。
   ========================================================================= */
function testSend() {
  doPost({ postData: { contents: JSON.stringify({
    "サイト名": "かんたん電気代シミュレーション",
    "送信日時": new Date().toLocaleString("ja-JP"),
    "流入経路": "test",
    "メールアドレス": "test@example.com",
    "会社名／屋号": "テスト商店",
    "ご担当者名": "山田 太郎",
    "電話番号": "000-0000-0000",
    "――― 入力内容 ―――": "",
    "事業形態": "法人",
    "業種": "飲食店・カフェ",
    "エリア": "関西",
    "契約数": "1件",
    "主な設備": "冷蔵・冷凍ショーケース",
    "現在の契約先": "大手電力",
    "直近1ヶ月の電気代": "120,000 円",
    "想定 月間削減額": "20,000 円",
    "想定 年間削減額": "240,000 円",
    "見直しの優先度": "高（いま見直す価値が大きい）"
  }) } });
}

/** Slackの疎通だけを確認したいとき */
function testSlackOnly() {
  if (!SLACK_WEBHOOK_URL) { throw new Error("SLACK_WEBHOOK_URL が未設定です"); }
  postToSlack_({
    "サイト名": "かんたん電気代シミュレーション",
    "会社名／屋号": "接続テスト",
    "メールアドレス": "test@example.com",
    "流入経路": "test",
    "送信日時": new Date().toLocaleString("ja-JP")
  });
}
