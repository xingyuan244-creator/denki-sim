/**
 * かんたん電気代シミュレーション ─ フォーム受信 → メール転送
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
 * 【スプレッドシートにも記録したい場合】
 *  ・新規スプレッドシートを作り、URL の /d/ と /edit の間のIDを SHEET_ID に設定
 *  ・空欄のままなら、メール送信のみ行われます
 *
 * 【注意】
 *  ・コードを修正したら、必ず「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
 *    バージョン「新バージョン」→「デプロイ」で再デプロイしてください（URLは変わりません）
 *  ・Gmail の1日あたりの送信上限は無料アカウントで100通程度です
 * =====================================================================
 */

var NOTIFY_TO = "";   // ← 通知先メールアドレスを設定（公開リポジトリのため実アドレスは記載していません。稼働中のGASプロジェクト側に設定済み）
var SHEET_ID  = "";   // 例: "1AbC...xyz"（不要なら空のまま）

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ---- 本文の組み立て ----
    var lines = [];
    Object.keys(data).forEach(function (k) {
      if (k.indexOf("―") === 0) { lines.push("", k, ""); }   // 区切り行
      else { lines.push(k + "： " + data[k]); }
    });
    var body = lines.join("\n");

    var company = data["会社名／屋号"] || "（未入力）";
    var siteName = data["サイト名"] || "電気代シミュレーション";
    var subject = "【" + siteName + "】お問い合わせ：" + company;

    // ---- メール送信 ----
    var options = { to: NOTIFY_TO, subject: subject, body: body };
    var replyTo = data["メールアドレス"];
    if (replyTo && replyTo.indexOf("@") > 0) { options.replyTo = replyTo; }
    MailApp.sendEmail(options);

    // ---- スプレッドシート記録（任意） ----
    if (SHEET_ID) { appendToSheet_(data); }

    return json_({ ok: true });
  } catch (err) {
    // 失敗しても管理者には通知しておく
    MailApp.sendEmail(NOTIFY_TO, "【電気代シミュレーション】受信エラー", String(err) + "\n\n" + (e && e.postData ? e.postData.contents : ""));
    return json_({ ok: false, error: String(err) });
  }
}

/** ブラウザで直接URLを開いたときの死活確認用 */
function doGet() {
  return json_({ ok: true, message: "endpoint is alive" });
}

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

/** 動作テスト用：エディタ上でこの関数を実行するとテストメールが届きます */
function testSend() {
  doPost({ postData: { contents: JSON.stringify({
    "サイト名": "かんたん電気代シミュレーション",
    "送信日時": new Date().toLocaleString("ja-JP"),
    "流入経路": "test",
    "メールアドレス": "test@example.com",
    "会社名／屋号": "テスト商店",
    "――― 入力内容 ―――": "",
    "直近1ヶ月の電気代": "120,000 円",
    "想定 年間削減額": "240,000 円"
  }) } });
}
