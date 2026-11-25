/**
 * セキュリティ訓練サイト用のバックエンド(Google Apps Script)
 *
 * 【使い方】
 * 1. Googleスプレッドシートを新規作成する
 * 2. 「拡張機能」→「Apps Script」を開く
 * 3. このファイルの内容を貼り付けて保存する
 * 4. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」を選択
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. 発行されたURLを landing_page.html の ENDPOINT_URL に貼り付ける
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.type === "visit") {
      var visitSheet = getOrCreateSheet(ss, "Visits", ["timestamp", "visitorId"]);
      visitSheet.appendRow([data.timestamp, data.visitorId]);

    } else if (data.type === "survey") {
      var surveySheet = getOrCreateSheet(ss, "SurveyResponses",
        ["timestamp", "visitorId", "q1", "q2", "q3", "q4", "q5"]);
      var a = data.answers || {};
      surveySheet.appendRow([
        data.timestamp,
        data.visitorId,
        a.q1 || "",
        a.q2 || "",
        a.q3 || "",
        a.q4 || "",
        a.q5 || ""
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(ss, name, headerRow) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headerRow);
  }
  return sheet;
}

/**
 * 管理用: ユニーク訪問者数を確認したいときに、このスプレッドシート上で
 * 下記のような数式をセルに入力すると集計できます。
 *
 * ユニーク訪問者数:
 *   =COUNTA(UNIQUE(Visits!B2:B))
 *
 * アンケート回答数:
 *   =COUNTA(SurveyResponses!A2:A)
 */
