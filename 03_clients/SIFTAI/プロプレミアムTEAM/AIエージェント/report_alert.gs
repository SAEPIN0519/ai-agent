/**
 * 日報提出状況 アラートスクリプト
 *
 * 動作概要:
 *   - Google Sheets「DB_Daily-Report」シートを読み取り
 *   - 会員ごとに「最後に日報を提出した日」から経過日数を計算
 *   - 3日以上未提出 → エルメでアナウンス送信
 *   - 5日以上未提出 → エルメで15分面談の申し込みを促進
 *
 * 設定方法:
 *   1. このスクリプトを会員カルテのスプレッドシートに貼り付け
 *   2. 下記「設定」の ERME_WEBHOOK_URL に実際のURLを入力
 *   3. トリガーを設定: 毎日 22:00 に checkReportSubmissions() を実行
 */

// ==========================================
// 設定（ここだけ変更すればOK）
// ==========================================
const CONFIG = {
  SHEET_NAME: "DB_Daily-Report",  // シート名
  COL_NAME: "氏名",               // 会員名の列ヘッダー
  COL_DATE: "日付",               // 日報提出日の列ヘッダー

  DAYS_ANNOUNCE: 3,               // 何日で全体アナウンスするか
  DAYS_MEETING: 5,                // 何日で面談促進するか

  // エルメ Webhook URL（エルメの管理画面から取得してここに貼り付け）
  ERME_WEBHOOK_URL: "https://your-erme-webhook-url-here",

  // 15分面談の申し込みURL
  MEETING_URL: "https://your-meeting-url-here",

  // 冴香さんへのアラート送信先（Gmail）
  ALERT_EMAIL: "your-email@example.com",
};

// ==========================================
// メイン処理（毎日22:00に自動実行）
// ==========================================
function checkReportSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log("エラー: シート「" + CONFIG.SHEET_NAME + "」が見つかりません");
    return;
  }

  // シートデータ取得
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 列番号を特定
  const nameCol = headers.indexOf(CONFIG.COL_NAME);
  const dateCol = headers.indexOf(CONFIG.COL_DATE);

  if (nameCol === -1 || dateCol === -1) {
    Logger.log("エラー: 列「" + CONFIG.COL_NAME + "」または「" + CONFIG.COL_DATE + "」が見つかりません");
    return;
  }

  // 会員ごとの最新提出日を集計
  const lastSubmitMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol];
    const date = row[dateCol];

    if (!name || !date) continue;

    const dateObj = new Date(date);
    if (isNaN(dateObj.getTime())) continue;

    if (!lastSubmitMap[name] || dateObj > lastSubmitMap[name]) {
      lastSubmitMap[name] = dateObj;
    }
  }

  // 今日の日付（時刻なし）
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 停滞会員を分類
  const announce3days = []; // 3日連続未提出（アナウンス対象）
  const meeting5days = [];  // 5日連続未提出（面談促進対象）

  for (const name in lastSubmitMap) {
    const last = new Date(lastSubmitMap[name]);
    last.setHours(0, 0, 0, 0);

    // 経過日数 = 今日 - 最終提出日
    const diffDays = Math.floor((today - last) / (1000 * 60 * 60 * 24));

    if (diffDays >= CONFIG.DAYS_MEETING) {
      meeting5days.push({ name: name, days: diffDays });
    } else if (diffDays >= CONFIG.DAYS_ANNOUNCE) {
      announce3days.push({ name: name, days: diffDays });
    }
  }

  Logger.log("3日以上未提出: " + announce3days.length + "名");
  Logger.log("5日以上未提出: " + meeting5days.length + "名");

  // エルメ送信
  if (announce3days.length > 0) {
    sendAnnouncement(announce3days);
  }
  if (meeting5days.length > 0) {
    sendMeetingPromotion(meeting5days);
  }

  // 冴香さんへ週次レポートメール（月曜のみ）
  if (today.getDay() === 1) {
    sendWeeklyReport(announce3days, meeting5days, lastSubmitMap, today);
  }
}

// ==========================================
// エルメ送信：3日連続アナウンス
// ==========================================
function sendAnnouncement(members) {
  const names = members.map(m => m.name).join("、");
  const message =
    "【日報リマインド】\n\n" +
    "最近、日報の投稿が止まっているメンバーへ。\n\n" +
    "小さなことでもOKです。\n" +
    "「今日はここまでやった」の記録が、\n" +
    "明日のあなたの行動につながります。\n\n" +
    "一緒に前に進みましょう！";

  callErmeWebhook(message);
  Logger.log("アナウンス送信完了: " + names);
}

// ==========================================
// エルメ送信：5日連続 面談促進
// ==========================================
function sendMeetingPromotion(members) {
  members.forEach(function(member) {
    const message =
      "【" + member.name + " さんへ】\n\n" +
      "ここ " + member.days + " 日間、活動が止まっているのが気になっています。\n\n" +
      "うまくいかない理由があるなら、\n" +
      "一緒に整理しませんか？\n\n" +
      "15分のオンライン面談を無料でご用意しています。\n" +
      "気軽に話しかけてください。\n\n" +
      "▼ 面談申し込みはこちら\n" +
      CONFIG.MEETING_URL;

    callErmeWebhook(message);
    Logger.log("面談促進送信: " + member.name + "（" + member.days + "日未提出）");
  });
}

// ==========================================
// エルメ Webhook 呼び出し
// ==========================================
function callErmeWebhook(message) {
  if (CONFIG.ERME_WEBHOOK_URL === "https://your-erme-webhook-url-here") {
    Logger.log("[エルメ未設定] 送信予定メッセージ:\n" + message);
    return;
  }

  const payload = JSON.stringify({ message: message });
  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(CONFIG.ERME_WEBHOOK_URL, options);
    Logger.log("エルメ送信ステータス: " + response.getResponseCode());
  } catch (e) {
    Logger.log("エルメ送信エラー: " + e.message);
  }
}

// ==========================================
// 冴香さんへの週次レポートメール（月曜のみ）
// ==========================================
function sendWeeklyReport(announce3, meeting5, lastSubmitMap, today) {
  const totalMembers = Object.keys(lastSubmitMap).length;
  const stoppedCount = announce3.length + meeting5.length;
  const activeCount = totalMembers - stoppedCount;
  const submitRate = totalMembers > 0
    ? Math.round((activeCount / totalMembers) * 100)
    : 0;

  let body = "【週次 日報提出状況レポート - " + Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd") + "】\n\n";

  body += "== 緊急フォロー（5日以上手が止まっている）==\n";
  if (meeting5.length > 0) {
    meeting5.forEach(m => { body += "  - " + m.name + ": " + m.days + "日連続未提出（面談促進送信済み）\n"; });
  } else {
    body += "  なし\n";
  }

  body += "\n== アナウンス送信済み（3日連続）==\n";
  if (announce3.length > 0) {
    announce3.forEach(m => { body += "  - " + m.name + ": " + m.days + "日連続未提出\n"; });
  } else {
    body += "  なし\n";
  }

  body += "\n== チーム全体 ==\n";
  body += "  全会員数: " + totalMembers + "名\n";
  body += "  アクティブ: " + activeCount + "名\n";
  body += "  提出率: " + submitRate + "%\n";

  GmailApp.sendEmail(
    CONFIG.ALERT_EMAIL,
    "【週次レポート】日報提出状況 " + Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd"),
    body
  );
  Logger.log("週次レポートメール送信完了");
}
