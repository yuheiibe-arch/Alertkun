/**
 * ==========================================
 * Slack通知用ヘルパー関数 (Block Kit対応)
 * ★UPDATE: @dspart @dsshift へのメンションを追加
 * ==========================================
 */

// ★ Webhook URLを設定（スクリプトプロパティから取得するよう修正）
const SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');

/**
 * マスター監査から呼び出されるSlack送信関数
 * @param {number} totalErrorsCount - 未対応エラーの合計件数
 * @param {object} breakdown - エラー項目の内訳オブジェクト
 * @param {string} sheetUrl - アラートリストシートのURL
 */
function sendSlackAlert(totalErrorsCount, breakdown, sheetUrl) {
  
  // 1. エラー内訳のテキストを生成
  let breakdownText = "";
  for (const [errName, count] of Object.entries(breakdown)) {
    breakdownText += `• *${errName}*: ${count}件\n`;
  }

  // 2. 「設定」シートから現在の除外状況を読み取る
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("設定");
  
  let holidaySetting = "なし";
  let specificClinics = "なし";

  if (settingsSheet) {
    const sData = settingsSheet.getDataRange().getValues();
    const clinicsArr = [];
    
    for (let r = 0; r < sData.length; r++) {
      const keyStr = String(sData[r][0]).trim();
      
      if (keyStr === "年末年始") {
        const isSkip = String(sData[r][3]).trim() === "検知なし";
        if (isSkip) {
          holidaySetting = "年末年始募集 (除外適用中)";
        }
      } 
      else if (keyStr === "特定拠点") {
        for (let c = 1; c < sData[r].length; c++) {
          const cName = String(sData[r][c]).trim();
          if (cName) clinicsArr.push(cName);
        }
      }
    }
    if (clinicsArr.length > 0) {
      specificClinics = clinicsArr.join(", ");
    }
  }

  // 3. Slack Block Kit を使ったリッチなメッセージの構築
  const payload = {
    // スマホ等のポップアップ通知で表示されるプレビューテキスト（ここにもメンションを入れて気づきやすくします）
    "text": `@dspart @dsshift 🚨 シフトチェッカー: 現在 ${totalErrorsCount}件 の未対応エラーがあります`, 
    "blocks": [
      {
        "type": "header",
        "text": {
          "type": "plain_text",
          "text": "🚨アラートくん参上",
          "emoji": true
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          // ★ 本文の冒頭にメンションを追加
          "text": `@dspart @dsshift\nDaily checkが完了しました。\n現在、*${totalErrorsCount}件* の未対応エラーがアラートリストに残っています。\n以下の内訳を確認し、スプレッドシートで対応をお願いします。`
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `*【エラー内訳】*\n${breakdownText}`
        }
      },
      {
        "type": "actions",
        "elements": [
          {
            "type": "button",
            "text": {
              "type": "plain_text",
              "text": "📊 アラートリストを開く",
              "emoji": true
            },
            "value": "open_sheet",
            "url": sheetUrl,
            "style": "primary"
          }
        ]
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "💡 *【担当者へのお願い】*\nアラートリストを確認し、*不要なもの・対応が完了したものには左端のチェックを入れて転記（アーカイブ）* してください。\n\n*＜募集シフト未掲載の除外拠点（現在の設定）＞*\n・除外設定： " + holidaySetting + "\n・特定拠点： " + specificClinics
        }
      }
    ]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  // 4. Webhook経由で送信
  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
    Logger.log("✅ Slack通知の送信に成功しました。");
  } catch (e) {
    Logger.log("❌ Slack送信エラー: " + e.message);
  }
}