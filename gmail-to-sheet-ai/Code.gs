/**
 * Gmail → AI抽出 → Googleスプレッドシート 自動登録
 *
 * 指定ラベルの未処理メール（本文＋PDF添付）から、Claude APIで項目を抽出し、
 * 「確認待ち」シートに追記する。人が確認してチェックを入れた行だけが「確定」シートへ移る。
 *
 * 設定: スクリプトプロパティに ANTHROPIC_API_KEY を登録すること。
 */

const CONFIG = {
  LABEL_TARGET: 'ai-import',        // 取り込み対象のGmailラベル
  LABEL_DONE: 'ai-import/done',     // 処理済みラベル（重複防止）
  LABEL_ERROR: 'ai-import/error',   // 失敗時ラベル（再実行対象）
  SHEET_REVIEW: '確認待ち',
  SHEET_FIXED: '確定',
  SHEET_LOG: 'ログ',
  MODEL: 'claude-sonnet-5',
  MAX_THREADS_PER_RUN: 20,
  // 抽出したい項目（ここを変えれば別業務に流用できる）
  FIELDS: [
    { key: 'property_name', label: '物件名' },
    { key: 'address', label: '所在地' },
    { key: 'price', label: '価格' },
    { key: 'area', label: '面積' },
    { key: 'contact', label: '担当者・連絡先' },
    { key: 'deadline', label: '期限' },
  ],
};

const HEADER = ['確認OK', 'メッセージID', '受信日時', '件名', ...CONFIG.FIELDS.map(f => f.label), '抽出メモ'];

/** 時間主導トリガーで定期実行する入口 */
function importMails() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) return; // 二重実行防止
  try {
    const sheet = getOrCreateSheet_(CONFIG.SHEET_REVIEW, HEADER);
    const known = new Set(sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat());
    const threads = GmailApp.search(`label:${CONFIG.LABEL_TARGET} -label:${CONFIG.LABEL_DONE}`, 0, CONFIG.MAX_THREADS_PER_RUN);

    threads.forEach(thread => {
      try {
        thread.getMessages().forEach(msg => {
          if (known.has(msg.getId())) return; // 同一メールの重複登録防止
          const result = extractWithClaude_(msg);
          sheet.appendRow([
            false, msg.getId(), msg.getDate(), msg.getSubject(),
            ...CONFIG.FIELDS.map(f => result[f.key] ?? ''),
            result.note ?? '',
          ]);
        });
        thread.addLabel(getLabel_(CONFIG.LABEL_DONE));
        thread.removeLabel(getLabel_(CONFIG.LABEL_ERROR));
      } catch (e) {
        thread.addLabel(getLabel_(CONFIG.LABEL_ERROR));
        log_('ERROR', thread.getFirstMessageSubject(), e.message);
      }
    });
    sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).insertCheckboxes();
  } finally {
    lock.releaseLock();
  }
}

/** 「確認OK」にチェックされた行を確定シートへ移す */
function commitReviewed() {
  const review = getOrCreateSheet_(CONFIG.SHEET_REVIEW, HEADER);
  const fixed = getOrCreateSheet_(CONFIG.SHEET_FIXED, HEADER.slice(1));
  const rows = review.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === true) {
      fixed.appendRow(rows[i].slice(1));
      review.deleteRow(i + 1);
    }
  }
}

/** エラーになったスレッドを再処理対象に戻す */
function retryErrors() {
  GmailApp.search(`label:${CONFIG.LABEL_ERROR}`).forEach(t => {
    t.removeLabel(getLabel_(CONFIG.LABEL_DONE));
    t.removeLabel(getLabel_(CONFIG.LABEL_ERROR));
  });
  importMails();
}

function extractWithClaude_(msg) {
  const content = [];
  msg.getAttachments().forEach(att => {
    if (att.getContentType() === 'application/pdf') {
      content.push({
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: Utilities.base64Encode(att.getBytes()) },
      });
    }
  });
  const fieldList = CONFIG.FIELDS.map(f => `- ${f.key}: ${f.label}`).join('\n');
  content.push({
    type: 'text',
    text: `以下のメール（と添付PDF）から項目を抽出し、JSONのみを返してください。
見つからない項目は空文字にし、推測で埋めないこと。判断に迷った点は note に書くこと。
項目:
${fieldList}
- note: 抽出メモ

件名: ${msg.getSubject()}
本文:
${msg.getPlainBody().slice(0, 20000)}`,
  });

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({ model: CONFIG.MODEL, max_tokens: 1024, messages: [{ role: 'user', content }] }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error(`API ${res.getResponseCode()}: ${res.getContentText().slice(0, 200)}`);
  const text = JSON.parse(res.getContentText()).content[0].text;
  const json = text.match(/\{[\s\S]*\}/);
  if (!json) throw new Error('JSONを抽出できませんでした');
  return JSON.parse(json[0]);
}

function getOrCreateSheet_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function log_(level, subject, message) {
  getOrCreateSheet_(CONFIG.SHEET_LOG, ['日時', 'レベル', '件名', '内容']).appendRow([new Date(), level, subject, message]);
}
