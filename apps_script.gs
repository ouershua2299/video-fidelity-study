/**
 * 视频保真度评测：接收网页提交的答案，追加到当前表格的 responses 工作表。
 * 每次请求是一个数组（一组或多组答案，包括之前未成功提交、自动重试的）。
 */
const HEADERS = ['received', 'rater', 'pair', 'left', 'right', 'reference', 'check',
                 'closer_to_reference', 'ms', 'answered_at'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('responses') || ss.insertSheet('responses');
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
    let rows = JSON.parse(e.postData.contents);
    if (!Array.isArray(rows)) rows = [rows];
    const now = new Date();
    rows.forEach(r => sh.appendRow([now, r.rater, r.pair, r.left, r.right, r.ref, r.check ? 1 : 0,
                                    r.q_ref, r.ms, new Date(r.at)]));
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
