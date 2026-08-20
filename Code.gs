function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  // 舊版本只有一個未命名分頁在用（=報價單資料），第一次跑到這裡順便正名，資料不會不見
  if (name === '報價單' && ss.getSheets().length === 1) {
    var only = ss.getSheets()[0];
    only.setName('報價單');
    return only;
  }
  return ss.insertSheet(name);
}

function doGet(e) {
  var p = e.parameter;
  var docType = p.doc === 'contract' ? 'contract' : 'quote';
  var uid = p.uid || '';
  var client = p.client || '';
  var total = p.total || '';
  var link = p.link || '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = docType === 'contract' ? '合約書' : '報價單';
  var sheet = getOrCreateSheet(ss, sheetName);

  if (sheet.getLastRow() === 0) {
    var amountLabel = docType === 'contract' ? '總金額' : '含稅總計';
    sheet.appendRow(['建立/更新時間', '客戶名稱', amountLabel, '編輯連結', 'uid']);
  }

  var lastRow = sheet.getLastRow();
  var uidCol = 5;
  var rowIndex = -1;
  if (lastRow > 1) {
    var uids = sheet.getRange(2, uidCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < uids.length; i++) {
      if (uids[i][0] === uid) {
        rowIndex = i + 2;
        break;
      }
    }
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  var row = [now, client, total, link, uid];

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return ContentService.createTextOutput('ok');
}
