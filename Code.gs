function doGet(e) {
  var p = e.parameter;
  var uid = p.uid || '';
  var client = p.client || '';
  var total = p.total || '';
  var link = p.link || '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['建立/更新時間', '客戶名稱', '含稅總計', '編輯連結', 'uid']);
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
