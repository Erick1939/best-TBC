/**
 * Apps Script Web App — thống kê sử dụng tool Trưởng Bưu Cục theo từng bưu cục.
 * Cách dùng: dán toàn bộ code này vào Google Sheet → Extensions → Apps Script,
 * rồi Deploy → New deployment → Web app (xem hướng dẫn chi tiết kèm theo).
 *
 * Mỗi lần userscript tải dữ liệu "hôm nay" (mở tool / bấm Tải lại), nó gọi
 * ?office=<tên bưu cục>&v=<version>. Script tự tính:
 *  - Cột B "Số lần sử dụng trong ngày": đếm số lần ping trong CÙNG 1 ngày,
 *    tự về 1 khi sang ngày mới.
 *  - Cột C "Số ngày sử dụng": tổng số ngày KHÁC NHAU bưu cục đó từng ping
 *    (cộng dồn, không bao giờ giảm).
 *  - Cột D "Ngày gần nhất" chỉ để script tự so sánh ngày, không cần quan tâm.
 */
function doGet(e) {
  var office = ((e && e.parameter && e.parameter.office) || '').toString().trim();
  var ver = ((e && e.parameter && e.parameter.v) || '').toString().trim();
  if (!office) return ContentService.createTextOutput('missing office param');

  var tz = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Data') || ss.insertSheet('Data');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Tên bưu cục', 'Số lần sử dụng trong ngày', 'Số ngày sử dụng', 'Ngày gần nhất', 'Phiên bản gần nhất']);
  }

  var data = sh.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === office) { foundRow = i; break; }
  }

  if (foundRow === -1) {
    // Bưu cục mới toanh: 1 lần dùng hôm nay, 1 ngày đã dùng
    sh.appendRow([office, 1, 1, todayStr, ver]);
  } else {
    var row = foundRow + 1; // 1-indexed trong Sheet, +1 vì có dòng tiêu đề
    var lastDateStr = Utilities.formatDate(new Date(data[foundRow][3]), tz, 'yyyy-MM-dd');
    if (lastDateStr === todayStr) {
      // Vẫn trong ngày hôm nay: chỉ tăng "Số lần sử dụng trong ngày"
      sh.getRange(row, 2).setValue((Number(data[foundRow][1]) || 0) + 1);
    } else {
      // Sang ngày mới: reset "Số lần sử dụng trong ngày" về 1, cộng thêm 1 "Số ngày sử dụng"
      sh.getRange(row, 2).setValue(1);
      sh.getRange(row, 3).setValue((Number(data[foundRow][2]) || 0) + 1);
      sh.getRange(row, 4).setValue(todayStr);
    }
    sh.getRange(row, 5).setValue(ver);
  }

  return ContentService.createTextOutput('ok');
}
