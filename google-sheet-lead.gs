/**
 * Nhận lead từ form "Đặt lịch kiểm tra hồ sơ EUDR" và ghi vào Google Sheet.
 *
 * Cách cài (làm một lần):
 * 1. Mở sheet nhận lead > Tiện ích mở rộng > Apps Script.
 * 2. Xoá code mẫu, dán toàn bộ file này vào, bấm Lưu.
 * 3. Triển khai > Tuỳ chọn triển khai mới > Loại: Ứng dụng web.
 *      Thực thi với tư cách: Tôi
 *      Người có quyền truy cập: Bất kỳ ai
 *    Bấm Triển khai, cấp quyền khi Google hỏi.
 * 4. Sao chép "URL ứng dụng web" (dạng https://script.google.com/macros/s/.../exec)
 *    và dán vào CONFIG.formEndpoint trong index.html.
 *
 * Sửa code sau này: Triển khai > Quản lý triển khai > sửa > Phiên bản: Phiên bản mới.
 * Nếu tạo triển khai mới thay vì sửa, URL sẽ đổi và phải dán lại vào trang.
 */

var SHEET_ID = '17-c13KvMGELMsiH3Mk5wUfK7CwbUvzhS9ATWuVCNMrA';
var SHEET_NAME = ''; // Để trống: ghi vào tab đầu tiên. Hoặc điền tên tab, VD 'Lead EUDR'.

// Thứ tự cột trong sheet: [tên trường gửi từ form, tiêu đề cột]
var COLUMNS = [
  ['time', 'Thời gian'],
  ['name', 'Họ và tên'],
  ['phone', 'Số điện thoại / Zalo'],
  ['email', 'Email'],
  ['company_name', 'Công ty'],
  ['industry', 'Ngành hàng'],
  ['page', 'Trang'],
  ['utm_campaign', 'utm_campaign'],
  ['utm_source', 'utm_source'],
  ['utm_medium', 'utm_medium'],
  ['utm_content', 'utm_content'],
  ['utm_term', 'utm_term'],
  ['gclid', 'gclid'],
  ['fbp', 'fbp'],
  ['fbclid', 'fbclid'],
  ['ip_address', 'ip_address'],
  ['company_address', 'Khu vực'],
  ['company_size', 'Quy mô nhân sự']
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // hai khách gửi cùng lúc không ghi đè lên nhau
  try {
    var p = (e && e.parameter) || {};
    // Hỗ trợ cả trường cũ và trường mới chuẩn hóa
    p.company_name = p.company_name || p.company;
    p.company_address = p.company_address || p.region;
    p.company_size = p.company_size || p.size;

    var book = SpreadsheetApp.openById(SHEET_ID);
    var sheet = SHEET_NAME ? book.getSheetByName(SHEET_NAME) : book.getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS.map(function (c) { return c[1]; }));
      sheet.setFrozenRows(1);
    } else {
      ensureHeaders(sheet);
    }
    var row = COLUMNS.map(function (c) {
      if (c[0] === 'time') return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
      return asText(p[c[0]]);
    });
    sheet.appendRow(row);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Sheet đã có dòng tiêu đề từ trước: điền tiêu đề cho các cột mới thêm vào COLUMNS (chỉ điền ô trống, không sửa ô đã có).
function ensureHeaders(sheet) {
  var range = sheet.getRange(1, 1, 1, COLUMNS.length);
  var head = range.getValues()[0];
  var changed = false;
  COLUMNS.forEach(function (c, i) {
    if (String(head[i]).trim() === '') { head[i] = c[1]; changed = true; }
  });
  if (changed) range.setValues([head]);
}

// Mở URL trên trình duyệt để kiểm tra script đã chạy.
function doGet() {
  return reply({ ok: true, info: 'Lead endpoint đang hoạt động' });
}

// Giữ số 0 đầu của số điện thoại, và không để ô bắt đầu bằng = + - @ bị hiểu là công thức.
function asText(v) {
  v = v == null ? '' : String(v).trim().slice(0, 500);
  return /^[=+\-@0]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
