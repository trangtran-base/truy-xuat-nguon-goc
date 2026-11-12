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
var SHEET_GID = '874816478'; // GID của tab nhận lead (từ URL: gid=874816478)
var SHEET_NAME = ''; // Để trống hoặc điền tên tab. Ưu tiên tìm theo SHEET_GID trước.
var NOTIFY_EMAILS = 'trang.tran@base.vn, hoa.nguyen02@base.vn'; // Danh sách email nhận thông báo khi có lead mới

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
    var sheet = getTargetSheet(book);
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

    // Gửi email thông báo cho team phụ trách lead
    sendLeadNotification(p);

    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Tìm đúng tab theo GID hoặc Tên tab
function getTargetSheet(book) {
  var sheets = book.getSheets();
  if (SHEET_GID) {
    for (var i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getSheetId()) === String(SHEET_GID)) {
        return sheets[i];
      }
    }
  }
  if (SHEET_NAME) {
    var s = book.getSheetByName(SHEET_NAME);
    if (s) return s;
  }
  return sheets[0];
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

// Gửi email thông báo khi có lead đăng ký mới
function sendLeadNotification(p) {
  p = p || {};
  if (!NOTIFY_EMAILS) return;
  try {
    var timeStr = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
    var subject = '[Lead Mới - EUDR] ' + (p.name || 'Khách hàng') + ' - ' + (p.company_name || p.phone || '');

    var htmlBody = ''
      + '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">'
      + '  <div style="background-color: #1b7a4e; color: #ffffff; padding: 18px 24px;">'
      + '    <h2 style="margin: 0; font-size: 20px;">Thông Báo Lead Đăng Ký Mới</h2>'
      + '    <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Giải pháp Truy xuất nguồn gốc chuẩn EUDR | Base.vn</p>'
      + '  </div>'
      + '  <div style="padding: 24px; background-color: #ffffff;">'
      + '    <h3 style="margin-top: 0; color: #1c2126; border-bottom: 2px solid #e6f2ec; padding-bottom: 8px;">Thông tin khách hàng:</h3>'
      + '    <table style="width: 100%; border-collapse: collapse; font-size: 14px; line-height: 1.6;">'
      + '      <tr><td style="padding: 6px 0; color: #6b737b; width: 140px;"><strong>Họ và tên:</strong></td><td style="padding: 6px 0; color: #1c2126;"><strong>' + (p.name || '-') + '</strong></td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Số điện thoại:</strong></td><td style="padding: 6px 0; color: #1b7a4e;"><strong><a href="tel:' + (p.phone || '') + '" style="color: #1b7a4e; text-decoration: none;">' + (p.phone || '-') + '</a></strong></td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Email:</strong></td><td style="padding: 6px 0; color: #1c2126;"><a href="mailto:' + (p.email || '') + '" style="color: #1b7a4e;">' + (p.email || '-') + '</a></td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Công ty:</strong></td><td style="padding: 6px 0; color: #1c2126;">' + (p.company_name || '-') + '</td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Khu vực:</strong></td><td style="padding: 6px 0; color: #1c2126;">' + (p.company_address || '-') + '</td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Quy mô nhân sự:</strong></td><td style="padding: 6px 0; color: #1c2126;">' + (p.company_size || '-') + '</td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Ngành hàng:</strong></td><td style="padding: 6px 0; color: #1c2126;">' + (p.industry || '-') + '</td></tr>'
      + '      <tr><td style="padding: 6px 0; color: #6b737b;"><strong>Thời gian:</strong></td><td style="padding: 6px 0; color: #1c2126;">' + timeStr + '</td></tr>'
      + '    </table>'
      + '    <h3 style="margin-top: 24px; color: #1c2126; border-bottom: 2px solid #e6f2ec; padding-bottom: 8px;">Thông tin nguồn chiến dịch:</h3>'
      + '    <table style="width: 100%; border-collapse: collapse; font-size: 13px; color: #4b535b;">'
      + '      <tr><td style="padding: 4px 0; width: 140px;">utm_source:</td><td>' + (p.utm_source || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">utm_campaign:</td><td>' + (p.utm_campaign || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">utm_medium:</td><td>' + (p.utm_medium || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">utm_content:</td><td>' + (p.utm_content || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">utm_term:</td><td>' + (p.utm_term || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">gclid:</td><td>' + (p.gclid || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">fbclid:</td><td>' + (p.fbclid || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">fbp:</td><td>' + (p.fbp || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">IP Address:</td><td>' + (p.ip_address || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">Trang đăng ký:</td><td>' + (p.page || '-') + '</td></tr>'
      + '    </table>'
      + '  </div>'
      + '  <div style="background-color: #f4f6f5; padding: 12px 24px; font-size: 12px; color: #6b737b; text-align: center;">'
      + '    Email tự động gửi từ hệ thống Landing Page Base EUDR'
      + '  </div>'
      + '</div>';

    MailApp.sendEmail({
      to: NOTIFY_EMAILS,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (mailErr) {
    Logger.log('Lỗi gửi email thông báo: ' + mailErr);
  }
}

// Hàm chạy thử để kích hoạt hộp thoại cấp quyền gửi email của Google (Chỉ cần chạy 1 lần)
function testSendEmail() {
  sendLeadNotification({
    name: 'Nguyễn Văn Test',
    phone: '0988888888',
    email: 'test@base.vn',
    company_name: 'Công ty Test Base',
    company_address: 'Khu vực Miền Bắc',
    company_size: '31 - 60 nhân sự',
    industry: 'Gỗ, giấy, bao bì',
    utm_source: 'facebook',
    utm_campaign: 'eudr_2026',
    page: 'https://signup.base.vn/truy-xuat-nguon-goc/'
  });
}
