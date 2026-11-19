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

var SCRIPT_VERSION = 'v5-ab-variant'; // Mở URL /exec trên trình duyệt để xem bản đang chạy

// [tên trường gửi từ form, tiêu đề cột CHUẨN trong sheet, ...tên tiêu đề cũ]
// Script ghi theo TÊN tiêu đề cột nên thứ tự cột trong sheet không quan trọng.
// Gặp tiêu đề cũ (vd "Họ và tên") thì tự đổi thành tiêu đề chuẩn (vd "name").
var COLUMNS = [
  ['time', 'time', 'Thời gian', 'Thời điểm', 'Ngày'],
  ['name', 'name', 'Họ và tên', 'Họ tên', 'Tên'],
  ['phone', 'phone', 'SĐT / Zalo', 'SĐT/Zalo', 'Số điện thoại / Zalo', 'SĐT', 'Số điện thoại', 'Điện thoại'],
  ['email', 'Email'],
  ['company_name', 'company_name', 'Công ty', 'Tên công ty', 'company'],
  ['industry', 'industry', 'Ngành hàng', 'Ngành'],
  ['page_url', 'page_url', 'Trang', 'page', 'URL', 'Landing page', 'Trang đăng ký'],
  ['variant', 'variant', 'Phiên bản'],
  ['utm_campaign', 'utm_campaign'],
  ['utm_source', 'utm_source'],
  ['utm_medium', 'utm_medium'],
  ['utm_content', 'utm_content'],
  ['utm_term', 'utm_term'],
  ['gclid', 'gclid'],
  ['fbp', 'fbp'],
  ['fbclid', 'fbclid'],
  ['ip_address', 'ip_address'],
  ['company_address', 'company_address', 'Khu vực', 'region'],
  ['company_size', 'company_size', 'Quy mô nhân sự', 'Quy mô', 'size']
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
    p.page_url = p.page_url || p.page;

    var book = SpreadsheetApp.openById(SHEET_ID);
    var sheet = getTargetSheet(book);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS.map(function (c) { return c[1]; }));
      sheet.setFrozenRows(1);
    }
    // Ghi theo TÊN tiêu đề cột (không theo vị trí), nên đổi thứ tự cột trong sheet vẫn ghi đúng.
    var colIndex = mapHeaders(sheet);
    var width = Math.max(sheet.getLastColumn(), 1);
    var row = [];
    for (var i = 0; i < width; i++) row.push('');
    COLUMNS.forEach(function (c) {
      var idx = colIndex[c[0]];
      if (idx == null) return;
      row[idx] = c[0] === 'time'
        ? Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss')
        : asText(p[c[0]]);
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

// Đọc dòng tiêu đề, trả về { tên trường: vị trí cột (0-based) }.
// Khớp tiêu đề với tên trường (vd "utm_source") hoặc tiêu đề hiển thị (vd "Họ và tên"), không phân biệt hoa thường.
// Trường nào chưa có cột thì thêm cột mới ở cuối sheet, không đụng tới các cột đã có.
function mapHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  var head = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var norm = function (s) { return String(s == null ? '' : s).normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase(); };
  var byTitle = {};
  head.forEach(function (h, i) {
    var k = norm(h);
    if (k && !(k in byTitle)) byTitle[k] = i;
  });
  var map = {};
  var missing = [];
  var renamed = false;
  COLUMNS.forEach(function (c) {
    var idx = null;
    for (var n = 0; n < c.length && idx == null; n++) {
      if (norm(c[n]) in byTitle) idx = byTitle[norm(c[n])];
    }
    if (idx == null) { missing.push(c); return; }
    map[c[0]] = idx;
    if (String(head[idx]).trim() !== c[1]) { head[idx] = c[1]; renamed = true; } // đổi tiêu đề cũ thành tiêu đề chuẩn
  });
  if (renamed) sheet.getRange(1, 1, 1, head.length).setValues([head]);
  if (missing.length) {
    // Thêm ngay sau ô tiêu đề cuối cùng có chữ (bản cũ có thể đã ghi dữ liệu vào cột chưa có tiêu đề).
    var start = 0;
    head.forEach(function (h, i) { if (norm(h)) start = i + 1; });
    sheet.getRange(1, start + 1, 1, missing.length).setValues([missing.map(function (c) { return c[1]; })]);
    missing.forEach(function (c, j) { map[c[0]] = start + j; });
  }
  return map;
}

/* ===== Sửa các dòng lead cũ bị ghi lệch cột =====
 * Bản cũ (12/11 – 16/11/2026) ghi theo VỊ TRÍ cột với thứ tự dưới đây, nên dữ liệu bị lệch
 * so với tiêu đề sheet (vd email nằm ở cột "Công ty", URL trang nằm ở cột "utm_source").
 * Cách dùng trong Apps Script (chọn tên hàm ở thanh trên cùng > Chạy):
 *   1. Chạy previewFixOldRows  -> xem Nhật ký thực thi: số dòng sẽ sửa + vài dòng mẫu trước/sau.
 *   2. Chạy fixOldRows         -> tự tạo tab sao lưu "Backup lead ..." rồi sửa.
 * Chỉ sửa dòng có dấu hiệu lệch: ô ở vị trí cột 4 chứa "@" (email) VÀ ô ở vị trí cột 7 bắt đầu bằng "http" (URL trang).
 * Chạy lại nhiều lần cũng không sao: dòng đã sửa không còn dấu hiệu này.
 */
var OLD_ORDER = ['time', 'name', 'phone', 'email', 'company_name', 'industry', 'page_url',
  'utm_campaign', 'utm_source', 'utm_medium', 'utm_content', 'utm_term', 'gclid', 'fbp',
  'fbclid', 'ip_address', 'company_address', 'company_size'];

function previewFixOldRows() { fixOldRows_(false); }
function fixOldRows() { fixOldRows_(true); }

function fixOldRows_(apply) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var book = SpreadsheetApp.openById(SHEET_ID);
    var sheet = getTargetSheet(book);
    var map = mapHeaders(sheet); // thêm cột còn thiếu (vd "Trang") nếu chưa có
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) { Logger.log('Sheet chưa có dữ liệu.'); return; }
    var head = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Cột trong vùng 18 cột đầu không thuộc trường nào: không đụng tới, chỉ cảnh báo.
    var targets = {};
    Object.keys(map).forEach(function (k) { targets[map[k]] = true; });
    var untouched = [];
    for (var c = 0; c < OLD_ORDER.length && c < lastCol; c++) {
      if (!targets[c]) untouched.push(head[c] || ('cột ' + (c + 1)));
    }
    if (untouched.length) Logger.log('Lưu ý: các cột sau không thuộc trường nào nên giữ nguyên: ' + untouched.join(', '));

    var width = 0;
    Object.keys(map).forEach(function (k) { width = Math.max(width, map[k] + 1); });
    width = Math.max(width, Math.min(OLD_ORDER.length, lastCol));

    var fixes = [];
    data.forEach(function (row, r) {
      var emailAt = String(row[3] == null ? '' : row[3]);
      var pageAt = String(row[6] == null ? '' : row[6]);
      if (emailAt.indexOf('@') === -1 || !/^https?:\/\//i.test(pageAt)) return;
      if (map.email === 3 && map.page_url === 6) return; // tiêu đề trùng thứ tự cũ: không lệch
      var out = row.slice(0, width);
      while (out.length < width) out.push('');
      OLD_ORDER.forEach(function (k, i) {
        if (map[k] != null) out[map[k]] = i < row.length ? row[i] : '';
      });
      fixes.push({ r: r + 2, before: row, after: out });
    });

    Logger.log('Số dòng bị lệch cần sửa: ' + fixes.length + ' / ' + data.length);
    fixes.slice(0, 5).forEach(function (f) {
      var show = function (arr) {
        return ['name', 'email', 'company_name', 'page_url', 'utm_source', 'utm_medium', 'utm_campaign']
          .map(function (k) { return k + '=' + arr[map[k]]; }).join(' | ');
      };
      Logger.log('Dòng ' + f.r + '\n  TRƯỚC: ' + show(f.before) + '\n  SAU:   ' + show(f.after));
    });
    if (!apply) { Logger.log('Đây là bản xem trước, chưa sửa gì. Chạy fixOldRows để sửa thật.'); return; }
    if (!fixes.length) return;

    var stamp = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd-MM HH:mm:ss');
    sheet.copyTo(book).setName('Backup lead ' + stamp);
    fixes.forEach(function (f) {
      var vals = f.after.map(function (v) { return typeof v === 'string' ? asText(v) : v; });
      sheet.getRange(f.r, 1, 1, width).setValues([vals]);
    });
    Logger.log('Đã sửa ' + fixes.length + ' dòng. Bản sao lưu: tab "Backup lead ' + stamp + '".');
  } finally {
    lock.releaseLock();
  }
}

// Mở URL trên trình duyệt để kiểm tra script đã chạy.
function doGet() {
  return reply({ ok: true, info: 'Lead endpoint đang hoạt động', version: SCRIPT_VERSION });
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
      + '      <tr><td style="padding: 4px 0;">Trang đăng ký:</td><td>' + (p.page_url || p.page || '-') + '</td></tr>'
      + '      <tr><td style="padding: 4px 0;">Phiên bản trang (A/B):</td><td>' + (p.variant || '-') + '</td></tr>'
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
    page_url: 'https://signup.base.vn/truy-xuat-nguon-goc/'
  });
}

// Chạy 1 lần trong Apps Script để đổi ngay tiêu đề cột cũ sang tên chuẩn (không cần chờ có lead mới).
function standardizeHeaders() {
  var sheet = getTargetSheet(SpreadsheetApp.openById(SHEET_ID));
  mapHeaders(sheet);
  Logger.log('Tiêu đề hiện tại: ' + sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].join(' | '));
}
