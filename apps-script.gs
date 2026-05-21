// Workshop "Từ Ý Tưởng Thành Website" – Google Apps Script
// Dán toàn bộ code này vào Apps Script Editor, sau đó Deploy lại → Web App

const SHEET_NAME = 'Đăng ký';
const SEPAY_API  = 'https://pay.sepay.vn/v1/checkout/init';

// ── Đọc credentials từ Script Properties (an toàn, không lộ trong code) ──
function getCreds() {
  const props = PropertiesService.getScriptProperties();
  return {
    merchantId: props.getProperty('SEPAY_MERCHANT_ID'),
    secretKey:  props.getProperty('SEPAY_SECRET_KEY'),
  };
}

// ── Nhận GET request từ landing page (form submit) ──
function doGet(e) {
  const p = e.parameter;
  // Chỉ lưu khi có đủ 3 trường bắt buộc: tên, email, SĐT
  if (p && p.name && p.name.trim() && p.email && p.email.trim() && p.phone && p.phone.trim()) {
    return handleFormSubmit(p);
  }
  return HtmlService.createHtmlOutput('<p>OK</p>');
}

// ── Nhận mọi POST request: từ form đăng ký HOẶC webhook SePay ──
function doPost(e) {
  try {
    // Form submit từ landing page (application/x-www-form-urlencoded)
    if (e.postData.type === 'application/x-www-form-urlencoded') {
      return handleFormSubmit(e.parameter);
    }

    // JSON: webhook SePay hoặc fetch từ browser
    const data = JSON.parse(e.postData.contents);

    // SePay bank monitoring webhook: có trường "content" (nội dung CK) và "transferType"
    // SePay checkout IPN: có trường "notification_type"
    if (data.transferType || data.content || data.notification_type === 'ORDER_PAID') {
      return handleSePayWebhook(data);
    }
    return handleRegistrationJson(data);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ result: 'error', message: err.toString() });
  }
}

// ── Xử lý form submit: lưu sheet + tạo đơn SePay + redirect ──
function handleFormSubmit(params) {
  const orderId = 'WS-' + Date.now();
  const sheet   = getOrCreateSheet();

  sheet.appendRow([
    new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }),
    params.name  || '',
    params.email || '',
    params.phone || '',
    params.job   || '',
    orderId,
    'Chờ thanh toán',
  ]);

  const bankId      = 'MB';
  const accountNo   = '0901277034';
  const accountName = 'BUI LE THAO VY';
  const amount      = 1860000;
  const qrUrl = 'https://img.vietqr.io/image/' + bankId + '-' + accountNo +
    '-compact2.png?amount=' + amount +
    '&addInfo=' + encodeURIComponent(orderId) +
    '&accountName=' + encodeURIComponent(accountName);

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Thanh toán Workshop</title>' +
    '<style>' +
    'body{font-family:sans-serif;text-align:center;padding:32px 16px;background:#f5f3ff;margin:0}' +
    '.card{background:#fff;border-radius:20px;padding:32px 24px;max-width:420px;margin:0 auto;box-shadow:0 8px 32px rgba(108,71,255,.12)}' +
    'h2{color:#6c47ff;margin-bottom:8px;font-size:1.4rem}' +
    '.order-id{background:#ede9ff;color:#4b2fe0;padding:8px 16px;border-radius:8px;font-weight:700;font-size:1rem;display:inline-block;margin:12px 0}' +
    'img{width:240px;height:240px;border-radius:12px;margin:16px 0}' +
    '.info{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:12px;margin:16px 0;font-size:.9rem;color:#166534}' +
    '.note{color:#6b7280;font-size:.82rem;margin-top:16px;line-height:1.6}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div style="font-size:2.5rem">🎉</div>' +
    '<h2>Đăng ký thành công!</h2>' +
    '<p style="color:#6b7280;font-size:.9rem">Quét mã QR để thanh toán học phí</p>' +
    '<div class="order-id">Mã đơn: ' + orderId + '</div>' +
    '<img src="' + qrUrl + '" alt="QR thanh toán" />' +
    '<div class="info">' +
    '🏦 MBBank · ' + accountNo + '<br>' +
    '👤 ' + accountName + '<br>' +
    '💰 1.860.000đ<br>' +
    '📝 Nội dung: <strong>' + orderId + '</strong>' +
    '</div>' +
    '<div class="note">⚠️ Vui lòng nhập đúng mã đơn vào nội dung chuyển khoản<br>để hệ thống tự động xác nhận thanh toán của bạn.</div>' +
    '</div></body></html>'
  );
}

// ── Xử lý đăng ký qua JSON (dự phòng) ──
function handleRegistrationJson(data) {
  const sheet   = getOrCreateSheet();
  const orderId = 'WS-' + Date.now(); // mã đơn hàng duy nhất

  // Lưu vào sheet (cột Thanh toán để trống, sẽ cập nhật khi webhook về)
  sheet.appendRow([
    new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }),
    data.name  || '',
    data.email || '',
    data.phone || '',
    data.job   || '',
    orderId,
    'Chờ thanh toán',
  ]);

  // Tạo đơn hàng SePay
  const checkoutUrl = createSePayOrder(orderId, data);

  return jsonResponse({ result: 'success', checkoutUrl });
}

// ── Gọi SePay API tạo đơn hàng ──
function createSePayOrder(orderId, data) {
  const { merchantId, secretKey } = getCreds();

  const payload = {
    merchantId:         merchantId,
    secretKey:          secretKey,
    orderInvoiceNumber: orderId,
    orderAmount:        1860000,
    currency:           'VND',
    operation:          'PURCHASE',
    orderDescription:   'Workshop Tu Y Tuong Thanh Website - ' + (data.name || ''),
    customerName:       data.name  || '',
    customerEmail:      data.email || '',
    customerPhone:      data.phone || '',
    successUrl:         'https://workshop.example.com/thank-you', // đổi thành URL cảm ơn của bạn
    cancelUrl:          'https://workshop.example.com/#register',
    errorUrl:           'https://workshop.example.com/#register',
  };

  const response = UrlFetchApp.fetch(SEPAY_API, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const raw = response.getContentText();
  Logger.log('SePay response: ' + raw);

  try {
    const result = JSON.parse(raw);
    if (result && result.checkoutUrl) return result.checkoutUrl;
    if (result && result.data && result.data.checkoutUrl) return result.data.checkoutUrl;
  } catch (e) {
    Logger.log('SePay parse error: ' + e.toString());
  }

  return null;
}

// ── Xử lý webhook từ SePay khi khách thanh toán xong ──
// SePay bank monitoring webhook gửi nội dung chuyển khoản trong trường "content"
// Format: { content: "WS-1234567890", transferAmount: 10000, transferType: "in", ... }
function handleSePayWebhook(data) {
  // Thử tất cả các trường có thể chứa mã đơn
  const raw =
    (data.content)                          ||  // bank monitoring webhook
    (data.code)                             ||  // một số phiên bản SePay
    (data.order && data.order.orderInvoiceNumber) || // checkout IPN
    '';

  // Tìm chuỗi "WS-<timestamp>" trong nội dung chuyển khoản
  const match = raw.toString().match(/WS-\d+/);
  const orderId = match ? match[0] : null;

  if (orderId) {
    updatePaymentStatus(orderId, 'Đã thanh toán ✅');
  }

  Logger.log('Webhook received: ' + JSON.stringify(data));
  Logger.log('orderId extracted: ' + orderId);

  return jsonResponse({ result: 'success' });
}

// ── Tìm đơn theo orderId và cập nhật trạng thái thanh toán ──
function updatePaymentStatus(orderId, status) {
  const sheet     = getOrCreateSheet();
  const lastRow   = sheet.getLastRow();
  const orderCol  = 6; // cột F = mã đơn hàng
  const statusCol = 7; // cột G = trạng thái

  for (let i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, orderCol).getValue() === orderId) {
      sheet.getRange(i, statusCol).setValue(status);
      break;
    }
  }
}

// ── Tạo / lấy sheet, thêm header nếu chưa có ──
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Thời gian', 'Họ và tên', 'Email', 'Số điện thoại', 'Nghề nghiệp', 'Mã đơn', 'Thanh toán']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#6c47ff').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 7, 160);
  }

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Test thủ công: chạy hàm này để kiểm tra ghi sheet ──
function testWrite() {
  const sheet = getOrCreateSheet();
  sheet.appendRow([
    new Date().toLocaleString('vi-VN'),
    'Nguyễn Test', 'test@email.com', '0900000000', 'Kinh doanh',
    'WS-TEST-001', 'Chờ thanh toán',
  ]);
  Logger.log('✅ Ghi thành công!');
}
