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
  if (e.parameter && e.parameter.name) {
    return handleFormSubmit(e.parameter);
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
    if (data.notification_type === 'ORDER_PAID') {
      return handleSePayWebhook(data);
    }
    return handleRegistrationJson(data);

  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h3>Có lỗi xảy ra: ' + err.toString() + '</h3>'
    );
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

  const checkoutUrl = createSePayOrder(orderId, params);

  if (checkoutUrl) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head>' +
      '<meta http-equiv="refresh" content="0;url=' + checkoutUrl + '">' +
      '</head><body>Đang chuyển sang trang thanh toán...</body></html>'
    );
  }

  // Fallback nếu SePay chưa trả URL (môi trường test chưa kích hoạt)
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><body>' +
    '<h2>✅ Đăng ký thành công!</h2>' +
    '<p>Mã đơn: <strong>' + orderId + '</strong></p>' +
    '<p>Chúng tôi sẽ liên hệ xác nhận thanh toán trong 24 giờ.</p>' +
    '</body></html>'
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

  const result = JSON.parse(response.getContentText());

  // SePay trả về URL trang thanh toán
  if (result && result.checkoutUrl) return result.checkoutUrl;
  if (result && result.data && result.data.checkoutUrl) return result.data.checkoutUrl;

  // Fallback nếu API chưa rõ response format
  Logger.log('SePay response: ' + response.getContentText());
  return null;
}

// ── Xử lý webhook từ SePay khi khách thanh toán xong ──
function handleSePayWebhook(data) {
  const orderId = data.order && data.order.orderInvoiceNumber;
  if (orderId) {
    updatePaymentStatus(orderId, 'Đã thanh toán ✅');
  }
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
