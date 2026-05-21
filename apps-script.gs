// Workshop "Từ Ý Tưởng Thành Website" – Google Apps Script
// Dán toàn bộ code này vào Apps Script Editor, sau đó Deploy lại → Web App

const SHEET_NAME = 'Đăng ký';
const SEPAY_API  = 'https://pay.sepay.vn/v1/checkout/init';

function getCreds() {
  const props = PropertiesService.getScriptProperties();
  return {
    merchantId: props.getProperty('SEPAY_MERCHANT_ID'),
    secretKey:  props.getProperty('SEPAY_SECRET_KEY'),
  };
}

function doGet(e) {
  const p = e.parameter;

  // ── CHECK PAYMENT STATUS (landing page polling) ──
  if (p.action === 'check') {
    const code = p.code || '';
    if (!code) return jsonResponse({paid: false});
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === code) {
        return jsonResponse({paid: data[i][6] === 'Đã thanh toán ✅'});
      }
    }
    return jsonResponse({paid: false});
  }

  // ── REGISTER ──
  if (p && p.name && p.name.trim() && p.email && p.email.trim() && p.phone && p.phone.trim()) {
    return handleFormSubmit(p);
  }

  return jsonResponse({ok: true});
}

function doPost(e) {
  try {
    if (e.postData.type === 'application/x-www-form-urlencoded') {
      return handleFormSubmit(e.parameter);
    }

    const data = JSON.parse(e.postData.contents);

    if (data.transferType || data.content || data.notification_type === 'ORDER_PAID') {
      return handleSePayWebhook(data);
    }
    return handleRegistrationJson(data);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ result: 'error', message: err.toString() });
  }
}

// ── Xử lý form submit: lưu sheet + trả JSON có code ──
function handleFormSubmit(params) {
  const orderId = 'WS' + Date.now();
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

  return jsonResponse({ success: true, code: orderId });
}

function handleRegistrationJson(data) {
  const sheet   = getOrCreateSheet();
  const orderId = 'WS' + Date.now();

  sheet.appendRow([
    new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }),
    data.name  || '',
    data.email || '',
    data.phone || '',
    data.job   || '',
    orderId,
    'Chờ thanh toán',
  ]);

  return jsonResponse({ result: 'success' });
}

function handleSePayWebhook(data) {
  const raw =
    (data.content) ||
    (data.code)    ||
    (data.order && data.order.orderInvoiceNumber) ||
    '';

  const match = raw.toString().match(/WS\d+/);
  const orderId = match ? match[0] : null;

  if (orderId) {
    updatePaymentStatus(orderId, 'Đã thanh toán ✅');
  }

  Logger.log('Webhook received: ' + JSON.stringify(data));
  Logger.log('orderId extracted: ' + orderId);

  return jsonResponse({ result: 'success' });
}

function updatePaymentStatus(orderId, status) {
  const sheet     = getOrCreateSheet();
  const lastRow   = sheet.getLastRow();
  const orderCol  = 6;
  const statusCol = 7;

  for (let i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, orderCol).getValue() === orderId) {
      sheet.getRange(i, statusCol).setValue(status);
      break;
    }
  }
}

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

function testWrite() {
  const sheet = getOrCreateSheet();
  sheet.appendRow([
    new Date().toLocaleString('vi-VN'),
    'Nguyễn Test', 'test@email.com', '0900000000', 'Kinh doanh',
    'WS-TEST-001', 'Chờ thanh toán',
  ]);
  Logger.log('✅ Ghi thành công!');
}
