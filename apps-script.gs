// Workshop "Từ Ý Tưởng Thành Website" – Google Apps Script
// Bản cập nhật sửa lỗi Webhook SePay & Đồng bộ dữ liệu

const SHEET_NAME = 'Đăng ký';

function doGet(e) {
  const p = e.parameter;

  // ── CHECK PAYMENT STATUS ──
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
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const data = JSON.parse(e.postData.contents);

    if (data.gateway_content || data.transaction_content || data.content || data.notification_type === 'ORDER_PAID') {
      return handleSePayWebhook(data);
    }
    return handleRegistrationJson(data);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ result: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function handleFormSubmit(params) {
  const orderId = 'WS' + Math.floor(100000 + Math.random() * 900000);
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
  const orderId = 'WS' + Math.floor(100000 + Math.random() * 900000);

  sheet.appendRow([
    new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }),
    data.name  || '',
    data.email || '',
    data.phone || '',
    data.job   || '',
    orderId,
    'Chờ thanh toán',
  ]);

  return jsonResponse({ result: 'success', code: orderId });
}

function handleSePayWebhook(data) {
  const rawContent =
    data.gateway_content ||
    data.transaction_content ||
    data.content ||
    (data.order && data.order.orderInvoiceNumber) ||
    '';

  Logger.log('Nội dung nhận từ SePay: ' + rawContent);

  const match   = rawContent.toString().match(/WS\d{6}/i);
  const orderId = match ? match[0].toUpperCase() : null;

  if (orderId) {
    const isUpdated = updatePaymentStatus(orderId, 'Đã thanh toán ✅');
    if (isUpdated) {
      Logger.log('Cập nhật thành công: ' + orderId);
      return jsonResponse({ result: 'success' });
    } else {
      Logger.log('Không tìm thấy mã đơn: ' + orderId);
      return jsonResponse({ result: 'error', message: 'Mã đơn không tồn tại' });
    }
  }

  Logger.log('Không tìm thấy mã đơn trong: ' + rawContent);
  return jsonResponse({ result: 'error', message: 'Không tìm thấy mã đơn' });
}

function updatePaymentStatus(orderId, status) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][5] && data[i][5].toString().trim().toUpperCase() === orderId.trim().toUpperCase()) {
      sheet.getRange(i + 1, 7).setValue(status);
      return true;
    }
  }
  return false;
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
