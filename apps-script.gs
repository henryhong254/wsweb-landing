// Workshop "Từ Ý Tưởng Thành Website" – Google Apps Script

var TELEGRAM_TOKEN   = '8427417613:AAEFwBt2TCpTqzi16QEFh3piLkmcV2WTp-Q';
var TELEGRAM_CHAT_ID = '1086190321';
var RESEND_API_KEY   = 're_7deuF8gG_6ZpmDfMqjsq1SXCGEMiQxkbi';

function doGet(e) {
  var action = e.parameter.action;

  // ── CHECK PAYMENT STATUS (landing page polling) ──
  if (action === 'check') {
    var code = e.parameter.code || '';
    if (!code) return json({paid: false});

    var sheet = SpreadsheetApp.getActiveSheet();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][5] === code) {
        return json({paid: data[i][6] === 'Đã thanh toán'});
      }
    }
    return json({paid: false});
  }

  // ── REGISTER ──
  var name  = (e.parameter.name  || '').trim();
  var phone = (e.parameter.phone || '').trim();
  var email = (e.parameter.email || '').trim();
  var job   = (e.parameter.job   || '').trim();

  if (!name || !phone || !email) {
    return json({error: 'missing fields'});
  }

  var code  = 'WS' + Math.random().toString(36).substring(2, 8).toUpperCase();
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([new Date(), name, phone, email, job, code, 'Chờ thanh toán', '', '']);

  sendTelegram('🔔 Đăng ký mới!\n👤 ' + name + '\n📱 ' + phone + '\n📧 ' + email + '\n🔖 ' + code);

  return json({success: true, code: code});
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var desc   = (body.description || body.content || '').toUpperCase();
    var amount = body.transferAmount || body.amount || 0;
    var time   = body.transactionDate || new Date().toISOString();

    var match = desc.match(/WS[A-Z0-9]{6}/);
    if (!match) return json({status: 'no match'});

    var code  = match[0];
    var sheet = SpreadsheetApp.getActiveSheet();
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][5] === code) {
        sheet.getRange(i + 1, 7).setValue('Đã thanh toán');
        sheet.getRange(i + 1, 8).setValue(time);
        sheet.getRange(i + 1, 9).setValue(amount);

        var name  = data[i][1];
        var email = data[i][3];

        sendTelegram('✅ Thanh toán mới!\n👤 ' + name + '\n📧 ' + email + '\n🔖 ' + code + '\n💰 ' + amount + 'đ');

        if (email) {
          sendConfirmEmail(String(email).trim(), String(name).trim(), String(code));
        }
        break;
      }
    }

    return json({status: 'ok'});
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return json({error: err.toString()});
  }
}

function sendConfirmEmail(toEmail, toName, orderCode) {
  var html =
    '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.8;color:#222;max-width:600px;">' +
    '<p>Chào <strong>' + toName + '</strong>,</p>' +
    '<p>Chúc mừng bạn đã đăng ký thành công <strong>Workshop Từ Ý Tưởng Thành Website!</strong></p>' +
    '<p>Chúng tôi đã nhận được thanh toán của bạn và xác nhận suất tham gia.</p>' +
    '<hr style="border:none;border-top:1px solid #eee;margin:24px 0;">' +
    '<p><strong>LỊCH TRÌNH WORKSHOP</strong></p>' +
    '<p>📅 3 – 5 tháng 7, 2025<br>' +
    '💻 Online + Offline tại TP.HCM</p>' +
    '<p>Thông tin chi tiết sẽ được gửi qua email và Group Zalo trước ngày khai giảng.</p>' +
    '<hr style="border:none;border-top:1px solid #eee;margin:24px 0;">' +
    '<p>Hẹn gặp bạn tại Workshop! 🎓</p>' +
    '<p>Trân trọng,<br><strong>Hồng Thiên Ý</strong><br>Sáng Lập Founder\'s North</p>' +
    '</div>';

  try {
    var response = UrlFetchApp.fetch('https://api.resend.com/emails', {
      method: 'post',
      contentType: 'application/json',
      headers: {'Authorization': 'Bearer ' + RESEND_API_KEY},
      payload: JSON.stringify({
        from: 'Founders North <workshop@mail.foundersnorth.com.vn>',
        to: [toEmail],
        subject: '[Xác nhận] Chúc mừng bạn đã đăng ký Workshop Từ Ý Tưởng Thành Website! 🎉',
        html: html
      }),
      muteHttpExceptions: true
    });
    sendTelegram('🟢 Email gửi: ' + response.getResponseCode());
  } catch(err) {
    sendTelegram('🔴 Lỗi email: ' + err.toString());
  }
}

function sendTelegram(msg) {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({chat_id: TELEGRAM_CHAT_ID, text: msg})
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testTelegram() {
  sendTelegram('✅ Test từ Workshop Apps Script – hoạt động rồi!');
}

function testWrite() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([new Date(), 'Nguyễn Test', '0900000000', 'test@email.com', 'Kinh doanh', 'WSTEST01', 'Chờ thanh toán', '', '']);
  Logger.log('✅ Ghi thành công!');
}
