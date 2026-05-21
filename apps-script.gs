// Workshop "Từ Ý Tưởng Thành Website" – Google Apps Script

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
        break;
      }
    }

    return json({status: 'ok'});
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return json({error: err.toString()});
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testWrite() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([new Date(), 'Nguyễn Test', '0900000000', 'test@email.com', 'Kinh doanh', 'WSTEST01', 'Chờ thanh toán', '', '']);
  Logger.log('✅ Ghi thành công!');
}
