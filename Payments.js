/** ============================================================
 *  Payments.js — チケット注文の決済導線（Square／銀行振込）
 *  PD2026_EntryManager（Pathos Drive 2026 エントリー管理・JUNYAが過去に作成した
 *  デモモード付きリマインド/督促メールの仕組み）の設計を踏襲。
 *  対象シート：注文（個人チケット・団体予約 共通）／決済設定
 * ============================================================ */

const PAY_SHEET_NAME = { ORDERS: '注文', SETTINGS: '決済設定' };

// 注文シート列定義（1-based）
const COL_ORDER = {
  NO: 1, TIMESTAMP: 2, TYPE: 3, NAME: 4, CONTACT: 5, BREAKDOWN: 6,
  TOTAL: 7, METHOD: 8, DEADLINE: 9, STATUS: 10,
  CONFIRM_SENT: 11, REMINDER_SENT: 12, OVERDUE_SENT: 13, NOTES: 14, QTY: 15
};
const ORDER_LAST_COL = COL_ORDER.QTY;

// 決済設定シートの行定義（1-based）— A列ラベル・B列値
const PAY_SETTINGS_ROW = {
  HEADER: 1, DEMO_MODE: 2, TEST_EMAIL: 3, DEADLINE_DAYS: 4, REMINDER_DAY: 5,
  ADMIN_EMAIL: 6, EVENT_NAME: 7, BANK_NAME: 8, BANK_BRANCH: 9,
  BANK_TYPE: 10, BANK_NUMBER: 11, BANK_HOLDER: 12
};

function getPaySettings_(){
  var sh = soloSs_().getSheetByName(PAY_SHEET_NAME.SETTINGS);
  var data = sh.getRange(PAY_SETTINGS_ROW.DEMO_MODE, 2, 11, 1).getValues();
  return {
    demoMode:     data[0][0]===true || String(data[0][0]).toUpperCase()==='TRUE',
    testEmail:    String(data[1][0]||'').trim(),
    deadlineDays: Number(data[2][0])||7,
    reminderDay:  Number(data[3][0])||3,
    adminEmail:   String(data[4][0]||'').trim(),
    eventName:    String(data[5][0]||'HARIBOW単独公演').trim(),
    bankName:     String(data[6][0]||'').trim(),
    bankBranch:   String(data[7][0]||'').trim(),
    bankType:     String(data[8][0]||'').trim(),
    bankNumber:   String(data[9][0]||'').trim(),
    bankHolder:   String(data[10][0]||'').trim()
  };
}

function payRecipientEmail_(realEmail, settings){
  if (settings.demoMode && settings.testEmail) return settings.testEmail;
  return realEmail;
}

function calcPayDeadline_(orderDate, settings){
  var d = new Date(orderDate.getTime());
  d.setDate(d.getDate() + settings.deadlineDays);
  return d;
}

function formatPayDate_(date){
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

/** ------------------------------------------------------------
 *  doPost 受け口：survey==='ticket_order'
 *  payload: { orderType:'individual'|'group', name, contact, items:[{label,qty,price}],
 *             total, method:'square'|'bank' }
 *  戻り値: { ok:true, method:'bank' } または { ok:true, method:'square', url } または
 *          { ok:false, error:'square_not_configured' }
 * ------------------------------------------------------------ */
function submitTicketOrder(payload){
  var name = String(payload.name||'').trim();
  var contact = String(payload.contact||'').trim();
  var items = payload.items||[];
  var total = Number(payload.total)||0;
  var method = payload.method==='square' ? 'square' : 'bank';
  var orderType = payload.orderType==='group' ? '団体' : '個人';

  if (!name || !contact || !items.length || total<=0){
    return { ok:false, error:'invalid_order' };
  }

  var breakdown = items.map(function(it){
    return it.label + '×' + it.qty + '(¥' + Number(it.price).toLocaleString() + ')';
  }).join('、');
  var qty = sumTicketQty_(items);

  var settings = getPaySettings_();
  var now = new Date();

  if (method==='square'){
    var link = createSquarePaymentLink_(name, breakdown, total);
    if (!link){
      return { ok:false, error:'square_not_configured' };
    }
    writeOrderRow_({
      type: orderType, name: name, contact: contact, breakdown: breakdown,
      total: total, method: 'Square', deadline: '', status: '未',
      timestamp: now, confirmSent: now, qty: qty
    });
    return { ok:true, method:'square', url: link };
  }

  // 銀行振込
  var deadline = calcPayDeadline_(now, settings);
  var rowNum = writeOrderRow_({
    type: orderType, name: name, contact: contact, breakdown: breakdown,
    total: total, method: '振込', deadline: deadline, status: '未',
    timestamp: now, confirmSent: null, qty: qty
  });
  sendBankConfirmation_({
    rowNum: rowNum, name: name, contact: contact, breakdown: breakdown,
    total: total, deadline: deadline
  }, settings);

  return { ok:true, method:'bank', deadline: formatPayDate_(deadline) };
}

function writeOrderRow_(o){
  var sh = soloSs_().getSheetByName(PAY_SHEET_NAME.ORDERS);
  var nextRow = sh.getLastRow()+1;
  var no = nextRow-1;
  var row = new Array(ORDER_LAST_COL).fill('');
  row[COL_ORDER.NO-1]        = no;
  row[COL_ORDER.TIMESTAMP-1] = o.timestamp;
  row[COL_ORDER.TYPE-1]      = o.type;
  row[COL_ORDER.NAME-1]      = o.name;
  row[COL_ORDER.CONTACT-1]   = o.contact;
  row[COL_ORDER.BREAKDOWN-1] = o.breakdown;
  row[COL_ORDER.TOTAL-1]     = o.total;
  row[COL_ORDER.METHOD-1]    = o.method;
  row[COL_ORDER.DEADLINE-1]  = o.deadline||'';
  row[COL_ORDER.STATUS-1]    = o.status;
  row[COL_ORDER.CONFIRM_SENT-1] = o.confirmSent||'';
  row[COL_ORDER.QTY-1]       = o.qty||0;
  sh.getRange(nextRow,1,1,ORDER_LAST_COL).setValues([row]);
  return nextRow;
}

/** 応援金(ドネーション)を除いた、実際のチケット枚数の合計を返す */
function sumTicketQty_(items){
  return items.reduce(function(sum, it){
    return it.label==='応援金' ? sum : sum + Number(it.qty||0);
  }, 0);
}

/** ダッシュボード「集客・収支」タブ向け：見込み(集客収支シミュレーター)＋実績(注文シート)のサマリーを返す */
function getBusinessSummary_(){
  var ss = soloSs_();
  var sim = ss.getSheetByName('集客収支シミュレーター');
  var forecast = {
    totalAttendance: sim.getRange('B33').getValue(),
    totalIncome:     sim.getRange('B44').getValue(),
    totalExpense:    sim.getRange('B47').getValue(),
    profitLoss:      sim.getRange('B48').getValue()
  };

  var orderSheet = ss.getSheetByName(PAY_SHEET_NAME.ORDERS);
  var lastRow = orderSheet.getLastRow();
  var actual = { orderCount: 0, ticketQty: 0, revenueTotal: 0, revenueConfirmed: 0 };
  if (lastRow > 1){
    var data = orderSheet.getRange(2, 1, lastRow - 1, ORDER_LAST_COL).getValues();
    data.forEach(function(r){
      if (String(r[COL_ORDER.NAME - 1] || '').trim() === '') return;
      actual.orderCount++;
      actual.ticketQty += Number(r[COL_ORDER.QTY - 1]) || 0;
      var total = Number(r[COL_ORDER.TOTAL - 1]) || 0;
      actual.revenueTotal += total;
      if (r[COL_ORDER.STATUS - 1] === '済') actual.revenueConfirmed += total;
    });
  }

  return {
    simulatorUrl: 'https://docs.google.com/spreadsheets/d/' + SOLO_SS_ID + '/edit#gid=1025959564',
    forecast: forecast,
    actual: actual
  };
}

/** ------------------------------------------------------------
 *  Square Payment Link 作成（動的・注文ごと）
 *  SQUARE_ACCESS_TOKEN/SQUARE_LOCATION_ID が未設定ならnullを返す。
 *  JPYは補助単位が無いため amount はそのまま円額（×100しない）。
 * ------------------------------------------------------------ */
function createSquarePaymentLink_(buyerName, breakdown, total){
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SQUARE_ACCESS_TOKEN');
  var locationId = props.getProperty('SQUARE_LOCATION_ID');
  if (!token || !locationId) return null;

  var body = {
    idempotency_key: Utilities.getUuid(),
    order: {
      location_id: locationId,
      line_items: [{
        name: 'HARIBOW単独公演チケット（' + buyerName + '様・' + breakdown + '）',
        quantity: '1',
        base_price_money: { amount: total, currency: 'JPY' }
      }]
    },
    checkout_options: {
      redirect_url: 'https://junyaunitydev.github.io/haribow-solo-show-survey/tickets.html'
    }
  };

  var res = UrlFetchApp.fetch('https://connect.squareup.com/v2/online-checkout/payment-links', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token, 'Square-Version': '2025-01-23' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var json = JSON.parse(res.getContentText());
  if (res.getResponseCode() >= 300 || !json.payment_link){
    Logger.log('Square決済リンク作成失敗: ' + res.getContentText());
    return null;
  }
  return json.payment_link.url;
}

/** ------------------------------------------------------------
 *  銀行振込：確認メール
 * ------------------------------------------------------------ */
function sendBankConfirmation_(order, settings){
  var to = payRecipientEmail_(order.contact, settings);
  var subject = '【' + settings.eventName + '】ご予約確認・お振込みのご案内';
  var deadlineStr = formatPayDate_(order.deadline);

  var body = order.name + ' 様\n\n';
  body += 'この度は ' + settings.eventName + ' へのご予約、誠にありがとうございます。\n';
  body += '以下の内容で受け付けました。\n\n';
  body += '────────────────────\n';
  body += '【ご注文内容】 ' + order.breakdown + '\n';
  body += '【合計金額】 ¥' + Number(order.total).toLocaleString() + '\n';
  body += '【お振込み期限】 ' + deadlineStr + '\n';
  body += '────────────────────\n\n';
  body += '【振込先】\n';
  body += settings.bankName + ' ' + settings.bankBranch + '\n';
  body += settings.bankType + ' ' + settings.bankNumber + '\n';
  body += settings.bankHolder + '\n\n';
  body += '恐れ入りますが、振込手数料はご負担をお願いいたします。\n';
  body += '期限までにお振込みください。\n\n';
  body += '※ このメールは ' + settings.eventName + ' 予約システムから自動送信されています。\n';
  body += 'ご不明な点は本メールへの返信、または info@haribow.com までお問い合わせください。\n\n';
  body += '---\n';
  body += settings.eventName + ' 運営事務局';

  if (settings.demoMode){
    body = '[DEMO MODE] 本来の宛先: ' + order.contact + '\n─────────────────────────────\n\n' + body;
  }

  GmailApp.sendEmail(to, subject, body);

  var sh = soloSs_().getSheetByName(PAY_SHEET_NAME.ORDERS);
  sh.getRange(order.rowNum, COL_ORDER.CONFIRM_SENT).setValue(new Date());
}

/** ------------------------------------------------------------
 *  日次トリガー：リマインド（期限前）／督促（期限超過）
 * ------------------------------------------------------------ */
function checkPaymentDeadlines(){
  var settings = getPaySettings_();
  var sh = soloSs_().getSheetByName(PAY_SHEET_NAME.ORDERS);
  var lastRow = sh.getLastRow();
  if (lastRow<=1) { Logger.log('checkPaymentDeadlines: 注文なし'); return; }

  var now = new Date();
  var data = sh.getRange(2,1,lastRow-1,ORDER_LAST_COL).getValues();
  var overdueForAdmin = [];

  data.forEach(function(row, i){
    var rowNum = i+2;
    var method = row[COL_ORDER.METHOD-1];
    var status = row[COL_ORDER.STATUS-1];
    var deadlineVal = row[COL_ORDER.DEADLINE-1];
    if (method!=='振込' || status==='済' || !deadlineVal) return;

    var deadline = (deadlineVal instanceof Date) ? deadlineVal : new Date(deadlineVal);
    var timestamp = row[COL_ORDER.TIMESTAMP-1];
    var orderDate = (timestamp instanceof Date) ? timestamp : new Date(timestamp);
    var reminderSent = row[COL_ORDER.REMINDER_SENT-1];
    var overdueSent = row[COL_ORDER.OVERDUE_SENT-1];

    var order = {
      rowNum: rowNum, name: row[COL_ORDER.NAME-1], contact: row[COL_ORDER.CONTACT-1],
      breakdown: row[COL_ORDER.BREAKDOWN-1], total: row[COL_ORDER.TOTAL-1], deadline: deadline
    };

    // リマインド（指定日数経過・未送信）
    var reminderDate = new Date(orderDate.getTime());
    reminderDate.setDate(reminderDate.getDate() + settings.reminderDay);
    if (!reminderSent && now >= reminderDate && now <= deadline){
      sendBankReminder_(order, settings);
      sh.getRange(rowNum, COL_ORDER.REMINDER_SENT).setValue(new Date());
    }

    // 督促（期限超過・未送信）
    if (!overdueSent && now > deadline){
      sendBankOverdue_(order, settings);
      sh.getRange(rowNum, COL_ORDER.OVERDUE_SENT).setValue(new Date());
      overdueForAdmin.push(order);
    }
  });

  if (overdueForAdmin.length){
    sendOverdueAdminSummary_(overdueForAdmin, settings);
  }
  Logger.log('checkPaymentDeadlines: 督促' + overdueForAdmin.length + '件');
}

function sendBankReminder_(order, settings){
  var to = payRecipientEmail_(order.contact, settings);
  var subject = '【' + settings.eventName + '】お振込み期限のリマインド';
  var body = order.name + ' 様\n\n';
  body += settings.eventName + ' チケット代金のお振込み期限が近づいております。\n\n';
  body += '────────────────────\n';
  body += '【ご注文内容】 ' + order.breakdown + '\n';
  body += '【合計金額】 ¥' + Number(order.total).toLocaleString() + '\n';
  body += '【お振込み期限】 ' + formatPayDate_(order.deadline) + '\n';
  body += '────────────────────\n\n';
  body += '振込先は受付時のメールをご確認ください。\n';
  body += '既にお振込み済みの場合、行き違いの可能性があります。本メールにご返信ください。\n\n';
  body += '---\n' + settings.eventName + ' 運営事務局';
  if (settings.demoMode){
    body = '[DEMO MODE] 本来の宛先: ' + order.contact + '\n─────────────────────────────\n\n' + body;
  }
  GmailApp.sendEmail(to, subject, body);
}

function sendBankOverdue_(order, settings){
  var to = payRecipientEmail_(order.contact, settings);
  var subject = '【' + settings.eventName + '】お振込み期限超過のお知らせ';
  var body = order.name + ' 様\n\n';
  body += settings.eventName + ' チケット代金のお振込み期限を超過しております。\n\n';
  body += '────────────────────\n';
  body += '【ご注文内容】 ' + order.breakdown + '\n';
  body += '【合計金額】 ¥' + Number(order.total).toLocaleString() + '\n';
  body += '【お振込み期限】 ' + formatPayDate_(order.deadline) + '（超過）\n';
  body += '────────────────────\n\n';
  body += 'お早めにお振込みください。\n';
  body += '既にお振込み済みの場合、行き違いの可能性があります。本メールにご返信ください。\n\n';
  body += '---\n' + settings.eventName + ' 運営事務局';
  if (settings.demoMode){
    body = '[DEMO MODE] 本来の宛先: ' + order.contact + '\n─────────────────────────────\n\n' + body;
  }
  GmailApp.sendEmail(to, subject, body);
}

function sendOverdueAdminSummary_(orders, settings){
  var adminTo = settings.demoMode && settings.testEmail ? settings.testEmail : settings.adminEmail;
  if (!adminTo){ Logger.log('運営メアド未設定のためサマリ送信をスキップ'); return; }

  var subject = '【' + settings.eventName + '】振込期限超過アラート（' + orders.length + '件）';
  var body = 'チケット代金の振込期限を超過している注文があります。\n\n────────────────────\n';
  orders.forEach(function(o, idx){
    body += (idx+1) + '. ' + o.name + '（' + o.contact + '）\n';
    body += '   内容: ' + o.breakdown + '\n';
    body += '   金額: ¥' + Number(o.total).toLocaleString() + '\n';
    body += '   期限: ' + formatPayDate_(o.deadline) + '\n\n';
  });
  body += '────────────────────\n入金確認後は「注文」シートの支払状態を「済」に変更してください。';
  if (settings.demoMode){
    body = '[DEMO MODE] 本来の宛先: ' + settings.adminEmail + '\n─────────────────────────────\n\n' + body;
  }
  GmailApp.sendEmail(adminTo, subject, body);
}

/** ------------------------------------------------------------
 *  手動テスト送信：既存の注文シートに一切触れず、ダミーデータで
 *  リマインド・督促・運営サマリの3通を送る（GASエディタから実行）
 * ------------------------------------------------------------ */
function testSendPaymentAlerts(){
  var settings = getPaySettings_();
  var mockOrder = {
    rowNum: -1, name: '【テスト送信】ダミー花子', contact: 'dummy@example.com',
    breakdown: '一般×2(¥7,000)', total: 7000,
    deadline: new Date(new Date().getTime() - 60*60*1000)
  };
  sendBankReminder_(mockOrder, settings);
  sendBankOverdue_(mockOrder, settings);
  sendOverdueAdminSummary_([mockOrder], settings);
  Logger.log('testSendPaymentAlerts: 3通送信完了（リマインド/督促/運営サマリ）。注文シートには一切影響していません。');
}

/** ------------------------------------------------------------
 *  トリガー登録（1回だけ手動実行）
 * ------------------------------------------------------------ */
function setupPaymentTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction()==='checkPaymentDeadlines') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkPaymentDeadlines').timeBased().everyHours(24).create();
  Logger.log('checkPaymentDeadlines を毎日実行するトリガーを登録しました。');
}
