/**
 * 音樂會訂票 - Google Apps Script 後端
 * 將此程式貼到「試算表 > 擴充功能 > Apps Script」，部署為網頁應用程式。
 *
 * 付款狀態流程：
 *   未匯款 →（使用者填寫後五碼）→ 待核對 →（主辦方核對銀行入帳後手動改）→ 已確認
 *   另有「已取消」可手動設定。
 */
const SHEET_NAME = '訂單';

// ===== 確認信設定（可自行修改）=====
const ORGANIZER = '聲動室內樂團 SOUNDANZE ENSEMBLE';  // 寄件者顯示名稱
const PAGE_URL = 'https://mikekuo2024-collab.github.io/concert-tickets/';  // 訂票網頁公開網址（確認信按鈕連到這裡）
const PAYMENT_INFO = '銀行：（807）永豐銀行 中壢分行（0243）\n戶名：聲動室內樂團\n帳號：02401800131556';  // 確認信中的匯款資訊（\n 換行）；留空則不顯示

const DISCOUNT = 0.7;
const PRICES = [600, 800, 1000, 1200];
const SESSION_IDS = [1, 2, 3];
const STATUS = ['未匯款', '待核對', '已確認', '已取消'];
const HEADERS = ['訂單編號', '訂購時間', '姓名', '電話', 'Email', '匯款後五碼',
  '觀賞場次', '場次一明細', '場次二明細', '場次三明細', '早鳥總金額', '樂團名稱',
  '付款狀態', '後五碼填寫時間', '核對備註', '查詢碼',
  '取票方式', 'OPENTIX會員姓名', 'OPENTIX會員電話', '自行選位', '負責老師'];
// 團購老師：代碼對應各老師頁面資料夾（例如 PAGE_URL + 'wang/'）
const TEACHERS = { wang: '王裕文', wei: '魏鴻達', hsieh: '謝政良' };
const TEACHER_IDS = {}; Object.keys(TEACHERS).forEach(k => TEACHER_IDS[TEACHERS[k]] = k);
// 依老師姓名取得訂單頁網址（無老師則為首頁）
function pageUrlFor_(teacherName) {
  const id = TEACHER_IDS[teacherName];
  return PAGE_URL + (id ? id + '/' : '');
}
const DELIVERY = ['電子票', '現場領票'];
// 欄位位置（1 起算）
const COL = { id: 1, time: 2, name: 3, phone: 4, email: 5, last5: 6, sessions: 7,
  detail1: 8, detail2: 9, detail3: 10, total: 11, band: 12, status: 13, last5Time: 14, note: 15, token: 16,
  delivery: 17, otName: 18, otPhone: 19, seatPick: 20, teacher: 21 };

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  try {
    // 唯讀或寄信的動作不鎖定，避免大量寄信時擋住訂票
    if (d.action === 'get') return json_(getOrder_(d));
    if (d.action === 'status') return json_({ ok: true, open: salesOpen_() });
    if (d.action === 'adminData') { checkAdmin_(d); return json_(adminData_()); }
    if (d.action === 'adminMail') { checkAdmin_(d); return json_(adminMail_(d)); }
    // 會寫入試算表的動作需要鎖定
    return json_(withLock_(() => {
      if (d.action === 'pay') return fillLast5_(d);
      if (d.action === 'adminSetStatus') { checkAdmin_(d); return adminSetStatus_(d); }
      if (d.action === 'adminSetSales') { checkAdmin_(d); return adminSetSales_(d); }
      return createOrder_(d);
    }));
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- 新增訂單 ---------- */
function createOrder_(d) {
  if (!salesOpen_()) throw new Error('本活動已停止接受訂票');
  if (!d.name || !d.phone || !d.email) throw new Error('必填欄位未填寫');
  if (DELIVERY.indexOf(d.delivery) < 0) throw new Error('請選擇取票方式');
  if (d.teacher && !TEACHERS[d.teacher]) throw new Error('團購老師代碼錯誤');
  const teacherName = d.teacher ? TEACHERS[d.teacher] : '';
  d.teacherName = teacherName;
  if (d.delivery === '電子票' && (!d.otName || !d.otPhone)) throw new Error('請填寫 OPENTIX 會員姓名與電話');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) throw new Error('Email 格式錯誤');
  if (d.last5 && !/^\d{5}$/.test(d.last5)) throw new Error('匯款後五碼格式錯誤');
  if (!Array.isArray(d.items) || !d.items.length) throw new Error('未選擇任何票券');

  // 後端重新計算金額（不信任前端傳來的總價）
  let orig = 0;
  d.items.forEach(it => {
    const p = Number(it.price), q = parseInt(it.qty, 10);
    if (PRICES.indexOf(p) < 0 || SESSION_IDS.indexOf(Number(it.session)) < 0 || !(q >= 1 && q <= 20)) {
      throw new Error('票券資料錯誤');
    }
    orig += p * q;
  });
  const total = Math.round(orig * DISCOUNT);

  const sheet = getSheet_();
  // 優先使用前端產生的訂單編號（格式 T + 12 碼數字 + 2 碼英文）
  const orderId = /^T\d{12}[A-Z]{2}$/.test(d.orderId) ? d.orderId :
    'T' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss') + 'ZZ';
  const now = new Date();
  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);  // 確認信連結用的查詢碼

  sheet.appendRow([
    orderId, now, d.name, "'" + d.phone, d.email, d.last5 ? "'" + d.last5 : '',
    (d.sessions || []).join('\n'), d.detail1 || '', d.detail2 || '', d.detail3 || '',
    total, d.band || '',
    d.last5 ? '待核對' : '未匯款', d.last5 ? now : '', '', token,
    d.delivery, d.delivery === '電子票' ? d.otName : '',
    d.delivery === '電子票' ? "'" + d.otPhone : '', d.seatPick ? '是' : '否', teacherName
  ]);
  // 寄送確認信；寄信失敗不影響訂單成立
  let emailSent = false;
  try {
    sendConfirmMail_(d, orderId, total, token);
    emailSent = true;
  } catch (err) {
    console.error('確認信寄送失敗：' + err.message);
  }
  return { ok: true, orderId: orderId, total: total, emailSent: emailSent };
}

/* ---------- 訂購確認信 ---------- */
function sendConfirmMail_(d, orderId, total, token) {
  const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = n => Number(n).toLocaleString('en-US');
  const td = 'padding:4px 8px;border-bottom:1px solid #eee;';

  let sessHtml = '';
  if (Array.isArray(d.view) && d.view.length) {
    d.view.forEach(v => {
      sessHtml += '<h3 style="margin:16px 0 4px;font-size:15px;">' + esc(v.title) + '</h3><ul style="margin:4px 0;">' +
        (v.lines || []).map(l => '<li>NT$' + fmt(l.price) + ' × ' + esc(l.qty) + ' 張　座位：' +
          esc((l.seats || []).join('、')) + '</li>').join('') +
        '</ul><div style="color:#555;">原價 NT$' + fmt(v.orig) + '　早鳥 7 折 NT$' + fmt(v.disc) + '</div>';
    });
  } else {
    sessHtml = [d.detail1, d.detail2, d.detail3].filter(Boolean).map(x => '<p>' + esc(x) + '</p>').join('');
  }

  const link = pageUrlFor_(d.teacherName) + '?order=' + encodeURIComponent(orderId) + '&t=' + token;
  const btnHtml =
    '<table cellpadding="0" cellspacing="0" style="margin:20px 0;"><tr><td style="background:#b4462a;border-radius:6px;">' +
    '<a href="' + esc(link) + '" style="display:inline-block;padding:12px 24px;color:#ffffff;text-decoration:none;font-weight:bold;">' +
    '查看訂單 / 填寫匯款資料</a></td></tr></table>' +
    '<p style="color:#888;font-size:12px;">若按鈕無法點擊，請複製此網址到瀏覽器：<br>' + esc(link) + '</p>';

  const deliveryHtml = d.delivery === '電子票'
    ? '<p>確認付款後，電子票將匯入您的 OPENTIX 會員票匣，請確認會員姓名與電話正確。</p>'
    : '<p>您選擇現場領票：請先完成付款並提供匯款帳號後五碼，演出當天可提早至現場領票。</p>';
  const payHtml = deliveryHtml + (d.last5
    ? '<p>已收到您的匯款帳號後五碼：<b>' + esc(d.last5) + '</b>，主辦方核對入帳後會再與您確認。</p>'
    : '<p>您尚未填寫匯款帳號後五碼。匯款後請點下方按鈕，填寫匯款帳號後五碼。</p>');

  const html =
    '<div style="font-family:sans-serif;font-size:14px;color:#222;max-width:600px;">' +
    '<p>' + esc(d.name) + ' 您好，感謝您訂購音樂會門票，以下是您的訂購內容：</p>' +
    '<table style="border-collapse:collapse;">' +
    '<tr><td style="' + td + 'color:#666;">訂單編號</td><td style="' + td + '"><b>' + esc(orderId) + '</b></td></tr>' +
    '<tr><td style="' + td + 'color:#666;">姓名</td><td style="' + td + '">' + esc(d.name) + '</td></tr>' +
    '<tr><td style="' + td + 'color:#666;">電話</td><td style="' + td + '">' + esc(d.phone) + '</td></tr>' +
    '<tr><td style="' + td + 'color:#666;">Email</td><td style="' + td + '">' + esc(d.email) + '</td></tr>' +
    '<tr><td style="' + td + 'color:#666;">樂團名稱</td><td style="' + td + '">' + esc(d.band || '（未填）') + '</td></tr>' +
    (d.teacherName ? '<tr><td style="' + td + 'color:#666;">團購老師</td><td style="' + td + '">' + esc(d.teacherName) + '老師</td></tr>' : '') +
    '<tr><td style="' + td + 'color:#666;">取票方式</td><td style="' + td + '">' +
      (d.delivery === '電子票' ? '電子票（OPENTIX 會員：' + esc(d.otName) + '／' + esc(d.otPhone) + '）' : esc(d.delivery)) + '</td></tr>' +
    '<tr><td style="' + td + 'color:#666;">選位</td><td style="' + td + '">' + (d.seatPick ? '自行指定（若已售出將代選鄰近位置）' : '由主辦方代選') + '</td></tr>' +
    '</table>' + sessHtml +
    '<p style="font-size:17px;margin-top:16px;">早鳥總金額：<b style="color:#b4462a;">NT$' + fmt(total) + '</b></p>' +
    (PAYMENT_INFO ? '<div style="border:2px solid #b4462a;border-radius:8px;padding:10px 14px;margin:16px 0;">' +
      '<b style="color:#b4462a;">匯款資訊</b><br>' + esc(PAYMENT_INFO).replace(/\n/g, '<br>') + '</div>' : '') +
    payHtml + btnHtml +
    '<p style="margin-top:20px;">' + esc(ORGANIZER) + ' 敬上</p>' +
    '<p style="color:#888;font-size:12px;">本信件由系統自動寄出，請勿直接回覆。</p></div>';

  MailApp.sendEmail({
    to: d.email,
    subject: '【訂購確認】聲動室內樂團 音樂會門票 ' + orderId,
    htmlBody: html,
    name: ORGANIZER
  });
}

/* 首次加入寄信功能時，請在編輯器選擇此函式按「執行」一次，完成寄信授權 */
function authorizeMail() {
  Logger.log('今日剩餘可寄信數：' + MailApp.getRemainingDailyQuota());
}

/* ---------- 補填匯款後五碼 ---------- */
function fillLast5_(d) {
  if (!/^\d{5}$/.test(d.last5)) throw new Error('匯款後五碼格式錯誤');
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const digits = s => String(s).replace(/\D/g, '');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.id - 1]).toUpperCase() !== String(d.orderId).toUpperCase()) continue;
    const byToken = d.token && String(rows[i][COL.token - 1]) === String(d.token);
    if (!byToken && digits(rows[i][COL.phone - 1]) !== digits(d.phone)) throw new Error('訂單編號與電話不符');
    const status = rows[i][COL.status - 1];
    if (status === '已確認') throw new Error('此訂單已確認付款，如需修改請聯絡主辦方');
    if (status === '已取消') throw new Error('此訂單已取消');
    const r = i + 1;
    sheet.getRange(r, COL.last5).setValue("'" + d.last5);
    sheet.getRange(r, COL.status).setValue('待核對');
    sheet.getRange(r, COL.last5Time).setValue(new Date());
    return { ok: true };
  }
  throw new Error('查無此訂單編號');
}

/* ---------- 依訂單編號 + 查詢碼取得訂單（唯讀） ---------- */
function getOrder_(d) {
  if (!d.orderId || !d.token) throw new Error('連結不完整');
  const rows = getSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[COL.id - 1]).toUpperCase() !== String(d.orderId).toUpperCase()) continue;
    if (!row[COL.token - 1] || String(row[COL.token - 1]) !== String(d.token)) break;
    const c = k => row[COL[k] - 1];
    const fmtTime = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : String(v || '');
    return { ok: true, order: {
      orderId: String(c('id')), time: fmtTime(c('time')), name: String(c('name')),
      phone: String(c('phone')), email: String(c('email')), band: String(c('band')),
      details: [c('detail1'), c('detail2'), c('detail3')].map(String).filter(x => x),
      sessions: String(c('sessions')), total: Number(c('total')),
      last5: String(c('last5') || ''), status: String(c('status') || ''),
      delivery: String(c('delivery') || ''), otName: String(c('otName') || ''), otPhone: String(c('otPhone') || ''),
      teacher: String(c('teacher') || '')
    } };
  }
  throw new Error('查無此訂單，請確認連結是否完整');
}

/* ---------- 工作表初始化：標題、狀態下拉選單、顏色 ---------- */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');

    const statusRange = sheet.getRange(2, COL.status, sheet.getMaxRows() - 1, 1);
    statusRange.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build());

    const colors = { '未匯款': '#fde2e1', '待核對': '#fff4c2', '已確認': '#d9f2dd', '已取消': '#e0e0e0' };
    sheet.setConditionalFormatRules(Object.keys(colors).map(s =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(s).setBackground(colors[s]).setRanges([statusRange]).build()));
  }
  // 舊版建立的工作表：補上後來新增欄位的標題
  for (let c = 16; c <= HEADERS.length; c++) {
    if (sheet.getRange(1, c).getValue() !== HEADERS[c - 1]) {
      sheet.getRange(1, c).setValue(HEADERS[c - 1]).setFontWeight('bold');
    }
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
 * 後台管理 API（admin.html 使用）
 * 管理密碼：Apps Script「專案設定 → 指令碼屬性」新增 ADMIN_KEY
 * ===================================================================== */
function checkAdmin_(d) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key) throw new Error('尚未設定管理密碼（指令碼屬性 ADMIN_KEY）');
  if (!d.key || String(d.key) !== key) throw new Error('管理密碼錯誤');
}

// 從明細文字解析票券，例如「NT$600 x 2 張（座位：…）」
function parseDetail_(text) {
  const items = [], re = /NT\$([\d,]+) x (\d+) 張/g;
  let m;
  while ((m = re.exec(String(text || '')))) {
    items.push({ price: Number(m[1].replace(/,/g, '')), qty: Number(m[2]) });
  }
  return items;
}

function readOrders_() {
  const rows = getSheet_().getDataRange().getValues();
  const fmtTime = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : String(v || '');
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i], c = k => r[COL[k] - 1];
    if (!c('id')) continue;
    out.push({
      row: i + 1, orderId: String(c('id')), time: fmtTime(c('time')),
      name: String(c('name')), phone: String(c('phone')), email: String(c('email')),
      band: String(c('band')), last5: String(c('last5') || ''), status: String(c('status') || ''),
      delivery: String(c('delivery') || ''), otName: String(c('otName') || ''), otPhone: String(c('otPhone') || ''),
      teacher: String(c('teacher') || ''), seatPick: String(c('seatPick') || ''),
      note: String(c('note') || ''), total: Number(c('total')) || 0,
      token: String(c('token') || ''),
      sessions: [1, 2, 3].map(n => ({ session: n, text: String(c('detail' + n) || ''), items: parseDetail_(c('detail' + n)) }))
        .filter(x => x.items.length)
    });
  }
  return out;
}

function adminData_() {
  const orders = readOrders_().map(o => { const x = Object.assign({}, o); delete x.token; return x; });
  return { ok: true, orders: orders, quota: MailApp.getRemainingDailyQuota(), organizer: ORGANIZER, salesOpen: salesOpen_() };
}

// 批次更新付款狀態：updates = [{ orderId, status, note }]
function adminSetStatus_(d) {
  if (!Array.isArray(d.updates) || !d.updates.length) throw new Error('沒有要更新的訂單');
  const sheet = getSheet_();
  const ids = sheet.getRange(2, COL.id, Math.max(sheet.getLastRow() - 1, 1), 1).getValues()
    .map(r => String(r[0]).toUpperCase());
  let count = 0;
  d.updates.forEach(u => {
    if (STATUS.indexOf(u.status) < 0) throw new Error('付款狀態錯誤：' + u.status);
    const idx = ids.indexOf(String(u.orderId).toUpperCase());
    if (idx < 0) return;
    const r = idx + 2;
    sheet.getRange(r, COL.status).setValue(u.status);
    if (u.note) {
      const old = String(sheet.getRange(r, COL.note).getValue() || '');
      sheet.getRange(r, COL.note).setValue(old ? old + '\n' + u.note : u.note);
    }
    count++;
  });
  return { ok: true, updated: count };
}

// 寄信給指定訂單：orderIds、subject、body（可用 {姓名} {訂單編號} {總金額} {付款狀態} {場次} {訂單連結}）
function adminMail_(d) {
  if (!Array.isArray(d.orderIds) || !d.orderIds.length) throw new Error('沒有選擇收件訂單');
  if (!d.subject || !d.body) throw new Error('請填寫主旨與內容');
  const want = d.orderIds.map(x => String(x).toUpperCase());
  const orders = readOrders_().filter(o => want.indexOf(o.orderId.toUpperCase()) >= 0);
  const quota = MailApp.getRemainingDailyQuota();
  if (orders.length > quota) throw new Error('今日剩餘寄信額度 ' + quota + ' 封，不足以寄送 ' + orders.length + ' 封');

  const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const names = { 1: '台中中山堂 11/28 19:30', 2: '高雄至善廳 12/26 14:30', 3: '高雄至善廳 12/26 19:00' };
  const sent = [], failed = [];

  orders.forEach(o => {
    try {
      const base = pageUrlFor_(o.teacher);
      const link = o.token ? base + '?order=' + encodeURIComponent(o.orderId) + '&t=' + o.token : base;
      const vars = {
        '{姓名}': o.name, '{訂單編號}': o.orderId,
        '{總金額}': 'NT$' + Number(o.total).toLocaleString('en-US'), '{付款狀態}': o.status,
        '{場次}': o.sessions.map(x => names[x.session]).join('、'),
        '{取票方式}': o.delivery || '未選擇',
        '{團購老師}': o.teacher ? o.teacher + '老師' : '聲動室內樂團'
      };
      let subject = String(d.subject), html = esc(d.body);
      Object.keys(vars).forEach(k => {
        subject = subject.split(k).join(vars[k]);
        html = html.split(esc(k)).join(esc(vars[k]));
      });
      const btn = '<a href="' + esc(link) + '" style="display:inline-block;padding:10px 20px;background:#b4462a;color:#fff;' +
        'text-decoration:none;border-radius:6px;font-weight:bold;">查看訂單 / 填寫匯款資料</a>';
      html = html.split(esc('{訂單連結}')).join(btn).replace(/\n/g, '<br>');
      MailApp.sendEmail({
        to: o.email, subject: subject, name: ORGANIZER,
        htmlBody: '<div style="font-family:sans-serif;font-size:14px;color:#222;max-width:600px;">' + html +
          '<p style="margin-top:20px;">' + esc(ORGANIZER) + ' 敬上</p></div>'
      });
      sent.push(o.orderId);
    } catch (err) {
      failed.push({ orderId: o.orderId, error: err.message });
    }
  });
  return { ok: true, sent: sent.length, failed: failed, quota: MailApp.getRemainingDailyQuota() };
}

/* ---------- 發售控制 ----------
 * 指令碼屬性 SALES_OPEN：'0' = 停止接受訂票；其他（含未設定）= 開放
 */
function salesOpen_() {
  return PropertiesService.getScriptProperties().getProperty('SALES_OPEN') !== '0';
}

// d.open：true 開始發售 / false 停止發售；d.clear：開始發售時是否清空現有訂單（先備份）
function adminSetSales_(d) {
  const res = { ok: true, open: !!d.open, cleared: 0, backup: '' };
  if (d.open && d.clear) {
    const sheet = getSheet_();
    const n = sheet.getLastRow() - 1;
    if (n > 0) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const name = SHEET_NAME + '-清空前備份-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss');
      sheet.copyTo(ss).setName(name);
      sheet.getRange(2, 1, n, sheet.getLastColumn()).clearContent();
      res.cleared = n;
      res.backup = name;
    }
  }
  PropertiesService.getScriptProperties().setProperty('SALES_OPEN', d.open ? '1' : '0');
  return res;
}
