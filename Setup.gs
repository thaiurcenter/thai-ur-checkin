/**
 * thai-ur office — ระบบเช็กอินตามตำแหน่ง (v2)
 * Setup.gs : ติดตั้งครั้งเดียว + เมนูแอดมิน
 */

// ─────────────────────────────────────────────────────────────
// เมนูในหน้า Sheet
// ─────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🕘 เช็กอิน thai-ur')
    .addItem('1. ติดตั้งระบบ (รันครั้งเดียว)', 'setupAll')
    .addItem('2. ลงทะเบียนอีเมลพนักงาน', 'registerEmails')
    .addSeparator()
    .addItem('ดู URL ของ API', 'showApiUrl')
    .addItem('ล้างแคชการตั้งค่า', 'clearConfigCache')
    .addItem('ทดสอบการตั้งค่า', 'selfTest')
    .addSeparator()
    .addItem('รีเซ็ตอุปกรณ์ที่ผูกไว้ (แถวที่เลือก)', 'resetDeviceForSelection')
    .addItem('สร้าง Daily ใหม่จาก Log ทั้งหมด', 'rebuildDaily')
    .addToUi();
}

// ─────────────────────────────────────────────────────────────
// ติดตั้งครั้งเดียว
// ─────────────────────────────────────────────────────────────
function setupAll() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(DEFAULT_CONFIG.TIMEZONE);

  setupConfigSheet_(ss);
  setupRosterSheet_(ss);
  setupLogSheet_(ss);
  setupDailySheet_(ss);

  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  ss.setActiveSheet(ss.getSheetByName(SHEET_CONFIG));
  clearConfigCache();

  SpreadsheetApp.getUi().alert(
    'ติดตั้งเรียบร้อย\n\n' +
    'ขั้นต่อไป (ดูละเอียดใน README):\n' +
    '1) สร้าง OAuth Client ID ที่ Google Cloud Console\n' +
    '2) ชีต Config: ใส่ OFFICE_LAT, OFFICE_LNG, GOOGLE_CLIENT_ID, APP_URL\n' +
    '3) ชีต Roster: ใส่ ชื่อ-สกุล / ตำแหน่ง / อีเมล Google ของพนักงาน\n' +
    '4) Deploy → Web app (Execute as: Me, Access: Anyone)\n' +
    '5) เมนู → "ลงทะเบียนอีเมลพนักงาน"\n' +
    '6) เมนู → "ทดสอบการตั้งค่า" เพื่อตรวจว่าครบ'
  );
}

function setupConfigSheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEET_CONFIG);
  sh.clear();
  const rows = [
    ['Key', 'Value', 'คำอธิบาย'],
    ['OFFICE_NAME',      DEFAULT_CONFIG.OFFICE_NAME,      'ชื่อสถานที่ที่แสดงบนหน้าเว็บ'],
    ['OFFICE_LAT',       DEFAULT_CONFIG.OFFICE_LAT,       '★ ละติจูดของออฟฟิศ'],
    ['OFFICE_LNG',       DEFAULT_CONFIG.OFFICE_LNG,       '★ ลองจิจูดของออฟฟิศ'],
    ['GOOGLE_CLIENT_ID', DEFAULT_CONFIG.GOOGLE_CLIENT_ID, '★ OAuth Client ID จาก Google Cloud Console (ลงท้าย .apps.googleusercontent.com)'],
    ['APP_URL',          DEFAULT_CONFIG.APP_URL,          '★ URL หน้าเว็บที่ GitHub Pages / Cloudflare Pages'],
    ['RADIUS_M',         DEFAULT_CONFIG.RADIUS_M,         'รัศมีที่ยอมให้เช็กอินได้ (เมตร)'],
    ['MAX_ACCURACY_M',   DEFAULT_CONFIG.MAX_ACCURACY_M,   'ค่า accuracy แย่สุดที่ยอมรับ (เมตร)'],
    ['MAX_AGE_SEC',      DEFAULT_CONFIG.MAX_AGE_SEC,      'อายุพิกัดสูงสุด (วินาที)'],
    ['WORK_START',       DEFAULT_CONFIG.WORK_START,       'เวลาเข้างาน HH:mm'],
    ['WORK_END',         DEFAULT_CONFIG.WORK_END,         'เวลาเลิกงาน HH:mm'],
    ['GRACE_MIN',        DEFAULT_CONFIG.GRACE_MIN,        'ผ่อนผัน (นาที) ทั้งเข้าสายและออกก่อน'],
    ['WORK_DAYS',        DEFAULT_CONFIG.WORK_DAYS,        'วันทำงาน 1=จันทร์ ... 7=อาทิตย์'],
    ['TIMEZONE',         DEFAULT_CONFIG.TIMEZONE,         'โซนเวลา'],
    ['ALLOW_REOPEN',     DEFAULT_CONFIG.ALLOW_REOPEN,     'TRUE = เช็กเอาต์แล้วกลับมาเช็กอินใหม่ในวันเดียวกันได้'],
    ['STORE_EMAIL',      DEFAULT_CONFIG.STORE_EMAIL,      'FALSE = เก็บอีเมลเป็นแบบย่อ som***@gmail.com / TRUE = เก็บอีเมลเต็ม']
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  sh.getRange(3, 2, 4, 1).setBackground('#fef3c7');     // ★ ช่องที่ต้องกรอกก่อนใช้งาน
  sh.setColumnWidth(1, 170).setColumnWidth(2, 380).setColumnWidth(3, 460);
  sh.setFrozenRows(1);
  sh.getRange(3, 2, 4, 1).setNumberFormat('@');          // lat / lng / client id / url เป็นข้อความ
  sh.getRange(10, 2, 2, 1).setNumberFormat('@');         // WORK_START / WORK_END
}

function setupRosterSheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEET_ROSTER);
  sh.clear();
  const head = ['รหัส', 'ชื่อ-สกุล', 'ตำแหน่ง', 'สถานะ', 'อีเมล Google', 'Hash', 'อุปกรณ์ที่ผูกไว้'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 80).setColumnWidth(2, 200).setColumnWidth(3, 180)
    .setColumnWidth(4, 90).setColumnWidth(5, 240).setColumnWidth(6, 200).setColumnWidth(7, 200);

  sh.getRange(2, 2, 1, 4).setValues([['(ชื่อ-สกุล)', '(ตำแหน่ง)', 'Active', '(อีเมล Google)']])
    .setFontColor('#9ca3af').setFontStyle('italic');

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true).setAllowInvalid(false).build();
  sh.getRange(2, 4, 500, 1).setDataValidation(rule);

  sh.getRange(1, 5, sh.getMaxRows(), 2).setNumberFormat('@');
  sh.getRange('F:G').setFontSize(8).setFontColor('#94a3b8');
  sh.getRange(1, 6, 1, 1).setNote('SHA-256 ของอีเมล — ระบบใช้ช่องนี้ในการจับคู่ ห้ามแก้เอง');
}

function setupLogSheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEET_LOG);
  sh.clear();
  const head = ['เวลา', 'วันที่', 'รหัส', 'ชื่อ-สกุล', 'ตำแหน่ง', 'ประเภท',
                'ผล', 'สถานะเวลา', 'ระยะ (ม.)', 'แม่นยำ (ม.)', 'ธง', 'หมายเหตุ'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  sh.setColumnWidth(1, 150).setColumnWidth(4, 180).setColumnWidth(5, 160).setColumnWidth(12, 260);
}

function setupDailySheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEET_DAILY);
  sh.clear();
  const head = ['วันที่', 'รหัส', 'ชื่อ-สกุล', 'ตำแหน่ง', 'เวลาเข้า', 'เวลาออก',
                'ชั่วโมงรวม', 'สถานะเข้า', 'สถานะออก', 'ธง'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('E:F').setNumberFormat('hh:mm');
  sh.getRange('G:G').setNumberFormat('0.00');
  sh.setColumnWidth(3, 180).setColumnWidth(4, 160);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ─────────────────────────────────────────────────────────────
// ลงทะเบียนอีเมล → เก็บเป็น hash
// ─────────────────────────────────────────────────────────────
function registerEmails() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  if (!sh) { ui.alert('ยังไม่มีชีต Roster — รัน "ติดตั้งระบบ" ก่อน'); return; }

  const last = sh.getLastRow();
  if (last < 2) { ui.alert('ยังไม่มีรายชื่อใน Roster'); return; }

  const cfg = getConfig();
  const keepEmail = String(cfg.STORE_EMAIL).toUpperCase() === 'TRUE';

  const rng = sh.getRange(2, 1, last - 1, 7);
  const vals = rng.getValues();
  let done = 0, skipped = 0, maxId = 0;
  const seen = {};
  const problems = [];

  vals.forEach(function (r) {
    const m = String(r[0] || '').match(/^EMP(\d+)$/);
    if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
  });

  vals.forEach(function (r, i) {
    const name  = String(r[1] || '').trim();
    const email = String(r[4] || '').trim();
    if (!name || name.charAt(0) === '(') return;

    if (!String(r[3] || '').trim()) vals[i][3] = 'Active';
    if (!String(r[0] || '').trim()) vals[i][0] = 'EMP' + String(++maxId).padStart(3, '0');

    if (!email || email.indexOf('***') !== -1) {          // ลงทะเบียนไปแล้ว หรือยังไม่ได้ใส่
      if (!String(r[5] || '').trim()) problems.push(name + ' — ยังไม่มีอีเมล');
      else skipped++;
      return;
    }
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      problems.push(name + ' — อีเมลผิดรูปแบบ: ' + email);
      return;
    }

    const h = emailHash_(email);
    if (seen[h]) { problems.push(name + ' — อีเมลซ้ำกับ ' + seen[h]); return; }
    seen[h] = name;

    vals[i][5] = h;
    vals[i][4] = keepEmail ? email.toLowerCase() : maskEmail_(email);
    done++;
  });

  rng.setValues(vals);
  clearConfigCache();

  let msg = 'ลงทะเบียนใหม่ ' + done + ' คน' + (skipped ? ' · เดิมอยู่แล้ว ' + skipped + ' คน' : '');
  if (!keepEmail) msg += '\n\nอีเมลถูกแทนด้วยแบบย่อแล้ว ระบบจับคู่ด้วย Hash ในคอลัมน์ F';
  if (problems.length) msg += '\n\nต้องแก้:\n• ' + problems.join('\n• ');
  msg += '\n\nบอกพนักงานให้เปิด ' + (cfg.APP_URL || '(ยังไม่ได้ตั้ง APP_URL)') +
         '\nแล้วลงชื่อเข้าใช้ด้วยบัญชี Google ที่ลงทะเบียนไว้';
  ui.alert(msg);
}

// ─────────────────────────────────────────────────────────────
// เครื่องมือแอดมิน
// ─────────────────────────────────────────────────────────────
function showApiUrl() {
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  SpreadsheetApp.getUi().alert(
    url ? 'URL ของ API (เอาไปใส่ใน web/config.js เป็น API_URL)\n\n' + url
        : 'ยังไม่ได้ Deploy เป็น Web App');
}

/** ตรวจว่าตั้งค่าครบและใช้การได้จริงไหม */
function selfTest() {
  const cfg = getConfig();
  const out = [];
  const bad = function (s) { out.push('✗ ' + s); };
  const ok  = function (s) { out.push('✓ ' + s); };

  const lat = Number(cfg.OFFICE_LAT), lng = Number(cfg.OFFICE_LNG);
  if (!isFinite(lat) || !isFinite(lng) || (lat === 0 && lng === 0)) bad('ยังไม่ได้ตั้งพิกัดออฟฟิศ');
  else ok('พิกัดออฟฟิศ ' + lat + ', ' + lng);

  const cid = String(cfg.GOOGLE_CLIENT_ID || '').trim();
  if (!cid) bad('ยังไม่ได้ใส่ GOOGLE_CLIENT_ID');
  else if (cid.indexOf('.apps.googleusercontent.com') === -1) bad('GOOGLE_CLIENT_ID รูปแบบไม่ถูกต้อง');
  else ok('GOOGLE_CLIENT_ID ตั้งแล้ว');

  if (!String(cfg.APP_URL || '').trim()) bad('ยังไม่ได้ใส่ APP_URL');
  else ok('APP_URL ' + cfg.APP_URL);

  if (hhmmToMinutes_(cfg.WORK_START) === null) bad('WORK_START ต้องเป็นรูปแบบ HH:mm');
  if (hhmmToMinutes_(cfg.WORK_END) === null)   bad('WORK_END ต้องเป็นรูปแบบ HH:mm');

  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  if (!url) bad('ยังไม่ได้ Deploy เป็น Web App');
  else ok('Deploy แล้ว');

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  let n = 0;
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 6, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      if (String(r[0]).trim().length === 64) n++;
    });
  }
  if (!n) bad('ยังไม่มีพนักงานที่ลงทะเบียนอีเมลแล้ว');
  else ok('พนักงานที่ลงทะเบียนแล้ว ' + n + ' คน');

  // เชื่อมต่อออกนอกได้ไหม (จำเป็นสำหรับตรวจ ID token)
  try {
    const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test',
                                { muteHttpExceptions: true });
    ok('ติดต่อ Google tokeninfo ได้ (HTTP ' + r.getResponseCode() + ')');
  } catch (e) {
    bad('ติดต่อ Google tokeninfo ไม่ได้: ' + e.message);
  }

  SpreadsheetApp.getUi().alert('ผลการตรวจ\n\n' + out.join('\n'));
}

function resetDeviceForSelection() {
  const sh = SpreadsheetApp.getActive().getActiveSheet();
  if (sh.getName() !== SHEET_ROSTER) {
    SpreadsheetApp.getUi().alert('เลือกแถวในชีต Roster ก่อน');
    return;
  }
  const r = sh.getActiveRange();
  sh.getRange(r.getRow(), 7, r.getNumRows(), 1).clearContent();
  SpreadsheetApp.getUi().alert('ล้างอุปกรณ์ที่ผูกไว้แล้ว ' + r.getNumRows() + ' แถว');
}

/** สร้างชีต Daily ใหม่ทั้งหมดจาก Log (ใช้เมื่อแก้ Log ย้อนหลัง) */
function rebuildDaily() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(SHEET_LOG);
  const daily = ss.getSheetByName(SHEET_DAILY);
  const cfg = getConfig();
  if (daily.getLastRow() > 1) {
    daily.getRange(2, 1, daily.getLastRow() - 1, daily.getLastColumn()).clearContent();
  }
  if (log.getLastRow() < 2) return;

  const rows = log.getRange(2, 1, log.getLastRow() - 1, 12).getValues();
  const map = {};
  rows.forEach(function (r) {
    if (String(r[6]) !== 'OK') return;
    const dateKey = Utilities.formatDate(new Date(r[1]), cfg.TIMEZONE, 'yyyy-MM-dd');
    const key = dateKey + '|' + r[2];
    if (!map[key]) {
      map[key] = { date: new Date(r[1]), id: r[2], name: r[3], pos: r[4],
                   inT: '', outT: '', inS: '', outS: '', flags: [] };
    }
    const d = map[key];
    const t = new Date(r[0]);
    if (String(r[5]) === 'IN') { if (!d.inT) { d.inT = t; d.inS = r[7]; } }
    else { d.outT = t; d.outS = r[7]; }
    if (r[10]) d.flags.push(String(r[10]));
  });

  const out = Object.keys(map).sort().map(function (k) {
    const d = map[k];
    const hrs = (d.inT && d.outT) ? ((d.outT - d.inT) / 3600000) : '';
    return [d.date, d.id, d.name, d.pos, d.inT, d.outT,
            hrs === '' ? '' : Math.round(hrs * 100) / 100,
            d.inS, d.outS, Array.from(new Set(d.flags)).join(', ')];
  });
  if (out.length) daily.getRange(2, 1, out.length, 10).setValues(out);
  SpreadsheetApp.getUi().alert('สร้าง Daily ใหม่แล้ว ' + out.length + ' แถว');
}
