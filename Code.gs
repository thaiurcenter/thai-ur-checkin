/**
 * thai-ur office — ระบบเช็กอินตามตำแหน่ง (v2)
 * Code.gs : API ฝั่งเซิร์ฟเวอร์
 *
 * v2 เปลี่ยนจาก v1 ตรงไหน
 *  - ยืนยันตัวตนด้วยบัญชี Google จริง ไม่ใช่ลิงก์ส่วนตัวที่ส่งต่อกันได้อีกแล้ว
 *  - ไม่เสิร์ฟ HTML แล้ว ทำหน้าที่เป็น API อย่างเดียว หน้าเว็บอยู่ที่ GitHub Pages
 *  - Roster เก็บ SHA-256 ของอีเมล ไม่เก็บอีเมลจริง (เปลี่ยนได้ด้วย STORE_EMAIL)
 *
 * หลักการที่ยังเหมือนเดิม
 *  - ตัดสิน "อยู่ในพื้นที่หรือไม่" ที่ฝั่งเซิร์ฟเวอร์เท่านั้น
 *  - พิกัดออฟฟิศไม่เคยถูกส่งไปฝั่งเบราว์เซอร์
 *  - ไม่บันทึกพิกัดของพนักงาน เก็บแค่ระยะห่างที่ปัดแล้ว
 */

const SHEET_ROSTER = 'Roster';
const SHEET_LOG    = 'Log';
const SHEET_DAILY  = 'Daily';
const SHEET_CONFIG = 'Config';

const DEFAULT_CONFIG = {
  OFFICE_NAME:      'thai-ur office',
  OFFICE_LAT:       '0',
  OFFICE_LNG:       '0',
  RADIUS_M:         '100',
  MAX_ACCURACY_M:   '100',
  MAX_AGE_SEC:      '60',
  WORK_START:       '08:30',
  WORK_END:         '17:30',
  GRACE_MIN:        '5',
  WORK_DAYS:        '1,2,3,4,5',
  TIMEZONE:         'Asia/Bangkok',
  ALLOW_REOPEN:     'FALSE',
  GOOGLE_CLIENT_ID: '',
  STORE_EMAIL:      'FALSE',
  APP_URL:          ''
};

// ─────────────────────────────────────────────────────────────
// การตั้งค่า
// ─────────────────────────────────────────────────────────────
function getConfig() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('cfg');
  if (hit) return JSON.parse(hit);

  const cfg = Object.assign({}, DEFAULT_CONFIG);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      const k = String(r[0]).trim();
      if (k) cfg[k] = String(r[1]).trim();
    });
  }
  cache.put('cfg', JSON.stringify(cfg), 300);
  return cfg;
}

function clearConfigCache() {
  CacheService.getScriptCache().remove('cfg');
}

// ─────────────────────────────────────────────────────────────
// จุดเข้า
//  - doPost : ทางหลัก ส่งเป็น Content-Type: text/plain เพื่อเลี่ยง CORS preflight
//  - doGet  : ทางสำรอง JSONP (?callback=fn) เผื่อเบราว์เซอร์บล็อก CORS
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  let p;
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, message: 'รูปแบบข้อมูลไม่ถูกต้อง' });
  }
  return json_(handleApi_(p));
}

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const cb = String(p.callback || '').trim();

  if (!cb) {
    const cfg = getConfig();
    const where = cfg.APP_URL
      ? 'เปิดหน้าเช็กอินได้ที่ ' + cfg.APP_URL
      : 'กรุณาเปิดจากลิงก์ที่แอดมินให้ไว้';
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
      '<body style="font-family:system-ui,sans-serif;padding:40px;text-align:center;color:#334155">' +
      '<p>นี่คือส่วนเบื้องหลังของระบบเช็กอิน ไม่ใช่หน้าสำหรับใช้งาน</p><p>' +
      escapeHtml_(where) + '</p></body></html>');
  }

  // JSONP — ชื่อฟังก์ชัน callback ต้องเป็นตัวอักษร ตัวเลข _ $ เท่านั้น
  if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)) {
    return ContentService.createTextOutput('/* bad callback */')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  const body = JSON.stringify(handleApi_(p));
  return ContentService.createTextOutput(cb + '(' + body + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fail_(msg, code) {
  return { ok: false, message: msg, code: code || '' };
}

// ─────────────────────────────────────────────────────────────
// ตัวจัดการคำขอ
// ─────────────────────────────────────────────────────────────
function handleApi_(p) {
  const action = String((p && p.action) || '').toUpperCase();
  if (action === 'STATE') return apiState_(p);
  if (action === 'IN' || action === 'OUT') return apiCheck_(p, action);
  return fail_('คำสั่งไม่ถูกต้อง');
}

/** ขอสถานะของวันนี้ ไม่ต้องใช้พิกัด */
function apiState_(p) {
  const cfg = getConfig();
  const who = authenticate_(p, cfg);
  if (who.error) return fail_(who.error, who.code);

  const now = new Date();
  const st = getTodayState_(who.member.id, dateKey_(now, cfg.TIMEZONE));
  return {
    ok: true,
    name:       who.member.name,
    position:   who.member.position,
    office:     cfg.OFFICE_NAME,
    workStart:  cfg.WORK_START,
    workEnd:    cfg.WORK_END,
    graceMin:   Number(cfg.GRACE_MIN),
    serverNow:  now.getTime(),
    serverTime: Utilities.formatDate(now, cfg.TIMEZONE, 'HH:mm'),
    serverDate: Utilities.formatDate(now, cfg.TIMEZONE, 'd MMM yyyy'),
    isWorkday:  isWorkday_(now, cfg),
    allowReopen: String(cfg.ALLOW_REOPEN).toUpperCase() === 'TRUE',
    state: stateOut_(st, cfg.TIMEZONE)
  };
}

function stateOut_(st, tz) {
  return {
    hasIn:  !!st.inTime,
    hasOut: !!st.outTime,
    inTime:  st.inTime  ? Utilities.formatDate(st.inTime,  tz, 'HH:mm') : '',
    outTime: st.outTime ? Utilities.formatDate(st.outTime, tz, 'HH:mm') : '',
    inStatus:  st.inStatus  || '',
    outStatus: st.outStatus || ''
  };
}

/** เช็กอิน / เช็กเอาต์ */
function apiCheck_(p, action) {
  const cfg = getConfig();
  const tz = cfg.TIMEZONE;

  // 1) ตรวจรูปแบบข้อมูลก่อน
  const lat = Number(p.lat), lng = Number(p.lng);
  const acc = Number(p.accuracy), ageMs = Number(p.ageMs);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    return fail_('ไม่ได้รับพิกัดที่ถูกต้อง ลองกดใหม่อีกครั้ง');
  }
  if (!isFinite(acc) || acc <= 0) return fail_('ไม่ได้รับค่าความแม่นยำ ลองกดใหม่อีกครั้ง');

  // 2) ตรวจตัวตน
  const who = authenticate_(p, cfg);
  if (who.error) return fail_(who.error, who.code);
  const member = who.member;

  // 3) กันกดรัว
  const cache = CacheService.getScriptCache();
  const rlKey = 'rl_' + member.id;
  if (cache.get(rlKey)) return fail_('กดถี่เกินไป รอสักครู่แล้วลองใหม่');
  cache.put(rlKey, '1', 8);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return fail_('ระบบกำลังยุ่ง ลองใหม่อีกครั้งใน 2-3 วินาที');
  }

  try {
    const now = new Date();
    const flags = [];

    // 4) ความสดของพิกัด — ageMs คำนวณจากนาฬิกาเครื่องเดียวกัน จึงไม่เพี้ยนตามเขตเวลา
    if (!isFinite(ageMs) || ageMs < -5000 || ageMs > Number(cfg.MAX_AGE_SEC) * 1000) {
      logRow_(now, member, action, 'REJECT', '', '', Math.round(acc), '', 'พิกัดเก่าเกินไป', cfg);
      return fail_('พิกัดเก่าเกินไป กรุณากดใหม่อีกครั้ง');
    }

    // 5) ความแม่นยำ
    if (acc > Number(cfg.MAX_ACCURACY_M)) {
      logRow_(now, member, action, 'REJECT', '', '', Math.round(acc), '', 'ความแม่นยำต่ำ', cfg);
      return fail_('สัญญาณตำแหน่งไม่แม่นพอ (คลาดเคลื่อน ' + Math.round(acc) + ' ม.)\n' +
                   'ลองออกไปที่โล่ง เปิด GPS แล้วกดใหม่');
    }

    // 6) ระยะห่าง
    const oLat = Number(cfg.OFFICE_LAT), oLng = Number(cfg.OFFICE_LNG);
    if (!isFinite(oLat) || !isFinite(oLng) || (oLat === 0 && oLng === 0)) {
      return fail_('ยังไม่ได้ตั้งพิกัดออฟฟิศในระบบ กรุณาแจ้งแอดมิน');
    }
    const dist = haversineMeters_(lat, lng, oLat, oLng);
    const distR = Math.round(dist / 10) * 10;
    const radius = Number(cfg.RADIUS_M);

    if (dist > radius) {
      logRow_(now, member, action, 'REJECT', '', distR, Math.round(acc), '', 'อยู่นอกรัศมี', cfg);
      return fail_('คุณอยู่ห่างจาก ' + cfg.OFFICE_NAME + ' ประมาณ ' + formatDist_(dist) + '\n' +
                   'ต้องอยู่ภายใน ' + radius + ' เมตร จึงจะเช็ก' +
                   (action === 'IN' ? 'อิน' : 'เอาต์') + 'ได้');
    }

    // 7) อุปกรณ์ที่ผูกไว้ (เป็นแค่ธง ไม่บล็อก เพราะตัวตนยืนยันด้วยบัญชี Google แล้ว)
    const devId = String((p && p.deviceId) || '').trim().substring(0, 40);
    if (devId) {
      if (!member.device) setMemberDevice_(member.row, devId);
      else if (member.device !== devId) flags.push('อุปกรณ์ต่างจากที่ผูกไว้');
    }

    // 8) นาฬิกาเครื่องเพี้ยน
    const skew = Number(p.clientNow) - now.getTime();
    if (isFinite(skew) && Math.abs(skew) > 5 * 60 * 1000) flags.push('นาฬิกาเครื่องคลาดเคลื่อน');

    // 9) สถานะเวลา
    const ts = timeStatus_(now, action, cfg);
    ts.flags.forEach(function (f) { flags.push(f); });

    // 10) สถานะของวันนี้
    const dk = dateKey_(now, tz);
    const st = getTodayState_(member.id, dk);
    const allowReopen = String(cfg.ALLOW_REOPEN).toUpperCase() === 'TRUE';

    if (action === 'IN') {
      if (st.inTime && !allowReopen) {
        return fail_('วันนี้เช็กอินไปแล้วเมื่อ ' + Utilities.formatDate(st.inTime, tz, 'HH:mm') + ' น.');
      }
      if (st.outTime && !allowReopen) {
        return fail_('วันนี้เช็กเอาต์ไปแล้ว ไม่สามารถเช็กอินซ้ำได้');
      }
    } else {
      if (!st.inTime) return fail_('ยังไม่ได้เช็กอินวันนี้ จึงยังเช็กเอาต์ไม่ได้');
      if (st.outTime && !allowReopen) {
        return fail_('วันนี้เช็กเอาต์ไปแล้วเมื่อ ' + Utilities.formatDate(st.outTime, tz, 'HH:mm') + ' น.');
      }
    }

    // 11) บันทึก
    const flagStr = Array.from(new Set(flags)).join(', ');
    logRow_(now, member, action, 'OK', ts.status, distR, Math.round(acc), flagStr, '', cfg);
    upsertDaily_(now, member, action, ts.status, flagStr, cfg);

    return {
      ok: true,
      action: action,
      status: ts.status,
      time: Utilities.formatDate(now, tz, 'HH:mm'),
      distance: formatDist_(dist),
      message: (action === 'IN' ? 'เช็กอินสำเร็จ' : 'เช็กเอาต์สำเร็จ'),
      state: stateOut_(getTodayState_(member.id, dk), tz)
    };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
// ยืนยันตัวตนด้วย Google ID token
// ─────────────────────────────────────────────────────────────
function authenticate_(p, cfg) {
  const idToken = String((p && p.idToken) || '').trim();
  if (!idToken) return { error: 'กรุณาลงชื่อเข้าใช้ด้วยบัญชี Google', code: 'NO_TOKEN' };

  const clientId = String(cfg.GOOGLE_CLIENT_ID || '').trim();
  if (!clientId) return { error: 'ระบบยังตั้งค่าไม่ครบ กรุณาแจ้งแอดมิน', code: 'NO_CLIENT_ID' };

  const email = verifyIdToken_(idToken, clientId);
  if (!email) return { error: 'การลงชื่อเข้าใช้หมดอายุ กรุณาลงชื่อเข้าใช้ใหม่', code: 'BAD_TOKEN' };

  const member = findMemberByEmail_(email);
  if (!member) {
    return { error: 'บัญชี ' + maskEmail_(email) + ' ยังไม่ได้ลงทะเบียนในระบบ\n' +
                    'ถ้าคุณมีหลายบัญชี Google ลองสลับเป็นบัญชีที่แจ้งไว้กับแอดมิน',
             code: 'NOT_REGISTERED' };
  }
  if (String(member.status).toLowerCase() !== 'active') {
    return { error: 'บัญชีนี้ถูกปิดการใช้งาน', code: 'INACTIVE' };
  }
  return { member: member, email: email };
}

/**
 * ตรวจ ID token กับ Google แล้วคืนอีเมลที่ยืนยันแล้ว หรือ null
 * ใช้ tokeninfo endpoint เพราะ Apps Script ไม่มีไลบรารีตรวจลายเซ็น JWT ในตัว
 */
function verifyIdToken_(idToken, clientId) {
  const cache = CacheService.getScriptCache();
  const key = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).substring(0, 40);
  const hit = cache.get(key);
  if (hit) return hit === '-' ? null : hit;

  let info;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { cache.put(key, '-', 60); return null; }
    info = JSON.parse(res.getContentText());
  } catch (err) {
    return null;                                    // เครือข่ายมีปัญหา อย่าเพิ่ง cache ผลลบ
  }

  const issOk = info.iss === 'accounts.google.com' || info.iss === 'https://accounts.google.com';
  const audOk = info.aud === clientId;
  const expOk = Number(info.exp) * 1000 > Date.now();
  const verOk = info.email_verified === true || String(info.email_verified) === 'true';

  if (!issOk || !audOk || !expOk || !verOk || !info.email) {
    cache.put(key, '-', 60);
    return null;
  }

  const email = String(info.email).trim().toLowerCase();
  // cache สั้นกว่าอายุจริงของ token เสมอ
  const ttl = Math.max(30, Math.min(300, Math.floor((Number(info.exp) * 1000 - Date.now()) / 1000) - 30));
  cache.put(key, email, ttl);
  return email;
}

/** SHA-256 ของอีเมล (ตัวพิมพ์เล็ก ตัดช่องว่าง) เป็น hex */
function emailHash_(email) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(email).trim().toLowerCase(), Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/** som***@gmail.com — ไว้แสดงในข้อความ ไม่ได้เก็บลงชีต */
function maskEmail_(email) {
  const s = String(email);
  const at = s.indexOf('@');
  if (at < 1) return '***';
  const user = s.substring(0, at);
  const keep = Math.min(3, user.length);
  return user.substring(0, keep) + '***' + s.substring(at);
}

// ─────────────────────────────────────────────────────────────
// เวลา
// ─────────────────────────────────────────────────────────────
function hhmmToMinutes_(s) {
  const m = String(s).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

/** 1 = จันทร์ ... 7 = อาทิตย์ (ตามโซนเวลาที่ตั้งไว้) */
function isoDayOfWeek_(d, tz) {
  const u = new Date(Utilities.formatDate(d, tz, 'yyyy-MM-dd') + 'T00:00:00Z');
  const wd = u.getUTCDay();
  return wd === 0 ? 7 : wd;
}

function isWorkday_(d, cfg) {
  const days = String(cfg.WORK_DAYS).split(',').map(function (x) { return parseInt(x.trim(), 10); });
  return days.indexOf(isoDayOfWeek_(d, cfg.TIMEZONE)) !== -1;
}

function dateKey_(d, tz) {
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/**
 * IN  : ถึง WORK_START + GRACE = ปกติ, หลังจากนั้น = สาย
 * OUT : ตั้งแต่ WORK_END - GRACE = ปกติ, ก่อนหน้านั้น = ออกก่อน
 */
function timeStatus_(now, action, cfg) {
  const tz = cfg.TIMEZONE;
  const mins = Number(Utilities.formatDate(now, tz, 'H')) * 60 +
               Number(Utilities.formatDate(now, tz, 'm'));
  const start = hhmmToMinutes_(cfg.WORK_START);
  const end   = hhmmToMinutes_(cfg.WORK_END);
  const grace = Number(cfg.GRACE_MIN) || 0;
  const flags = [];

  if (!isWorkday_(now, cfg)) flags.push('นอกวันทำงาน');

  let status;
  if (action === 'IN') {
    if (mins <= start + grace) {
      status = 'ปกติ';
      if (start - mins > 120) flags.push('เช็กอินเช้ากว่าเวลางานมาก');
    } else {
      status = 'สาย ' + (mins - start) + ' นาที';
    }
  } else {
    if (mins >= end - grace) {
      status = 'ปกติ';
      if (mins - end > 180) flags.push('เช็กเอาต์ดึกกว่าเวลางานมาก');
    } else {
      status = 'ออกก่อน ' + (end - mins) + ' นาที';
    }
  }
  return { status: status, flags: flags };
}

// ─────────────────────────────────────────────────────────────
// ระยะทาง
// ─────────────────────────────────────────────────────────────
function haversineMeters_(lat1, lon1, lat2, lon2) {
  const R = 6371008.8;
  const toRad = function (d) { return d * Math.PI / 180; };
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(a)));
}

function formatDist_(m) {
  if (m < 1000) return Math.round(m / 10) * 10 + ' เมตร';
  return (Math.round(m / 100) / 10) + ' กม.';
}

// ─────────────────────────────────────────────────────────────
// Roster
// คอลัมน์: A รหัส | B ชื่อ-สกุล | C ตำแหน่ง | D สถานะ | E อีเมล | F Hash | G อุปกรณ์
// ─────────────────────────────────────────────────────────────
function findMemberByEmail_(email) {
  const hash = emailHash_(email);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  if (!sh || sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][5]).trim().toLowerCase() === hash) {
      return {
        row:      i + 2,
        id:       String(vals[i][0]).trim(),
        name:     String(vals[i][1]).trim(),
        position: String(vals[i][2]).trim(),
        status:   String(vals[i][3]).trim() || 'Active',
        device:   String(vals[i][6]).trim()
      };
    }
  }
  return null;
}

function setMemberDevice_(row, devId) {
  SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER).getRange(row, 7).setValue(devId);
}

// ─────────────────────────────────────────────────────────────
// Log / Daily
// ─────────────────────────────────────────────────────────────
function logRow_(now, member, action, result, timeStatus, dist, acc, flags, note, cfg) {
  SpreadsheetApp.getActive().getSheetByName(SHEET_LOG).appendRow([
    now,
    new Date(dateKey_(now, cfg.TIMEZONE) + 'T00:00:00'),
    member.id, member.name, member.position,
    action, result, timeStatus || '',
    dist === '' ? '' : dist,
    acc === '' ? '' : acc,
    flags || '', note || ''
  ]);
}

function getTodayState_(id, dk) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DAILY);
  const empty = { row: 0, inTime: null, outTime: null, inStatus: '', outStatus: '' };
  if (!sh || sh.getLastRow() < 2) return empty;

  const cfg = getConfig();
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 1, n, 9).getValues();
  for (let i = n - 1; i >= 0; i--) {
    const d = vals[i][0];
    if (!d) continue;
    const key = (d instanceof Date) ? dateKey_(d, cfg.TIMEZONE) : String(d).substring(0, 10);
    if (key === dk && String(vals[i][1]).trim() === id) {
      return {
        row: i + 2,
        inTime:  vals[i][4] instanceof Date ? vals[i][4] : null,
        outTime: vals[i][5] instanceof Date ? vals[i][5] : null,
        inStatus:  String(vals[i][7] || ''),
        outStatus: String(vals[i][8] || '')
      };
    }
  }
  return empty;
}

/** เข้า = ครั้งแรกของวัน, ออก = ครั้งล่าสุดของวัน */
function upsertDaily_(now, member, action, timeStatus, flags, cfg) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DAILY);
  const dk = dateKey_(now, cfg.TIMEZONE);
  const st = getTodayState_(member.id, dk);

  if (!st.row) {
    sh.appendRow([
      new Date(dk + 'T00:00:00'), member.id, member.name, member.position,
      action === 'IN' ? now : '', action === 'OUT' ? now : '',
      '', action === 'IN' ? timeStatus : '', action === 'OUT' ? timeStatus : '',
      flags || ''
    ]);
    return;
  }

  const row = st.row;
  if (action === 'IN') {
    if (!st.inTime) {
      sh.getRange(row, 5).setValue(now);
      sh.getRange(row, 8).setValue(timeStatus);
    }
  } else {
    sh.getRange(row, 6).setValue(now);
    sh.getRange(row, 9).setValue(timeStatus);
  }

  const inT  = action === 'IN'  ? (st.inTime || now) : st.inTime;
  const outT = action === 'OUT' ? now : st.outTime;
  if (inT && outT) sh.getRange(row, 7).setValue(Math.round((outT - inT) / 36000) / 100);

  if (flags) {
    const prev = String(sh.getRange(row, 10).getValue() || '');
    const set = new Set(prev.split(',').map(function (x) { return x.trim(); }).filter(String));
    flags.split(',').forEach(function (f) { if (f.trim()) set.add(f.trim()); });
    sh.getRange(row, 10).setValue(Array.from(set).join(', '));
  }
}
