/**
 * ================================================================
 * Toy Station Plus+ | ระบบวงล้อเสี่ยงโชคประจำสาขา  (v2)
 * Backend : Google Apps Script + Google Sheets
 * ----------------------------------------------------------------
 * สิ่งที่เปลี่ยนจาก v1
 *  1) ระบบบังคับแชร์ย้ายมาตรวจฝั่ง SERVER ด้วย share token + จับเวลา
 *     (v1 ส่งค่า shared เป็น boolean จาก client ซึ่งปลอมได้)
 *  2) รองรับ Landing Page ภายนอกสำหรับ OG image + แคปชั่นสุ่ม
 *  3) เพิ่ม wheel_font_scale สำหรับปรับขนาดตัวอักษรบนวงล้อ
 * ================================================================
 */

const SHEET_ID = '';                 // ว่างไว้ = ใช้ไฟล์ที่สคริปต์ผูกอยู่
const TZ = 'Asia/Bangkok';

const SH = { PRIZES:'Prizes', LOG:'SpinLog', CFG:'Config', BR:'Branches', QUOTA:'Quota' };

const HEADERS = {
  Prizes  : ['prize_id','branch_code','prize_name','color','weight','stock','active'],
  SpinLog : ['timestamp','branch_code','phone','prize_id','prize_name','redeem_code','status','expire_at','redeemed_at','staff','shared'],
  Config  : ['key','value','note'],
  Branches: ['branch_code','branch_name','active'],
  Quota   : ['key','phone','branch_code','datekey','used','last_spin']
};

const DEFAULT_CFG = [
  ['event_name',      'กิจกรรมวงล้อเสี่ยงโชค Toy Station Plus+', 'ชื่อกิจกรรมที่แสดงบนหน้าเว็บ'],
  ['max_spin_phone',  '10',   'จำนวนครั้งสูงสุดต่อ 1 เบอร์'],
  ['limit_scope',     'DAY',  'DAY = นับใหม่ทุกวัน / EVENT = นับรวมตลอดกิจกรรม'],
  ['limit_per_branch','Y',    'Y = นับแยกรายสาขา / N = นับรวมทุกสาขา'],
  ['start_date',      '',     'วันเริ่มกิจกรรม (yyyy-MM-dd) ว่าง = ไม่จำกัด'],
  ['end_date',        '',     'วันสิ้นสุดกิจกรรม (yyyy-MM-dd) ว่าง = ไม่จำกัด'],
  ['redeem_minutes',  '30',   'อายุโค้ดรับรางวัล (นาที)'],
  ['share_mode',      'SOFT', 'SOFT = บังคับแชร์ (ตรวจเวลาฝั่ง server) / OFF = ไม่บังคับ'],
  ['share_min_seconds','8',   'ต้องออกจากหน้าไปอย่างน้อยกี่วินาทีจึงนับว่าแชร์แล้ว'],
  ['share_every_spin','N',    'Y = ต้องแชร์ทุกครั้งก่อนหมุน / N = แชร์ครั้งเดียวปลดล็อกทั้ง session'],
  ['share_landing_base','',   'URL โฟลเดอร์ Landing Page เช่น https://user.github.io/tsp-wheel (ว่าง = แชร์ลิงก์ตรงเข้าเว็บแอป ไม่มีภาพ)'],
  ['share_variants',  '4',    'จำนวนไฟล์ Landing (1.html..N.html) ที่ระบบจะสุ่ม'],
  ['share_captions',  'มาหมุนวงล้อลุ้นของรางวัลกับ Toy Station Plus+ กันครับ 🎁|วันนี้ที่ Toy Station Plus+ มีวงล้อเสี่ยงโชคแจกของรางวัลด้วย ลองมาหมุนกันดู 🕹️|ตู้คีบตุ๊กตา Toy Station Plus+ จัดกิจกรรมหมุนวงล้อ ลุ้นรางวัลฟรี ใครอยู่แถวนี้ห้ามพลาด ✨|หมุนวงล้อรับของรางวัลที่ Toy Station Plus+ สนุกมาก แนะนำเลย 🎯', 'แคปชั่นสุ่ม คั่นด้วยเครื่องหมาย |'],
  ['wheel_font_scale','1.0',  'ตัวคูณขนาดตัวอักษรบนวงล้อ (0.7 - 1.4) ระบบ Auto-fit อยู่แล้ว ใช้ปรับละเอียด'],
  ['pdpa_text',       'ข้าพเจ้ายินยอมให้บริษัทฯ เก็บเบอร์โทรศัพท์เพื่อใช้ยืนยันสิทธิ์และติดต่อเรื่องของรางวัลเท่านั้น', 'ข้อความยินยอม PDPA'],
  ['footer_note',     'ของรางวัลมีจำนวนจำกัด • ขอสงวนสิทธิ์เปลี่ยนแปลงโดยไม่ต้องแจ้งล่วงหน้า', 'ข้อความท้ายหน้า']
];

/* ================================================================
 * 1) ติดตั้ง / อัปเดตโครงสร้าง — รันทุกครั้งที่อัปเดตเวอร์ชัน (ปลอดภัยกับข้อมูลเดิม)
 * ================================================================ */
function setup() {
  const s = ss_();
  Object.keys(HEADERS).forEach(function (name) {
    let sh = s.getSheetByName(name);
    if (!sh) sh = s.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]])
        .setFontWeight('bold').setBackground('#E8EAED');
      sh.setFrozenRows(1);
    }
  });

  const br = s.getSheetByName(SH.BR);
  if (br.getLastRow() < 2) {
    const rows = [];
    for (let i = 1; i <= 10; i++) rows.push(['BR' + ('0' + i).slice(-2), 'สาขาที่ ' + i, i === 1 ? 'Y' : 'N']);
    br.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  const cf = s.getSheetByName(SH.CFG);
  const have = {};
  readAll_(SH.CFG).forEach(function (r) { have[String(r.key)] = 1; });
  const add = DEFAULT_CFG.filter(function (d) { return !have[d[0]]; });
  if (add.length) cf.getRange(cf.getLastRow() + 1, 1, add.length, 3).setValues(add);

  const pz = s.getSheetByName(SH.PRIZES);
  if (pz.getLastRow() < 2) {
    pz.getRange(2, 1, 5, 7).setValues([
      ['P001','BR01','ตุ๊กตาพวงกุญแจ','#8ED1FC', 40, 100,'Y'],
      ['P002','BR01','ตุ๊กตา',        '#F28B82', 10,  20,'Y'],
      ['P003','BR01','คูปองเล่นฟรี',  '#FBD35B', 25,  80,'Y'],
      ['P004','BR01','หมุนใหม่อีกครั้ง','#A8E6A3',20, 999,'Y'],
      ['P005','BR01','ส่วนลด 20 บาท', '#D9B8FF',  5,  30,'Y']
    ]);
  }

  const pr = PropertiesService.getScriptProperties();
  if (!pr.getProperty('ADMIN_PIN')) pr.setProperty('ADMIN_PIN', '246810');
  if (!pr.getProperty('STAFF_PIN')) pr.setProperty('STAFF_PIN', '1357');

  SpreadsheetApp.getActive().toast('อัปเดตโครงสร้างเรียบร้อย — เพิ่ม config ใหม่ ' + add.length + ' รายการ', 'Setup', 12);
}

/* ================================================================
 * 2) Router
 * ================================================================ */
function doGet(e) {
  const p = String((e && e.parameter && e.parameter.page) || 'spin').toLowerCase();
  const file = (p === 'admin') ? 'admin' : (p === 'verify') ? 'verify' : 'spin';

  const t = HtmlService.createTemplateFromFile(file);
  t.branch = String((e && e.parameter && e.parameter.b) || '').toUpperCase();
  t.webUrl = ScriptApp.getService().getUrl();

  return t.evaluate()
    .setTitle('Toy Station Plus+ | วงล้อเสี่ยงโชค')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/* ================================================================
 * 3) Helper
 * ================================================================ */
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive(); }

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีท ' + name + ' — กรุณารันฟังก์ชัน setup() ก่อน');
  return sh;
}

function readAll_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const head = values.shift().map(String);
  return values.map(function (r, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, c) { if (h) o[h] = r[c]; });
    return o;
  });
}

function getCfg_() {
  const o = {};
  readAll_(SH.CFG).forEach(function (r) { if (r.key) o[String(r.key).trim()] = String(r.value); });
  DEFAULT_CFG.forEach(function (d) { if (o[d[0]] === undefined) o[d[0]] = d[1]; });
  return o;
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function normPhone_(p) {
  const d = String(p || '').replace(/[^0-9]/g, '');
  return /^0[0-9]{8,9}$/.test(d) ? d : '';
}
function maskPhone_(p) { return String(p).replace(/^(\d{3})\d{3,4}(\d{3})$/, '$1-xxx-$2'); }
function makeCode_() {
  const A = 'ACDEFGHJKLMNPQRTUVWXY34679';
  let s = ''; for (let i = 0; i < 6; i++) s += A.charAt(Math.floor(Math.random() * A.length));
  return s;
}

/* ================================================================
 * 4) ระบบบังคับแชร์ — ตรวจฝั่ง SERVER
 * ----------------------------------------------------------------
 *   1. apiShareStart  -> server ออก token + จดเวลาเริ่ม (t0)
 *   2. หน้าเว็บเปิดหน้าต่างแชร์ Facebook
 *   3. apiShareConfirm -> server เช็ค (now - t0) >= share_min_seconds
 *   4. apiSpin ต้องแนบ token ที่ผ่านแล้วเท่านั้น
 *
 *   ข้อจำกัดที่ต้องรับทราบ: Facebook ไม่มี API ให้ตรวจว่าโพสต์จริงหรือไม่
 *   ระบบนี้ยืนยันได้แค่ "เปิดหน้าแชร์และอยู่นานพอ" ไม่ใช่ "โพสต์แล้ว"
 * ================================================================ */
function shareCache_() { return CacheService.getScriptCache(); }

function apiShareStart(branch, phone) {
  branch = String(branch || '').toUpperCase();
  const ph = normPhone_(phone);
  if (!ph) return { ok: false, msg: 'เบอร์โทรไม่ถูกต้อง' };

  const cfg = getCfg_();
  const token = 'shr_' + Utilities.getUuid();
  shareCache_().put(token, JSON.stringify({ p: ph, b: branch, t0: Date.now(), ok: false }), 1800);

  const n = Math.max(1, Number(cfg.share_variants) || 1);
  const v = 1 + Math.floor(Math.random() * n);
  const caps = String(cfg.share_captions).split('|').map(function (x) { return x.trim(); }).filter(String);

  return {
    ok: true, token: token,
    url: landingUrl_(cfg, branch, v),
    caption: caps.length ? caps[Math.floor(Math.random() * caps.length)] : '',
    minSec: Number(cfg.share_min_seconds) || 8
  };
}

function landingUrl_(cfg, branch, v) {
  const base = String(cfg.share_landing_base || '').trim().replace(/\/+$/, '');
  if (!base) return ScriptApp.getService().getUrl() + '?b=' + branch;
  return base + '/' + v + '.html?b=' + branch;
}

function apiShareConfirm(token) {
  const raw = shareCache_().get(String(token || ''));
  if (!raw) return { ok: false, msg: 'เซสชันแชร์หมดอายุ กรุณากดปุ่มแชร์ใหม่อีกครั้ง' };
  const o = JSON.parse(raw);
  const need = Number(getCfg_().share_min_seconds) || 8;
  const elapsed = (Date.now() - o.t0) / 1000;
  if (elapsed < need) {
    return { ok: false, wait: Math.ceil(need - elapsed),
             msg: 'กรุณาแชร์ให้เรียบร้อยก่อน (รออีก ' + Math.ceil(need - elapsed) + ' วินาที)' };
  }
  o.ok = true;
  shareCache_().put(token, JSON.stringify(o), 1800);
  return { ok: true };
}

function checkShare_(cfg, token, phone, branch) {
  if (String(cfg.share_mode).toUpperCase() === 'OFF') return { ok: true };
  const raw = shareCache_().get(String(token || ''));
  if (!raw) return { ok: false, msg: 'กรุณากดปุ่มแชร์กิจกรรมก่อนหมุนวงล้อ', needShare: true };
  const o = JSON.parse(raw);
  if (!o.ok) return { ok: false, msg: 'ยังไม่ได้ยืนยันการแชร์', needShare: true };
  if (o.p !== phone || o.b !== branch) return { ok: false, msg: 'ข้อมูลการแชร์ไม่ตรงกับผู้ใช้งาน', needShare: true };
  return { ok: true, consume: String(cfg.share_every_spin).toUpperCase() === 'Y' };
}

/* ================================================================
 * 5) API ฝั่งลูกค้า
 * ================================================================ */
function apiGetWheel(branch) {
  branch = String(branch || '').toUpperCase();
  const cfg = getCfg_();

  const period = checkPeriod_(cfg);
  if (!period.ok) return period;

  const br = readAll_(SH.BR).filter(function (b) { return String(b.branch_code).toUpperCase() === branch; })[0];
  if (!br) return { ok: false, msg: 'ไม่พบรหัสสาขา (' + branch + ')' };
  if (String(br.active).toUpperCase() !== 'Y') return { ok: false, msg: 'สาขานี้ยังไม่เปิดกิจกรรม' };

  const prizes = livePrizes_(branch);
  if (prizes.length < 2) return { ok: false, msg: 'ของรางวัลของสาขานี้หมดแล้ว หรือยังตั้งค่าไม่ครบ' };

  return {
    ok: true, branch: branch, branchName: String(br.branch_name), eventName: cfg.event_name,
    prizes: prizes.map(function (p) {
      return { id: String(p.prize_id), name: String(p.prize_name), color: String(p.color || '#8ED1FC') };
    }),
    maxSpin: Number(cfg.max_spin_phone),
    shareMode: String(cfg.share_mode).toUpperCase(),
    shareEverySpin: String(cfg.share_every_spin).toUpperCase() === 'Y',
    fontScale: Number(cfg.wheel_font_scale) || 1,
    pdpaText: cfg.pdpa_text, footerNote: cfg.footer_note,
    redeemMinutes: Number(cfg.redeem_minutes)
  };
}

function livePrizes_(branch) {
  return readAll_(SH.PRIZES).filter(function (p) {
    return String(p.branch_code).toUpperCase() === branch
      && String(p.active).toUpperCase() === 'Y'
      && Number(p.stock) > 0 && Number(p.weight) > 0;
  });
}

function checkPeriod_(cfg) {
  const d = today_();
  if (cfg.start_date && d < cfg.start_date) return { ok: false, msg: 'กิจกรรมยังไม่เริ่ม (เริ่ม ' + cfg.start_date + ')' };
  if (cfg.end_date && d > cfg.end_date)     return { ok: false, msg: 'กิจกรรมสิ้นสุดแล้ว (' + cfg.end_date + ')' };
  return { ok: true };
}

function apiCheckPhone(branch, phone) {
  const ph = normPhone_(phone);
  if (!ph) return { ok: false, msg: 'รูปแบบเบอร์โทรไม่ถูกต้อง (ต้องขึ้นต้นด้วย 0 และมี 9-10 หลัก)' };
  const cfg = getCfg_();
  const q = quotaOf_(cfg, ph, String(branch || '').toUpperCase());
  return { ok: true, phone: ph, used: q.used,
           remaining: Math.max(0, Number(cfg.max_spin_phone) - q.used), max: Number(cfg.max_spin_phone) };
}

function quotaKey_(cfg, phone, branch) {
  const scopeBranch = String(cfg.limit_per_branch).toUpperCase() === 'Y' ? branch : 'ALL';
  const scopeDate   = String(cfg.limit_scope).toUpperCase() === 'DAY' ? today_() : 'EVENT';
  return phone + '|' + scopeBranch + '|' + scopeDate;
}

function quotaOf_(cfg, phone, branch) {
  const key = quotaKey_(cfg, phone, branch);
  const rows = readAll_(SH.QUOTA);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].key) === key) return { key: key, used: Number(rows[i].used) || 0, row: rows[i]._row };
  }
  return { key: key, used: 0, row: 0 };
}

function apiSpin(branch, phone, shareToken, clientIds) {
  branch = String(branch || '').toUpperCase();
  const ph = normPhone_(phone);
  if (!ph) return { ok: false, msg: 'รูปแบบเบอร์โทรไม่ถูกต้อง' };

  const cfg = getCfg_();
  const period = checkPeriod_(cfg);
  if (!period.ok) return period;

  // >>> ด่านตรวจการแชร์ ฝั่ง server <<<
  const shr = checkShare_(cfg, shareToken, ph, branch);
  if (!shr.ok) return shr;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, msg: 'ระบบกำลังมีผู้ใช้งานพร้อมกันจำนวนมาก กรุณากดใหม่อีกครั้ง' };

  try {
    const max = Number(cfg.max_spin_phone);
    const q = quotaOf_(cfg, ph, branch);
    if (q.used >= max) return { ok: false, msg: 'เบอร์นี้ใช้สิทธิ์ครบ ' + max + ' ครั้งแล้ว', remaining: 0 };

    const pool = livePrizes_(branch);
    if (pool.length < 2) return { ok: false, msg: 'ของรางวัลหมดแล้ว', reload: true };

    const total = pool.reduce(function (s, p) { return s + Number(p.weight); }, 0);
    let r = Math.random() * total, win = pool[pool.length - 1];
    for (let i = 0; i < pool.length; i++) {
      r -= Number(pool[i].weight);
      if (r <= 0) { win = pool[i]; break; }
    }

    sheet_(SH.PRIZES).getRange(win._row, HEADERS.Prizes.indexOf('stock') + 1).setValue(Number(win.stock) - 1);

    const code = uniqueCode_();
    const now = new Date();
    const expire = new Date(now.getTime() + Number(cfg.redeem_minutes) * 60000);
    sheet_(SH.LOG).appendRow([now, branch, ph, String(win.prize_id), String(win.prize_name),
      code, 'PENDING', expire, '', '', 'Y']);

    const used = q.used + 1;
    const qs = sheet_(SH.QUOTA);
    if (q.row) qs.getRange(q.row, 5, 1, 2).setValues([[used, now]]);
    else { const parts = q.key.split('|'); qs.appendRow([q.key, ph, parts[1], parts[2], used, now]); }

    if (shr.consume) shareCache_().remove(String(shareToken));   // โหมดแชร์ทุกครั้ง

    let idx = -1;
    if (clientIds && clientIds.length) idx = clientIds.indexOf(String(win.prize_id));

    SpreadsheetApp.flush();
    return {
      ok: true, index: idx,
      prizeId: String(win.prize_id), prizeName: String(win.prize_name), code: code,
      expireAt: Utilities.formatDate(expire, TZ, 'HH:mm'), expireMs: expire.getTime(),
      remaining: max - used, needShareAgain: !!shr.consume, reload: (idx < 0)
    };
  } finally { lock.releaseLock(); }
}

function uniqueCode_() {
  const used = {};
  readAll_(SH.LOG).forEach(function (r) { if (String(r.status) === 'PENDING') used[String(r.redeem_code)] = 1; });
  let c = makeCode_(), guard = 0;
  while (used[c] && guard++ < 50) c = makeCode_();
  return c;
}

/* ================================================================
 * 6) API ฝั่งพนักงาน
 * ================================================================ */
function staffLogin(pin) {
  const pr = PropertiesService.getScriptProperties();
  if (String(pin) !== String(pr.getProperty('STAFF_PIN')) && String(pin) !== String(pr.getProperty('ADMIN_PIN')))
    return { ok: false, msg: 'PIN ไม่ถูกต้อง' };
  return { ok: true, token: newToken_('staff') };
}

function apiLookup(token, code) {
  chkToken_(token, 'staff');
  const c = String(code || '').trim().toUpperCase();
  const rows = readAll_(SH.LOG);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i].redeem_code).toUpperCase() === c) {
      const exp = rows[i].expire_at ? new Date(rows[i].expire_at) : null;
      return {
        ok: true, found: true, row: rows[i]._row,
        branch: String(rows[i].branch_code), phone: maskPhone_(rows[i].phone),
        prize: String(rows[i].prize_name), status: String(rows[i].status),
        time: Utilities.formatDate(new Date(rows[i].timestamp), TZ, 'dd/MM/yyyy HH:mm'),
        expireTime: exp ? Utilities.formatDate(exp, TZ, 'dd/MM/yyyy HH:mm') : '-',
        expired: !!(exp && new Date() > exp)
      };
    }
  }
  return { ok: true, found: false };
}

function apiRedeem(token, row, staffName) {
  chkToken_(token, 'staff');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SH.LOG);
    const cStatus = HEADERS.SpinLog.indexOf('status') + 1;
    const cRedeem = HEADERS.SpinLog.indexOf('redeemed_at') + 1;
    const cStaff  = HEADERS.SpinLog.indexOf('staff') + 1;
    const cExpire = HEADERS.SpinLog.indexOf('expire_at') + 1;

    if (String(sh.getRange(row, cStatus).getValue()) === 'REDEEMED')
      return { ok: false, msg: 'โค้ดนี้ถูกใช้แลกรางวัลไปแล้ว' };

    const exp = sh.getRange(row, cExpire).getValue();
    if (exp && new Date() > new Date(exp)) {
      sh.getRange(row, cStatus).setValue('EXPIRED');
      return { ok: false, msg: 'โค้ดหมดอายุแล้ว' };
    }

    sh.getRange(row, cStatus).setValue('REDEEMED');
    sh.getRange(row, cRedeem).setValue(new Date());
    sh.getRange(row, cStaff).setValue(String(staffName || ''));
    SpreadsheetApp.flush();
    return { ok: true, msg: 'ยืนยันการรับรางวัลเรียบร้อย' };
  } finally { lock.releaseLock(); }
}

/* ================================================================
 * 7) API ฝั่งแอดมิน
 * ================================================================ */
function newToken_(role) {
  const t = role + '_' + Utilities.getUuid();
  CacheService.getScriptCache().put(t, '1', 3600);
  return t;
}
function chkToken_(token, role) {
  if (!token || String(token).indexOf(role + '_') !== 0) throw new Error('ไม่มีสิทธิ์เข้าถึง');
  if (!CacheService.getScriptCache().get(token)) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return true;
}
function adminLogin(pin) {
  if (String(pin) !== String(PropertiesService.getScriptProperties().getProperty('ADMIN_PIN')))
    return { ok: false, msg: 'PIN ไม่ถูกต้อง' };
  return { ok: true, token: newToken_('admin') };
}

function adminBootstrap(token) {
  chkToken_(token, 'admin');
  return {
    ok: true, cfg: getCfg_(), webUrl: ScriptApp.getService().getUrl(),
    prizes: readAll_(SH.PRIZES).map(function (p) {
      return { row: p._row, id: String(p.prize_id), branch: String(p.branch_code).toUpperCase(),
        name: String(p.prize_name), color: String(p.color), weight: Number(p.weight),
        stock: Number(p.stock), active: String(p.active).toUpperCase() };
    }),
    branches: readAll_(SH.BR).map(function (b) {
      return { row: b._row, code: String(b.branch_code).toUpperCase(),
        name: String(b.branch_name), active: String(b.active).toUpperCase() };
    })
  };
}

function adminSavePrize(token, p) {
  chkToken_(token, 'admin');
  const sh = sheet_(SH.PRIZES);
  const row = [String(p.id || nextPrizeId_()).toUpperCase(), String(p.branch).toUpperCase(),
    String(p.name), String(p.color || '#8ED1FC'), Number(p.weight) || 0, Number(p.stock) || 0,
    String(p.active).toUpperCase() === 'Y' ? 'Y' : 'N'];
  if (p.row) sh.getRange(Number(p.row), 1, 1, 7).setValues([row]); else sh.appendRow(row);
  SpreadsheetApp.flush();
  return { ok: true };
}

function adminDeletePrize(token, row) {
  chkToken_(token, 'admin'); sheet_(SH.PRIZES).deleteRow(Number(row)); return { ok: true };
}

function nextPrizeId_() {
  let max = 0;
  readAll_(SH.PRIZES).forEach(function (r) {
    const m = String(r.prize_id).match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1]));
  });
  return 'P' + ('00' + (max + 1)).slice(-3);
}

function adminSaveBranch(token, b) {
  chkToken_(token, 'admin');
  const sh = sheet_(SH.BR);
  const row = [String(b.code).toUpperCase(), String(b.name), String(b.active).toUpperCase() === 'Y' ? 'Y' : 'N'];
  if (b.row) sh.getRange(Number(b.row), 1, 1, 3).setValues([row]); else sh.appendRow(row);
  return { ok: true };
}

function adminSaveConfig(token, obj) {
  chkToken_(token, 'admin');
  const sh = sheet_(SH.CFG);
  const map = {};
  readAll_(SH.CFG).forEach(function (r) { map[String(r.key)] = r._row; });
  Object.keys(obj).forEach(function (k) {
    if (map[k]) sh.getRange(map[k], 2).setValue(String(obj[k]));
    else sh.appendRow([k, String(obj[k]), '']);
  });
  return { ok: true };
}

function adminChangePin(token, which, newPin) {
  chkToken_(token, 'admin');
  const p = String(newPin || '').trim();
  if (!/^[0-9]{4,8}$/.test(p)) return { ok: false, msg: 'PIN ต้องเป็นตัวเลข 4-8 หลัก' };
  PropertiesService.getScriptProperties().setProperty(which === 'staff' ? 'STAFF_PIN' : 'ADMIN_PIN', p);
  return { ok: true, msg: 'เปลี่ยน PIN เรียบร้อย' };
}

function adminReport(token, branch, limit) {
  chkToken_(token, 'admin');
  const b = String(branch || '').toUpperCase();
  const rows = readAll_(SH.LOG).filter(function (r) { return !b || String(r.branch_code).toUpperCase() === b; });

  const byPrize = {}, phones = {};
  rows.forEach(function (r) {
    const k = String(r.prize_name);
    if (!byPrize[k]) byPrize[k] = { spin: 0, redeem: 0 };
    byPrize[k].spin++;
    if (String(r.status) === 'REDEEMED') byPrize[k].redeem++;
    phones[String(r.phone)] = 1;
  });

  const recent = rows.slice(-(Number(limit) || 100)).reverse().map(function (r) {
    return { time: Utilities.formatDate(new Date(r.timestamp), TZ, 'dd/MM HH:mm'),
      branch: String(r.branch_code), phone: maskPhone_(r.phone), prize: String(r.prize_name),
      code: String(r.redeem_code), status: String(r.status) };
  });

  return { ok: true, totalSpin: rows.length,
    totalRedeem: rows.filter(function (r) { return String(r.status) === 'REDEEMED'; }).length,
    uniquePhone: Object.keys(phones).length, byPrize: byPrize, recent: recent };
}

function adminRestoreExpired(token) {
  chkToken_(token, 'admin');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const logSh = sheet_(SH.LOG);
    const cStatus = HEADERS.SpinLog.indexOf('status') + 1;
    const pmap = {};
    readAll_(SH.PRIZES).forEach(function (p) { pmap[String(p.branch_code).toUpperCase() + '|' + String(p.prize_id)] = p; });

    const now = new Date(); let n = 0;
    readAll_(SH.LOG).forEach(function (r) {
      if (String(r.status) !== 'PENDING') return;
      if (!r.expire_at || now <= new Date(r.expire_at)) return;
      logSh.getRange(r._row, cStatus).setValue('EXPIRED');
      const p = pmap[String(r.branch_code).toUpperCase() + '|' + String(r.prize_id)];
      if (p) {
        sheet_(SH.PRIZES).getRange(p._row, HEADERS.Prizes.indexOf('stock') + 1).setValue(Number(p.stock) + 1);
        p.stock = Number(p.stock) + 1;
      }
      n++;
    });
    return { ok: true, msg: 'คืนสต็อกจากโค้ดหมดอายุ ' + n + ' รายการ' };
  } finally { lock.releaseLock(); }
}

function adminResetQuota(token, phone) {
  chkToken_(token, 'admin');
  const ph = normPhone_(phone);
  if (!ph) return { ok: false, msg: 'เบอร์ไม่ถูกต้อง' };
  const sh = sheet_(SH.QUOTA);
  const rows = readAll_(SH.QUOTA).filter(function (r) { return String(r.phone) === ph; });
  rows.sort(function (a, b) { return b._row - a._row; }).forEach(function (r) { sh.deleteRow(r._row); });
  return { ok: true, msg: 'รีเซ็ตสิทธิ์เบอร์ ' + maskPhone_(ph) + ' แล้ว (' + rows.length + ' รายการ)' };
}
