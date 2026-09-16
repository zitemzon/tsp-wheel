/**
 * TSP_Monitor_API.gs
 * ------------------------------------------------------------------
 * Web App endpoint สำหรับจอมอนิเตอร์ ESP32
 *   GET  ?k=<token>&b=TPS-03[&h=<hash เดิมของจอ>][&debug=1]
 *   ->   JSON สรุปยอด วัน/สัปดาห์/เดือน/เช็คพอยต์/เป้า/pace
 *
 * ข้อจำกัดที่ต้องรู้: Apps Script Web App ตั้ง HTTP header เองไม่ได้
 * จึงใช้ ETag/304 ไม่ได้ -> ใช้กลไก hash แทน
 *   จอส่ง ?h=<hash ที่ถือไว้> ถ้าข้อมูลไม่เปลี่ยน server ตอบ {"nc":1} ก้อนจิ๋ว
 *
 * ตั้งค่าครั้งแรก: รัน TSP_setup() ในตัวแก้ไขสคริปต์ 1 ครั้ง (แก้ค่าในฟังก์ชันก่อน)
 * Deploy: Execute as = Me, Who has access = Anyone
 * ------------------------------------------------------------------
 */

var TSP_DEFAULTS = {
  MASTER_ID: '1FU1iosNzzepz8AVAFOz-ZRwzPzi_O5wgiOV3TYwcomk',   // TSP_Master_Database Ver.2
  TARGET_ID: '1MmioNBkITy7HiZTHXJcJOQFzRW0eTUwam55jovaxnJ4',   // TSP_TargetBot_Config
  TZ: 'Asia/Bangkok',
  DATE_ORDER: 'MDY',
  CACHE_SEC: '90',
  SCHEDULE_URL: ''                                             // web app ตารางกะเดิม (เว้นว่าง = ปิดหน้า 5)
};

/** รันครั้งเดียวในตัวแก้ไข: ตั้ง token แล้วบันทึกค่าลง Script Properties */
function TSP_setup() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('MONITOR_TOKEN', 'เปลี่ยนเป็นรหัสยาว ๆ ของคุณเอง');   // <<< แก้ก่อนรัน
  for (var k in TSP_DEFAULTS) if (TSP_DEFAULTS.hasOwnProperty(k) && !props.getProperty(k)) {
    props.setProperty(k, TSP_DEFAULTS[k]);
  }
  Logger.log('ตั้งค่าเรียบร้อย: ' + JSON.stringify(props.getProperties()));
}

function TSP_cfg_() {
  var p = PropertiesService.getScriptProperties().getProperties();
  var cfg = {};
  for (var k in TSP_DEFAULTS) if (TSP_DEFAULTS.hasOwnProperty(k)) cfg[k] = p[k] || TSP_DEFAULTS[k];
  cfg.TOKEN = p.MONITOR_TOKEN || '';
  return cfg;
}

/* ------------------------- อ่านชีตแบบทนชื่อแท็บเปลี่ยน ------------------------- */

/**
 * หาแท็บจาก "ลายเซ็นหัวคอลัมน์" แทนการอ้างชื่อแท็บ
 * เพราะชื่อแท็บใน Master DB เปลี่ยนได้ แต่หัวคอลัมน์เป็นสัญญากับ AppSheet
 * สแกนหัวตารางใน 8 แถวแรกของทุกแท็บ
 */
function TSP_findSheet_(ss, mustHave) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lastCol = sh.getLastColumn();
    if (lastCol < 1) continue;
    var scan = Math.min(8, sh.getLastRow());
    if (scan < 1) continue;
    var top = sh.getRange(1, 1, scan, lastCol).getDisplayValues();
    for (var r = 0; r < top.length; r++) {
      var hdr = top[r].map(function (v) { return String(v).trim(); });
      var ok = true;
      for (var m = 0; m < mustHave.length; m++) if (hdr.indexOf(mustHave[m]) < 0) { ok = false; break; }
      if (ok) return { sheet: sh, headerRow: r + 1, header: hdr };
    }
  }
  return null;
}

/** อ่านทั้งแท็บเป็น array ของ object (คีย์ = ชื่อคอลัมน์) */
function TSP_readTable_(ss, mustHave, label) {
  var found = TSP_findSheet_(ss, mustHave);
  if (!found) throw new Error('หาแท็บ ' + (label || mustHave.join('/')) + ' ไม่เจอ (ต้องมีคอลัมน์: ' + mustHave.join(', ') + ')');
  var sh = found.sheet;
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow <= found.headerRow) return [];
  var values = sh.getRange(found.headerRow + 1, 1, lastRow - found.headerRow, lastCol).getValues();
  var hdr = found.header;
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i], o = {}, empty = true;
    for (var c = 0; c < hdr.length; c++) {
      if (!hdr[c]) continue;
      o[hdr[c]] = row[c];
      if (row[c] !== '' && row[c] !== null) empty = false;
    }
    if (!empty) out.push(o);
  }
  return out;
}

/* ------------------------------ ตารางกะ (P2) ------------------------------ */

/**
 * ดึงตารางกะจาก web app เดิม แล้วบีบให้เหลือรูปแบบที่จอวาดได้
 * ยังไม่รู้โครงสร้างจริงของ web app นั้น -> ตัวนี้ทนทั้ง array และ {rows:[...]}
 * ผลลัพธ์: { hdr:['จ','อ',...], rows:[['ไอริน','D','D','OFF',...]] }
 */
function TSP_fetchSchedule_(cfg, branchId) {
  if (!cfg.SCHEDULE_URL) return null;
  var cache = CacheService.getScriptCache();
  var ck = 'sched:' + branchId;
  var hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }

  try {
    var url = cfg.SCHEDULE_URL + (cfg.SCHEDULE_URL.indexOf('?') >= 0 ? '&' : '?') +
      'b=' + encodeURIComponent(branchId) + '&format=json';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return null;
    var body = res.getContentText();
    if (body.indexOf('{') !== 0 && body.indexOf('[') !== 0) return null;   // เป็น HTML ไม่ใช่ JSON
    var data = JSON.parse(body);
    var rows = data.rows || data.schedule || (data instanceof Array ? data : []);
    var out = { hdr: data.hdr || data.header || ['จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส', 'อา'], rows: [] };
    for (var i = 0; i < rows.length && i < 8; i++) {
      var r = rows[i];
      if (r instanceof Array) out.rows.push(r.slice(0, 8).map(String));
      else if (r && typeof r === 'object') {
        var line = [String(r.name || r.emp || r.Emp_Name || '')];
        var days = r.days || r.shifts || [];
        for (var d = 0; d < 7; d++) line.push(String(days[d] || ''));
        out.rows.push(line);
      }
    }
    cache.put(ck, JSON.stringify(out), 600);
    return out;
  } catch (err) {
    return null;
  }
}

/* ------------------------------ ตัวสร้าง payload ------------------------------ */

function TSP_payloadFor_(branchId, cfg) {
  var master = SpreadsheetApp.openById(cfg.MASTER_ID);
  var logs = TSP_readTable_(master, ['Log_ID', 'Log_Date', 'Machine_ID', 'Revenue'], 'Daily_Logs');
  var machines = TSP_readTable_(master, ['Machine_ID', 'Coin_Price', 'Revenue_Method'], 'Machines');
  var branches = TSP_readTable_(master, ['Branch_ID', 'Branch_Name', 'Machine_Prefix'], 'Branches');

  var targetRows = [], targetBranchRows = [], seasonalRows = [], milestones = [50, 80, 100, 110];
  try {
    var tb = SpreadsheetApp.openById(cfg.TARGET_ID);
    targetRows = TSP_readTable_(tb, ['period', 'branch_code', 'target'], 'Target');
    targetBranchRows = TSP_readTable_(tb, ['branch_code', 'R0', 'active'], 'Branch');
    seasonalRows = TSP_readTable_(tb, ['month_no', 'seasonal_index'], 'SeasonalIndex');
    var settings = TSP_readTable_(tb, ['key', 'value'], 'Settings');
    for (var i = 0; i < settings.length; i++) {
      if (String(settings[i].key).trim() === 'MILESTONES') {
        var parts = String(settings[i].value).split(',').map(function (s) { return Number(s.trim()); })
          .filter(function (n) { return isFinite(n) && n > 0; });
        if (parts.length) milestones = parts;
      }
    }
  } catch (e) {
    // ไม่มีสิทธิ์/ไม่มีไฟล์เป้า -> ยังแสดงยอดได้ แค่ไม่มี Progress Bar
  }

  return TSP_buildPayload_({
    branchId: branchId,
    now: new Date(),
    dateOrder: cfg.DATE_ORDER,
    logs: logs, machines: machines, branches: branches,
    targetRows: targetRows, targetBranchRows: targetBranchRows, seasonalRows: seasonalRows,
    milestones: milestones,
    schedule: TSP_fetchSchedule_(cfg, branchId)
  });
}

/* ------------------------------ endpoint ------------------------------ */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var cfg = TSP_cfg_();

  if (!cfg.TOKEN || params.k !== cfg.TOKEN) {
    return TSP_json_({ e: 'auth' });
  }

  var branchId = params.b ? String(params.b).toUpperCase() : 'ALL';
  var cache = CacheService.getScriptCache();
  var ckey = 'mon:v1:' + branchId;
  var json = cache.get(ckey);

  if (!json) {
    try {
      json = JSON.stringify(TSP_payloadFor_(branchId, cfg));
      cache.put(ckey, json, Number(cfg.CACHE_SEC) || 90);
    } catch (err) {
      return TSP_json_({ e: 'calc', msg: String(err && err.message ? err.message : err) });
    }
  }

  // จอถือ hash เดิมอยู่ และข้อมูลไม่เปลี่ยน -> ตอบก้อนจิ๋ว ประหยัดเวลา parse ฝั่งบอร์ด
  if (params.h) {
    try {
      var cur = JSON.parse(json);
      if (cur.h && cur.h === params.h) return TSP_json_({ nc: 1, ts: cur.ts });
    } catch (e2) { }
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function TSP_json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ ใช้ทดสอบในตัวแก้ไข ------------------------------ */

/** รันแล้วดูผลใน Execution log — ไม่ต้อง deploy */
function TSP_test_today() {
  var cfg = TSP_cfg_();
  var p = TSP_payloadFor_('TPS-03', cfg);
  Logger.log(JSON.stringify(p, null, 2));
  return p;
}

/** ตรวจว่าหาแท็บเจอครบไหม (รันก่อน deploy ทุกครั้งที่มีคนแก้โครงชีต) */
function TSP_test_sheets() {
  var cfg = TSP_cfg_();
  var master = SpreadsheetApp.openById(cfg.MASTER_ID);
  var checks = [
    ['Daily_Logs', ['Log_ID', 'Log_Date', 'Machine_ID', 'Revenue']],
    ['Machines', ['Machine_ID', 'Coin_Price', 'Revenue_Method']],
    ['Branches', ['Branch_ID', 'Branch_Name', 'Machine_Prefix']]
  ];
  checks.forEach(function (c) {
    var f = TSP_findSheet_(master, c[1]);
    Logger.log(c[0] + ' -> ' + (f ? f.sheet.getName() + ' (หัวตารางแถว ' + f.headerRow + ')' : '❌ ไม่เจอ'));
  });
  var tb = SpreadsheetApp.openById(cfg.TARGET_ID);
  [['Target', ['period', 'branch_code', 'target']], ['Branch', ['branch_code', 'R0', 'active']],
   ['SeasonalIndex', ['month_no', 'seasonal_index']]].forEach(function (c) {
    var f = TSP_findSheet_(tb, c[1]);
    Logger.log('TargetBot/' + c[0] + ' -> ' + (f ? f.sheet.getName() : '❌ ไม่เจอ'));
  });
}
