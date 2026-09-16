/**
 * TSP_Monitor_Route.gs  (v2)
 * ------------------------------------------------------------------
 * Web App endpoint สำหรับจอมอนิเตอร์ ESP32
 *   GET {WEB_APP_URL}/exec?route=monitor&k=<token>&b=TPS-03[&h=<hash เดิม>]
 *
 * ⚠ ไฟล์นี้ต้องอยู่ใน **โปรเจกต์เดียวกับ TSP_Unified_Report.gs** เพราะใช้ของเดิมร่วมกัน:
 *     table_() · toDate_() · ymd_() · branchNames_() · getBranches_() · SH · C · CFG
 *   เพื่อให้ยอดบนจอ = ยอดในการ์ด LINE = ยอดใน PDF (ไม่มีตัวอ่านชีต 2 ชุด)
 *
 * ── ติดตั้ง ──────────────────────────────────────────────
 * 1) วางไฟล์นี้ + TSP_Monitor_Calc.gs ในโปรเจกต์ที่ผูกกับ TSP_Master_Database Ver.2
 * 2) ใน TSP_Unified_Report.gs เปลี่ยนชื่อ `function doGet()` เดิม เป็น
 *      `function TSP_dashboardRedirect_()`   (เนื้อในไม่ต้องแก้)
 *    ถ้าลืมข้อนี้ ระบบยังทำงานได้ — router จะสร้าง redirect จาก DASHBOARD_URL ให้เอง
 * 3) รัน TSP_monSetup() หนึ่งครั้ง (แก้ token ในฟังก์ชันก่อน)
 * 4) Deploy: Manage deployments → ✏️ → Version: **New version** → Deploy
 *    ⚠ ห้ามกด "New deployment" เพราะจะได้ URL ใหม่ แล้ว LINE webhook เดิมจะชี้ผิดที่
 * ------------------------------------------------------------------
 */

var TSP_MON_DEFAULTS = {
  TARGET_ID:    '1MmioNBkITy7HiZTHXJcJOQFzRW0eTUwam55jovaxnJ4',  // TSP_TargetBot_Config
  DATE_ORDER:   'MDY',       // ใช้เมื่อ Log_Date เป็นสตริงที่ชี้ขาดไม่ได้ (เช่น 2/8/2026)
  CACHE_SEC:    '90',
  SCHEDULE_URL: ''           // web app ตารางกะเดิม (เว้นว่าง = ปิดหน้า 5)
};

/** รันครั้งเดียวในตัวแก้ไข: ตั้ง token แล้วบันทึกค่าตั้งต้นลง Script Properties */
function TSP_monSetup() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('MONITOR_TOKEN', 'เปลี่ยนเป็นรหัสยาว ๆ ของคุณเอง');   // <<< แก้ก่อนรัน
  for (var k in TSP_MON_DEFAULTS) {
    if (TSP_MON_DEFAULTS.hasOwnProperty(k) && !props.getProperty(k)) props.setProperty(k, TSP_MON_DEFAULTS[k]);
  }
  Logger.log('ตั้งค่าเรียบร้อย: ' + JSON.stringify(props.getProperties()));
}

function TSP_monCfg_() {
  var p = PropertiesService.getScriptProperties().getProperties();
  var cfg = {};
  for (var k in TSP_MON_DEFAULTS) if (TSP_MON_DEFAULTS.hasOwnProperty(k)) cfg[k] = p[k] || TSP_MON_DEFAULTS[k];
  cfg.TOKEN = p.MONITOR_TOKEN || '';
  cfg.TZ = (typeof CFG === 'object' && CFG.TZ) ? CFG.TZ : 'Asia/Bangkok';
  return cfg;
}

/** เอนจินเดิม (TSP_Unified_Report.gs) ถูกโหลดมาด้วยหรือยัง */
function TSP_engineReady_() {
  return typeof table_ === 'function' && typeof SH === 'object' && typeof C === 'object';
}

/* ------------------------- อ่านข้อมูลผ่านเอนจินเดิม ------------------------- */

/**
 * แปลง Daily_Logs เป็นแถวที่ TSP_Monitor_Calc ใช้ได้ — อ่านผ่าน table_() ที่แคชไว้แล้ว
 * Log_Date แปลงด้วย toDate_() ของเอนจินเดิม เพื่อไม่ให้ตีความวันที่ต่างจากรายงาน
 * (dlIndex_() ใช้ไม่ได้ตรง ๆ เพราะไม่ได้ index คอลัมน์ Cash_Collected / Is_Collection / Timestamp)
 */
function TSP_monLogs_() {
  var t = table_(SH.DAILY);
  if (!t.exists) return [];
  var out = [];
  for (var i = 0; i < t.rows.length; i++) {
    var r = t.rows[i];
    var raw = g_(t, r, C.DL.DATE);
    out.push({
      Log_Date:       toDate_(raw) || raw,
      Branch_ID:      g_(t, r, C.DL.BRANCH),
      Machine_ID:     g_(t, r, C.DL.MACHINE),
      Revenue:        g_(t, r, C.DL.REV),
      Meter_Diff:     g_(t, r, C.DL.MDIFF),
      Cash_Collected: g_(t, r, C.DL.CASH),
      Is_Collection:  g_(t, r, C.DL.ISCOL),
      Timestamp:      g_(t, r, C.DL.TS)
    });
  }
  return out;
}

function TSP_monMachines_() {
  var t = table_(SH.MACHINES);
  if (!t.exists) return [];
  var out = [];
  for (var i = 0; i < t.rows.length; i++) {
    var r = t.rows[i];
    out.push({
      Machine_ID:     g_(t, r, C.MC.ID),
      Branch_ID:      g_(t, r, C.MC.BRANCH),
      Status:         g_(t, r, C.MC.STATUS),
      Revenue_Method: g_(t, r, C.MC.METHOD)
    });
  }
  return out;
}

/** ค่าเป้าจาก TSP_TargetBot_Config (คนละไฟล์ จึงอ่านตรง ไม่ผ่าน table_) */
function TSP_monTargets_(cfg) {
  var out = { targetRows: [], targetBranchRows: [], seasonalRows: [], milestones: [50, 80, 100, 110] };
  try {
    var tb = SpreadsheetApp.openById(cfg.TARGET_ID);
    out.targetRows       = TSP_monSheetRows_(tb, ['period', 'branch_code', 'target']);
    out.targetBranchRows = TSP_monSheetRows_(tb, ['branch_code', 'R0', 'active']);
    out.seasonalRows     = TSP_monSheetRows_(tb, ['month_no', 'seasonal_index']);
    var settings = TSP_monSheetRows_(tb, ['key', 'value']);
    for (var i = 0; i < settings.length; i++) {
      if (String(settings[i].key).trim() === 'MILESTONES') {
        var parts = String(settings[i].value).split(',')
          .map(function (s) { return Number(String(s).trim()); })
          .filter(function (n) { return isFinite(n) && n > 0; });
        if (parts.length) out.milestones = parts;
      }
    }
  } catch (e) {
    // ไม่มีสิทธิ์/ไฟล์เป้าหาย -> ยังแสดงยอดได้ แค่ไม่มี Progress Bar
  }
  return out;
}

/**
 * อ่านแท็บจาก "ลายเซ็นหัวคอลัมน์" — ใช้เฉพาะไฟล์ TargetBot ที่ชื่อแท็บไม่แน่นอน
 * (หัวตารางอาจไม่ได้อยู่แถวแรก เพราะมีแท็บ README นำหน้า)
 */
function TSP_monSheetRows_(ss, mustHave) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i], lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    var scan = Math.min(8, lastRow);
    var top = sh.getRange(1, 1, scan, lastCol).getDisplayValues();
    for (var r = 0; r < top.length; r++) {
      var hdr = top[r].map(function (v) { return String(v).trim(); }), ok = true;
      for (var m = 0; m < mustHave.length; m++) if (hdr.indexOf(mustHave[m]) < 0) { ok = false; break; }
      if (!ok) continue;
      if (lastRow <= r + 1) return [];
      var values = sh.getRange(r + 2, 1, lastRow - r - 1, lastCol).getValues(), out = [];
      for (var v = 0; v < values.length; v++) {
        var row = values[v], o = {}, empty = true;
        for (var c = 0; c < hdr.length; c++) {
          if (!hdr[c]) continue;
          o[hdr[c]] = row[c];
          if (row[c] !== '' && row[c] !== null) empty = false;
        }
        if (!empty) out.push(o);
      }
      return out;
    }
  }
  return [];
}

/**
 * ดึงตารางกะจาก web app เดิม แล้วบีบให้เหลือรูปแบบที่จอวาดได้
 * ยังไม่รู้โครงสร้างจริง -> ทนทั้ง array และ {rows:[...]}
 */
function TSP_monSchedule_(cfg, branchId) {
  if (!cfg.SCHEDULE_URL) return null;
  var cache = CacheService.getScriptCache(), ck = 'monsched:' + branchId;
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
    for (var i = 0; i < rows.length && i < 6; i++) {
      var r = rows[i];
      if (r instanceof Array) out.rows.push(r.slice(0, 8).map(String));
      else if (r && typeof r === 'object') {
        var line = [String(r.name || r.emp || r.Emp_Name || '')], days = r.days || r.shifts || [];
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

/* ------------------------------ payload ------------------------------ */

function TSP_monPayload_(branchId, cfg) {
  var names = branchNames_();
  var tg = TSP_monTargets_(cfg);        // เปิดไฟล์เป้าครั้งเดียวต่อการสร้าง payload
  return TSP_buildPayload_({
    branchId:   branchId,
    branchName: names[branchId] || names[TSP_normBranch_(branchId)] || '',
    now:        new Date(),
    dateOrder:  cfg.DATE_ORDER,
    logs:       TSP_monLogs_(),
    machines:   TSP_monMachines_(),
    targetRows:       tg.targetRows,
    targetBranchRows: tg.targetBranchRows,
    seasonalRows:     tg.seasonalRows,
    milestones:       tg.milestones,
    schedule:   TSP_monSchedule_(cfg, branchId)
  });
}

/* ------------------------------ endpoint ------------------------------ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.route === 'monitor') return TSP_monitorJson_(p);
  return TSP_dashboardOrDefault_();
}

/** พฤติกรรมเดิมของ doGet — เรียกของเดิมถ้ามี ไม่มีก็สร้าง redirect จาก DASHBOARD_URL ให้ */
function TSP_dashboardOrDefault_() {
  if (typeof TSP_dashboardRedirect_ === 'function') return TSP_dashboardRedirect_();
  var url = (typeof DASHBOARD_URL === 'string') ? DASHBOARD_URL : '';
  if (url) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<meta http-equiv="refresh" content="0;url=' + url + '">' +
      '</head><body style="font-family:sans-serif;padding:24px">' +
      'กำลังไปที่ Dashboard… <a href="' + url + '">คลิกที่นี่</a></body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutput('TSP Monitor API');
}

function TSP_monitorJson_(params) {
  var cfg = TSP_monCfg_();
  if (!cfg.TOKEN || params.k !== cfg.TOKEN) return TSP_monJson_({ e: 'auth' });
  if (!TSP_engineReady_()) return TSP_monJson_({ e: 'engine', msg: 'ต้องวางไฟล์นี้ในโปรเจกต์เดียวกับ TSP_Unified_Report.gs' });

  var branchId = params.b ? TSP_normBranch_(params.b) : 'ALL';
  if (String(params.b || '').toUpperCase() === 'ALL') branchId = 'ALL';

  var cache = CacheService.getScriptCache(), ckey = 'mon:v2:' + branchId;
  var json = cache.get(ckey);
  if (!json) {
    try {
      json = JSON.stringify(TSP_monPayload_(branchId, cfg));
      cache.put(ckey, json, Number(cfg.CACHE_SEC) || 90);
    } catch (err) {
      return TSP_monJson_({ e: 'calc', msg: String(err && err.message ? err.message : err) });
    }
  }
  // จอถือ hash เดิมอยู่ และข้อมูลไม่เปลี่ยน -> ตอบก้อนจิ๋ว (Apps Script ตั้ง header เองไม่ได้ จึงใช้ ETag/304 ไม่ได้)
  if (params.h) {
    try {
      var cur = JSON.parse(json);
      if (cur.h && cur.h === params.h) return TSP_monJson_({ nc: 1, ts: cur.ts });
    } catch (e2) { }
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function TSP_monJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ ทดสอบในตัวแก้ไข ------------------------------ */

/** ดู payload ของสาขาหนึ่งใน Execution log — ไม่ต้อง deploy */
function TSP_test_today() {
  var p = TSP_monPayload_('TPS-03', TSP_monCfg_());
  Logger.log(JSON.stringify(p, null, 2));
  return p;
}

/**
 * ตัวชี้ขาด — ยอดบนจอต้องเท่ากับที่รายงานเดิมคำนวณ ทุกบาท
 * เทียบ 3 จุด: MTD กับ sumRevenue_() · ยอดวันนี้กับ buildBranchData_() · 7 วันกับ lnTrend_()
 */
function TSP_test_vs_report() {
  if (!TSP_engineReady_()) { Logger.log('❌ ยังไม่ได้วางไฟล์ในโปรเจกต์เดียวกับ TSP_Unified_Report.gs'); return; }
  var branches = getBranches_(true), lines = [], allOk = true;

  for (var i = 0; i < branches.length; i++) {
    var b = branches[i];
    clearCache_();
    var p = TSP_monPayload_(b.id, TSP_monCfg_());

    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var mStart = new Date(now.getFullYear(), now.getMonth(), 1);

    var mtdRef = Math.round(sumRevenue_(b.id, mStart, today));
    var dayRef = Math.round(sumRevenue_(b.id, today, today));
    var trendRef = (typeof lnTrend_ === 'function')
      ? Math.round(lnTrend_(b.id, today, 7)[6].v) : null;

    var okM = (p.m.r === mtdRef), okD = (p.d.r === dayRef);
    var okT = (trendRef === null) || (p.w.v[6] === trendRef);
    if (!okM || !okD || !okT) allOk = false;

    lines.push([
      (okM && okD && okT) ? '✅' : '❌', b.id, b.name,
      'MTD จอ ' + p.m.r + ' / รายงาน ' + mtdRef,
      'วันนี้ จอ ' + p.d.r + ' / รายงาน ' + dayRef,
      trendRef === null ? '' : ('กราฟ จอ ' + p.w.v[6] + ' / lnTrend ' + trendRef),
      p.wn.length ? ('warn: ' + p.wn.join(',')) : ''
    ].join(' · '));
  }
  Logger.log(lines.join('\n'));
  Logger.log(allOk ? '✅ ตรงกันทุกสาขา' : '❌ มีสาขาที่ไม่ตรง — ห้าม deploy จนกว่าจะแก้');
  return allOk;
}

/** ตรวจว่าเอนจินเดิมพร้อมและอ่านคอลัมน์ที่จอต้องใช้ได้ครบ */
function TSP_test_columns() {
  if (!TSP_engineReady_()) { Logger.log('❌ เอนจินเดิมยังไม่ถูกโหลด'); return; }
  var t = table_(SH.DAILY);
  [C.DL.DATE, C.DL.BRANCH, C.DL.MACHINE, C.DL.REV, C.DL.MDIFF, C.DL.CASH, C.DL.ISCOL, C.DL.TS]
    .forEach(function (c) { Logger.log((t.idx[c] === undefined ? '❌ ไม่พบ ' : '✅ ') + c); });
  var mc = table_(SH.MACHINES), methods = {};
  for (var i = 0; i < mc.rows.length; i++) {
    var v = String(g_(mc, mc.rows[i], C.MC.METHOD) || '(ว่าง)').toUpperCase();
    methods[v] = (methods[v] || 0) + 1;
  }
  Logger.log('Revenue_Method: ' + JSON.stringify(methods));
}
