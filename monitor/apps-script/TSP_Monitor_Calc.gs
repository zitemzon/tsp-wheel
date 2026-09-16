/**
 * TSP_Monitor_Calc.gs  (v2 — ต่อกับเอนจินเดิม)
 * ------------------------------------------------------------------
 * สูตรคำนวณล้วนสำหรับจอมอนิเตอร์รายได้
 * ทุกฟังก์ชันเป็น pure function ไม่แตะ SpreadsheetApp/CacheService
 * → รันทดสอบนอก Apps Script ได้: monitor/tests/calc.test.js
 *
 * กติกาในไฟล์นี้ยึดตามระบบเดิมเป๊ะ เพื่อให้ยอดบนจอ = ยอดในการ์ด LINE = ยอดใน PDF
 *   - รายได้ = ผลรวมคอลัมน์ Revenue ล้วน            (TSP_Unified_Report.gs > buildBranchData_)
 *   - แถวที่ "ต้องมีเงินสด" = OR(Is_Collection, Revenue_Method <> 'METER')   (Code.gs > shouldHaveCash_)
 *   - ปี พ.ศ. ที่หลุดมาจาก AppSheet ต้องลบ 543                              (Code.gs > fixBuddhistYear_)
 *
 * ฝั่ง I/O อยู่ใน TSP_Monitor_Route.gs
 * ------------------------------------------------------------------
 */

/* ------------------------------ ค่าพื้นฐาน ------------------------------ */

/** แปลงเป็นตัวเลขแบบเดียวกับ n_() ของ TSP_Unified_Report ("1,250" -> 1250) */
function TSP_num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var x = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(x) ? 0 : x;
}

function TSP_isTrue_(v) {
  if (v === true) return true;
  var s = String(v === null || v === undefined ? '' : v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'เก็บ' || s === '✔' || s === '✓';
}

/** แปลงรหัสสาขาให้เป็นรูปแบบเดียวกัน — Master DB ใช้ TPS-xx, TargetBot ใช้ TSP-xx */
function TSP_normBranch_(id) {
  if (id === null || id === undefined) return '';
  var s = String(id).trim().toUpperCase().replace(/\s+/g, '');
  var m = s.match(/^(?:TPS|TSP)[-_]?([0-9]+[A-Z]?)$/);
  return m ? 'TPS-' + m[1] : s;
}

/**
 * ผังหน่วยเป้า — TargetBot ยุบบางสาขาเข้าด้วยกัน
 *   TPS-06A + TPS-06B -> TSP-06
 *   TPS-05  + TPS-07  -> TSP-57
 */
function TSP_targetUnit_(branchId) {
  var b = TSP_normBranch_(branchId);
  var UNITS = [
    { unit: 'TSP-01', members: ['TPS-01'] },
    { unit: 'TSP-02', members: ['TPS-02'] },
    { unit: 'TSP-03', members: ['TPS-03'] },
    { unit: 'TSP-04', members: ['TPS-04'] },
    { unit: 'TSP-06', members: ['TPS-06A', 'TPS-06B'] },
    { unit: 'TSP-57', members: ['TPS-05', 'TPS-07'] }
  ];
  for (var i = 0; i < UNITS.length; i++) {
    if (UNITS[i].members.indexOf(b) >= 0) {
      return { unit: UNITS[i].unit, members: UNITS[i].members.slice(), shared: UNITS[i].members.length > 1 };
    }
  }
  return { unit: '', members: [b], shared: false };
}

function TSP_allUnits_() {
  return ['TSP-01', 'TSP-02', 'TSP-03', 'TSP-04', 'TSP-06', 'TSP-57'];
}

/* ------------------------------ วันที่ ------------------------------ */

/**
 * ปีเกิน 2400 = เป็น พ.ศ. แน่นอน -> ลบ 543 (เก็บเวลาเดิมไว้ครบ)
 * พบจริง 30 ก.ค. 2026: AppSheet เขียน 2569 ลงชีต — Code.gs มี repairBuddhistDates() ไว้ซ่อมต้นเหตุ
 */
function TSP_fixBE_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return d;
  var y = d.getFullYear();
  if (y <= 2400) return d;
  return new Date(y - 543, d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
}

/**
 * แปลงค่าจากชีตเป็น Date (ตัดเวลาออก)
 * สตริง d/m/Y: ตัวไหนเกิน 12 ตัวนั้นคือวัน ถ้าไม่ชี้ขาดค่อยใช้ order ('MDY' ค่าเริ่มต้น)
 * หมายเหตุ: Unified_Report ตีความเป็น DMY ส่วน Code.gs ปล่อยให้ JS ตีความ (= MDY) — สองที่ไม่ตรงกันอยู่แล้ว
 */
function TSP_toDate_(v, order) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    var f = TSP_fixBE_(v);
    return new Date(f.getFullYear(), f.getMonth(), f.getDate());
  }
  if (typeof v === 'number' && v > 20000) {           // serial number ของสเปรดชีต
    var sd = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return new Date(sd.getUTCFullYear(), sd.getUTCMonth(), sd.getUTCDate());
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    var yy = Number(m[1]); if (yy > 2400) yy -= 543;
    return new Date(yy, Number(m[2]) - 1, Number(m[3]));
  }
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    var a = Number(m[1]), b = Number(m[2]), y = Number(m[3]), mm, dd;
    if (a > 12) { dd = a; mm = b; }
    else if (b > 12) { mm = a; dd = b; }
    else if ((order || 'MDY') === 'DMY') { dd = a; mm = b; }
    else { mm = a; dd = b; }
    if (y > 2400) y -= 543;
    return new Date(y, mm - 1, dd);
  }
  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  d = TSP_fixBE_(d);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** เหมือน TSP_toDate_ แต่เก็บเวลาไว้ — ใช้กับ Timestamp ตอนจัดกลุ่มรอบเก็บเงิน */
function TSP_toDateTime_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : TSP_fixBE_(v);
  if (typeof v === 'number' && v > 20000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
  var s = String(v).trim();
  var d = new Date(s);
  if (isNaN(d.getTime())) {
    var only = TSP_toDate_(s);                       // อ่านเวลาไม่ได้ ใช้แค่วันที่
    return only || null;
  }
  return TSP_fixBE_(d);
}

function TSP_dkey_(d) {
  if (!d) return '';
  var mm = d.getMonth() + 1, dd = d.getDate();
  return d.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

function TSP_addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function TSP_daysInMonth_(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate(); }

/** งวดแบบ TargetBot: พ.ศ.-MM เช่น 2569-09 */
function TSP_period_(d) {
  var mm = d.getMonth() + 1;
  return (d.getFullYear() + 543) + '-' + (mm < 10 ? '0' : '') + mm;
}

/* ------------------------------ รายได้ / เงินสด ------------------------------ */

/**
 * รายได้ของ log 1 แถว = คอลัมน์ Revenue ล้วน (ตรงกับ buildBranchData_ ของรายงานเดิม)
 * ไม่มี fallback Meter_Diff x Coin_Price — ถ้าใส่ จอจะโชว์มากกว่าการ์ด LINE
 * แต่ยังยกธง rev-missing ให้เห็นว่ามีแถวที่มิเตอร์เดินแต่ Revenue ไม่ถูกคำนวณ
 */
function TSP_rowRevenue_(row) {
  var rev = TSP_num_(row.Revenue);
  if (rev !== 0) return { rev: rev, warn: null };
  var diff = TSP_num_(row.Meter_Diff);
  if (diff > 0) return { rev: 0, warn: 'rev-missing:' + (row.Machine_ID || '?') };
  return { rev: 0, warn: null };
}

/**
 * แถวนี้ "ต้องมีเงินที่นับได้" ไหม — ต้องตรงกับ Show_If / Require_If ของ Cash_Collected ใน AppSheet:
 *     OR([Is_Collection], [Machine_ID].[Revenue_Method] <> "METER")
 * ห้ามใช้ Is_Collection เดี่ยว ๆ เพราะตู้ CAPSULE ถูกซ่อนช่องติ๊กไว้ จึงเป็น FALSE เสมอ
 */
function TSP_shouldHaveCash_(row, revenueMethod) {
  var method = String(revenueMethod || '').trim().toUpperCase();
  if (method && method !== 'METER') return true;
  return TSP_isTrue_(row.Is_Collection);
}

/**
 * รวมรายได้รายวัน + เตรียมแถวเงินสดไว้ให้ตัวจัดกลุ่มเช็คพอยต์
 * logs        = แถวจาก Daily_Logs (object ตามชื่อคอลัมน์)
 * methodById  = { Machine_ID: 'METER' | 'CAPSULE' }
 * branchSet   = รายการสาขาที่สนใจ (null = ทุกสาขา)
 */
function TSP_aggregate_(logs, methodById, branchSet, dateOrder) {
  var byDate = {}, byDateCash = {}, warns = {}, kept = [];
  var bump = function (code) { if (code) { var k = code.split(':')[0]; warns[k] = (warns[k] || 0) + 1; } };

  for (var i = 0; i < logs.length; i++) {
    var r = logs[i];
    var bid = TSP_normBranch_(r.Branch_ID);
    if (branchSet && branchSet.indexOf(bid) < 0) continue;

    var d = TSP_toDate_(r.Log_Date, dateOrder);
    if (!d) { bump('bad-date'); continue; }
    var k = TSP_dkey_(d);

    var res = TSP_rowRevenue_(r);
    bump(res.warn);
    byDate[k] = (byDate[k] || 0) + res.rev;

    var cash = null;
    if (TSP_shouldHaveCash_(r, methodById ? methodById[r.Machine_ID] : '')) {
      var amt = TSP_num_(r.Cash_Collected);
      if (amt > 0) { cash = amt; byDateCash[k] = (byDateCash[k] || 0) + amt; }
      else bump('cash-missing');                     // ควรมีเงินสดแต่ยังไม่ได้กรอก
    }

    kept.push({ d: d, k: k, rev: res.rev, bid: bid, cash: cash, ts: TSP_toDateTime_(r.Timestamp) || d });
  }

  kept.sort(function (a, b) { return a.d.getTime() - b.d.getTime(); });
  return { byDate: byDate, byDateCash: byDateCash, rows: kept, warns: warns };
}

/**
 * จัดกลุ่มแถวเก็บเงินเป็น "รอบเช็คพอยต์"
 * แถวที่มีเงินสด และเวลาห่างกันไม่เกิน gapMin นาที = รอบเดียวกัน
 * คืน [{ t:'HH:mm', a: ยอดรวมรอบนั้น, n: จำนวนตู้ }]
 */
function TSP_clusterCheckpoints_(entries, gapMin) {
  var gap = (gapMin || 45) * 60000;
  var pts = [];
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    if (!e.cash || !e.ts) continue;
    pts.push({ ts: e.ts.getTime(), a: e.cash });
  }
  pts.sort(function (a, b) { return a.ts - b.ts; });

  var out = [];
  for (var j = 0; j < pts.length; j++) {
    var last = out[out.length - 1];
    if (last && pts[j].ts - last.last <= gap) {
      last.a += pts[j].a; last.n += 1; last.last = pts[j].ts;
    } else {
      out.push({ first: pts[j].ts, last: pts[j].ts, a: pts[j].a, n: 1 });
    }
  }
  return out.map(function (c) {
    var d = new Date(c.first), hh = d.getHours(), mi = d.getMinutes();
    return { t: (hh < 10 ? '0' : '') + hh + ':' + (mi < 10 ? '0' : '') + mi, a: Math.round(c.a), n: c.n };
  });
}

/* ------------------------------ เป้า / pace ------------------------------ */

function TSP_pct_(actual, target) {
  var t = Number(target);
  if (!isFinite(t) || t <= 0) return null;
  return Math.round((Number(actual) / t) * 1000) / 10;
}

/** คาดการณ์ยอดปิดเดือน — MTD / วันที่ผ่านไป x จำนวนวันในเดือน */
function TSP_pace_(mtd, elapsedDays, daysInMonth) {
  var e = Number(elapsedDays);
  if (!isFinite(e) || e <= 0) return 0;
  return Math.round((Number(mtd) / e) * Number(daysInMonth));
}

/** G >= 100% · Y 90-99.9% · R < 90% · N ไม่มีเป้า */
function TSP_paceStatus_(forecast, target) {
  var p = TSP_pct_(forecast, target);
  if (p === null) return 'N';
  if (p >= 100) return 'G';
  if (p >= 90) return 'Y';
  return 'R';
}

function TSP_milestone_(pct, list) {
  if (pct === null || pct === undefined) return 0;
  var ms = (list || [50, 80, 100, 110]).slice().sort(function (a, b) { return a - b; });
  var hit = 0;
  for (var i = 0; i < ms.length; i++) if (pct >= ms[i]) hit = ms[i];
  return hit;
}

/**
 * หาเป้าของหน่วย/งวด จากแท็บ Target ของ TargetBot
 * แท็บนี้มีแถวซ้ำ (งวด 2569-08 ซ้ำ 3 ชุด) -> ยึดแถวท้ายสุดของคู่ (period, branch_code)
 * ไม่เจอ -> fallback R0 x seasonal_index
 */
function TSP_findTarget_(targetRows, branchRows, seasonalRows, unit, period, monthNo) {
  var found = null;
  for (var i = 0; i < targetRows.length; i++) {
    var r = targetRows[i];
    if (String(r.period).trim() === period && TSP_normBranch_(r.branch_code) === TSP_normBranch_(unit)) found = r;
  }
  if (found) {
    var t = TSP_num_(found.target);
    if (t > 0) return { target: t, src: 'Target' };
  }
  var r0 = 0;
  for (var j = 0; j < (branchRows || []).length; j++) {
    if (TSP_normBranch_(branchRows[j].branch_code) === TSP_normBranch_(unit)) r0 = TSP_num_(branchRows[j].R0);
  }
  var idx = 0;
  for (var k = 0; k < (seasonalRows || []).length; k++) {
    if (TSP_num_(seasonalRows[k].month_no) === Number(monthNo)) idx = TSP_num_(seasonalRows[k].seasonal_index);
  }
  if (r0 > 0 && idx > 0) return { target: Math.round(r0 * idx), src: 'R0xIndex' };
  return { target: 0, src: 'none' };
}

/* ------------------------------ hash ------------------------------ */

/** FNV-1a 32-bit -> hex 8 ตัว ใช้บอกจอว่าข้อมูลเปลี่ยนหรือยัง (ทำงานได้ทั้งใน GAS และ node) */
function TSP_hash_(str) {
  var h = 0x811c9dc5;
  for (var i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return ('0000000' + h.toString(16)).slice(-8);
}

/* ------------------------------ payload ------------------------------ */

/**
 * ประกอบ payload ทั้งก้อน — pure ล้วน
 * ctx = {
 *   branchId, now:Date, dateOrder,
 *   logs:[], machines:[], branchName:'',
 *   targetRows:[], targetBranchRows:[], seasonalRows:[],
 *   milestones:[50,80,100,110], schedule: null|object
 * }
 */
function TSP_buildPayload_(ctx) {
  var now = ctx.now || new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var order = ctx.dateOrder || 'MDY';

  var isAll = !ctx.branchId || String(ctx.branchId).toUpperCase() === 'ALL';
  var unitInfo = isAll ? { unit: 'ALL', members: null, shared: false } : TSP_targetUnit_(ctx.branchId);
  var members = unitInfo.members;

  var methodById = {}, machineCount = 0;
  for (var i = 0; i < (ctx.machines || []).length; i++) {
    var m = ctx.machines[i];
    if (!m.Machine_ID) continue;
    methodById[m.Machine_ID] = m.Revenue_Method;
    var st = String(m.Status || '').trim().toLowerCase();
    if (st && st !== 'active') continue;
    var bidm = TSP_normBranch_(m.Branch_ID);
    if (!members || members.indexOf(bidm) >= 0) machineCount++;
  }

  var agg = TSP_aggregate_(ctx.logs || [], methodById, members, order);

  // ---- วันนี้ ----
  var kToday = TSP_dkey_(today);
  var revToday = agg.byDate[kToday] || 0;
  var todayRows = [], logsToday = 0;
  for (var a = 0; a < agg.rows.length; a++) {
    if (agg.rows[a].k === kToday) { todayRows.push(agg.rows[a]); logsToday++; }
  }

  // ---- สัปดาห์: 7 วันจบที่วันนี้ + สัปดาห์ก่อนหน้า ----
  var week = [], wSum = 0, pSum = 0;
  for (var w = 6; w >= 0; w--) {
    var v = agg.byDate[TSP_dkey_(TSP_addDays_(today, -w))] || 0;
    week.push(Math.round(v));
    wSum += v;
  }
  for (var p = 13; p >= 7; p--) pSum += agg.byDate[TSP_dkey_(TSP_addDays_(today, -p))] || 0;

  // ---- เดือน (MTD) ----
  var mStart = new Date(today.getFullYear(), today.getMonth(), 1);
  var dim = TSP_daysInMonth_(today), mtd = 0, mtdCash = 0;
  for (var dd = 0; dd < dim; dd++) {
    var dkey = TSP_dkey_(TSP_addDays_(mStart, dd));
    mtd += agg.byDate[dkey] || 0;
    mtdCash += agg.byDateCash[dkey] || 0;
  }
  var elapsed = today.getDate();
  var period = TSP_period_(today);

  var tgt;
  if (isAll) {
    var units = TSP_allUnits_(), sum = 0, any = false;
    for (var u = 0; u < units.length; u++) {
      var t1 = TSP_findTarget_(ctx.targetRows || [], ctx.targetBranchRows || [], ctx.seasonalRows || [],
                               units[u], period, today.getMonth() + 1);
      if (t1.target > 0) { sum += t1.target; any = true; }
    }
    tgt = { target: sum, src: any ? 'sum' : 'none' };
  } else {
    tgt = TSP_findTarget_(ctx.targetRows || [], ctx.targetBranchRows || [], ctx.seasonalRows || [],
                          unitInfo.unit, period, today.getMonth() + 1);
  }

  var tgtDay = tgt.target > 0 ? Math.round(tgt.target / dim) : 0;
  var forecast = TSP_pace_(mtd, elapsed, dim);
  var pctMonth = TSP_pct_(mtd, tgt.target);

  // ---- เช็คพอยต์ของวันนี้ ----
  var cps = TSP_clusterCheckpoints_(todayRows, 45);
  var cashToday = agg.byDateCash[kToday] || 0;
  var diffCash = Math.round(cashToday - revToday);
  var cashFlag = (revToday > 0 && cashToday > 0 && Math.abs(diffCash) / revToday > 0.03) ? 1 : 0;

  var warns = [];
  for (var wk in agg.warns) if (agg.warns.hasOwnProperty(wk)) warns.push(wk + ':' + agg.warns[wk]);
  if (tgt.src === 'none') warns.push('no-target');
  if (unitInfo.shared) warns.push('target-shared:' + unitInfo.unit);

  var payload = {
    v: 1,
    b: isAll ? 'ALL' : TSP_normBranch_(ctx.branchId),
    bn: isAll ? 'ทุกสาขา' : String(ctx.branchName || ''),
    nm: machineCount,
    d: { r: Math.round(revToday), t: tgtDay, p: TSP_pct_(revToday, tgtDay), n: logsToday },
    c: { l: cps.slice(0, 6), cash: Math.round(cashToday), df: diffCash, fl: cashFlag },
    w: {
      v: week, s: Math.round(wSum), pv: Math.round(pSum),
      ch: pSum > 0 ? Math.round(((wSum - pSum) / pSum) * 1000) / 10 : null
    },
    m: {
      pd: period, r: Math.round(mtd), t: tgt.target, p: pctMonth,
      f: forecast, fs: TSP_paceStatus_(forecast, tgt.target),
      ms: TSP_milestone_(pctMonth, ctx.milestones), sh: unitInfo.shared ? 1 : 0,
      cash: Math.round(mtdCash), ed: elapsed, dm: dim
    },
    wn: warns
  };
  if (ctx.schedule) payload.s = ctx.schedule;

  payload.h = TSP_hash_(JSON.stringify(payload));
  payload.ts = TSP_dkey_(today) + ' ' +
    ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
  return payload;
}

/* export สำหรับรันทดสอบด้วย node — Apps Script จะไม่เข้าบล็อกนี้ */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    TSP_num_: TSP_num_, TSP_isTrue_: TSP_isTrue_,
    TSP_normBranch_: TSP_normBranch_, TSP_targetUnit_: TSP_targetUnit_, TSP_allUnits_: TSP_allUnits_,
    TSP_fixBE_: TSP_fixBE_, TSP_toDate_: TSP_toDate_, TSP_toDateTime_: TSP_toDateTime_,
    TSP_dkey_: TSP_dkey_, TSP_addDays_: TSP_addDays_, TSP_daysInMonth_: TSP_daysInMonth_,
    TSP_period_: TSP_period_, TSP_rowRevenue_: TSP_rowRevenue_, TSP_shouldHaveCash_: TSP_shouldHaveCash_,
    TSP_aggregate_: TSP_aggregate_, TSP_clusterCheckpoints_: TSP_clusterCheckpoints_,
    TSP_pct_: TSP_pct_, TSP_pace_: TSP_pace_, TSP_paceStatus_: TSP_paceStatus_,
    TSP_milestone_: TSP_milestone_, TSP_findTarget_: TSP_findTarget_,
    TSP_hash_: TSP_hash_, TSP_buildPayload_: TSP_buildPayload_
  };
}
