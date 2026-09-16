/**
 * TSP_Monitor_Calc.gs
 * ------------------------------------------------------------------
 * สูตรคำนวณล้วน ๆ สำหรับจอมอนิเตอร์รายได้ (TSP Revenue Monitor)
 *
 * ทุกฟังก์ชันในไฟล์นี้เป็น pure function — ไม่แตะ SpreadsheetApp / CacheService
 * รับเข้าเป็น array ของ object แล้วคืนค่าออกมาอย่างเดียว
 * เหตุผล: เอาไปรันทดสอบนอก Apps Script ได้ (ดู monitor/tests/calc.test.js)
 *
 * ฝั่ง I/O ทั้งหมดอยู่ใน TSP_Monitor_API.gs
 * ------------------------------------------------------------------
 */

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
 * คืน { unit: รหัสหน่วยเป้าฝั่ง TargetBot, members: [สาขาใน Master DB], shared: boolean }
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

/** คืนรายการหน่วยเป้าทั้งหมด (ใช้ตอน b=ALL) */
function TSP_allUnits_() {
  return ['TSP-01', 'TSP-02', 'TSP-03', 'TSP-04', 'TSP-06', 'TSP-57'];
}

/* ------------------------------ วันที่ ------------------------------ */

/**
 * แปลงค่าจากชีตเป็น Date (เที่ยงคืนตามเวลาไทย)
 * order = 'MDY' (ค่าเริ่มต้น — AppSheet เขียน Log_Date เป็น 07/28/2026) หรือ 'DMY'
 * คืน null เมื่อแปลงไม่ได้
 */
function TSP_toDate_(v, order) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    var a = Number(m[1]), b = Number(m[2]), y = Number(m[3]);
    var mm, dd;
    if (a > 12) { dd = a; mm = b; }                 // 28/07/2026 -> ชัดเจนว่า DMY
    else if (b > 12) { mm = a; dd = b; }            // 07/28/2026 -> ชัดเจนว่า MDY
    else if ((order || 'MDY') === 'DMY') { dd = a; mm = b; }
    else { mm = a; dd = b; }
    if (y > 2400) y -= 543;                         // เผื่อมีคนกรอกเป็น พ.ศ.
    return new Date(y, mm - 1, dd);
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** 'YYYY-MM-DD' */
function TSP_dkey_(d) {
  if (!d) return '';
  var mm = d.getMonth() + 1, dd = d.getDate();
  return d.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

function TSP_addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

function TSP_daysInMonth_(d) {
  return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
}

/** งวดแบบ TargetBot: พ.ศ.-MM เช่น 2569-09 */
function TSP_period_(d) {
  var mm = d.getMonth() + 1;
  return (d.getFullYear() + 543) + '-' + (mm < 10 ? '0' : '') + mm;
}

/* ------------------------------ รายได้ ------------------------------ */

/**
 * รายได้ของ log 1 แถว
 *  1) ถ้า Revenue > 0 เชื่อค่าที่ AppSheet คำนวณไว้
 *  2) ถ้าไม่มี ใช้ Meter_Diff x Coin_Price
 *  3) กัน first reading: Meter_Previous = 0 และไม่เคยมี log ก่อนหน้า -> 0
 *  4) Meter_Diff ติดลบ (มิเตอร์รีเซ็ต/เปลี่ยนบอร์ด) -> 0 + ใส่ warning
 * คืน { rev: number, warn: string|null }
 */
function TSP_rowRevenue_(row, coinPriceById, seenMachines) {
  var rev = Number(row.Revenue);
  if (isFinite(rev) && rev > 0) return { rev: rev, warn: null };

  var diff = Number(row.Meter_Diff);
  var cur = Number(row.Meter_Current);
  var prev = Number(row.Meter_Previous);
  if (!isFinite(diff) || diff === 0) {
    if (isFinite(cur) && isFinite(prev) && prev > 0) diff = cur - prev; else diff = 0;
  }
  var mid = row.Machine_ID;
  var firstEver = !(seenMachines && seenMachines[mid]);
  if ((!isFinite(prev) || prev === 0) && firstEver) {
    return { rev: 0, warn: 'first-reading:' + mid };   // อ่านมิเตอร์ครั้งแรกของตู้ ไม่นับเป็นรายได้
  }
  if (diff < 0) return { rev: 0, warn: 'meter-reset:' + mid };
  if (diff === 0) return { rev: 0, warn: null };

  var price = Number(coinPriceById && coinPriceById[mid]);
  if (!isFinite(price) || price <= 0) return { rev: 0, warn: 'no-coin-price:' + mid };
  return { rev: diff * price, warn: null };
}

function TSP_isTrue_(v) {
  if (v === true) return true;
  var s = String(v === null || v === undefined ? '' : v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'เก็บ' || s === '✔' || s === '✓';
}

/**
 * รวมรายได้รายวัน
 * logs   = แถวจากแท็บ Daily_Logs (object ตามชื่อคอลัมน์)
 * คืน { byDate: {'YYYY-MM-DD': บาท}, byDateCash: {...}, rows: [...], warns: [...] }
 * เฉพาะสาขาใน branchSet เท่านั้น (branchSet = null คือเอาทุกสาขา)
 */
function TSP_aggregate_(logs, coinPriceById, branchSet, dateOrder) {
  var byDate = {}, byDateCash = {}, warns = {}, seen = {}, kept = [];
  var sorted = logs.slice().sort(function (a, b) {
    var da = TSP_toDate_(a.Log_Date, dateOrder), db = TSP_toDate_(b.Log_Date, dateOrder);
    return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
  });

  for (var i = 0; i < sorted.length; i++) {
    var r = sorted[i];
    var bid = TSP_normBranch_(r.Branch_ID);
    if (branchSet && branchSet.indexOf(bid) < 0) continue;
    var d = TSP_toDate_(r.Log_Date, dateOrder);
    if (!d) { warns['bad-date'] = 1; continue; }
    var k = TSP_dkey_(d);

    var res = TSP_rowRevenue_(r, coinPriceById, seen);
    if (res.warn) warns[res.warn.split(':')[0]] = (warns[res.warn.split(':')[0]] || 0) + 1;
    seen[r.Machine_ID] = 1;

    byDate[k] = (byDate[k] || 0) + res.rev;

    if (TSP_isTrue_(r.Is_Collection)) {
      var cash = Number(r.Cash_Collected);
      if (isFinite(cash) && cash !== 0) byDateCash[k] = (byDateCash[k] || 0) + cash;
    }
    kept.push({ d: d, k: k, rev: res.rev, row: r, bid: bid });
  }
  return { byDate: byDate, byDateCash: byDateCash, rows: kept, warns: warns };
}

/**
 * จัดกลุ่มแถวเก็บเงินเป็น "รอบเช็คพอยต์"
 * แถวที่ Is_Collection = TRUE และเวลาห่างกันไม่เกิน gapMin นาที = รอบเดียวกัน
 * คืน [{ t:'HH:mm', a: ยอดเงินสดรวมรอบนั้น, n: จำนวนตู้ }]
 */
function TSP_clusterCheckpoints_(rows, gapMin) {
  var gap = (gapMin || 45) * 60000;
  var pts = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i].row || rows[i];
    if (!TSP_isTrue_(r.Is_Collection)) continue;
    var ts = r.Timestamp instanceof Date ? r.Timestamp : new Date(String(r.Timestamp));
    if (isNaN(ts.getTime())) ts = rows[i].d || null;
    if (!ts) continue;
    var cash = Number(r.Cash_Collected);
    pts.push({ ts: ts.getTime(), a: isFinite(cash) ? cash : 0 });
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
    var d = new Date(c.first);
    var hh = d.getHours(), mi = d.getMinutes();
    return { t: (hh < 10 ? '0' : '') + hh + ':' + (mi < 10 ? '0' : '') + mi, a: Math.round(c.a), n: c.n };
  });
}

/* ------------------------------ เป้า / pace ------------------------------ */

/** % แบบปลอดหารศูนย์ ทศนิยม 1 ตำแหน่ง */
function TSP_pct_(actual, target) {
  var t = Number(target);
  if (!isFinite(t) || t <= 0) return null;
  return Math.round((Number(actual) / t) * 1000) / 10;
}

/**
 * คาดการณ์ยอดปิดเดือน — MTD / วันที่ผ่านไป x จำนวนวันในเดือน
 * elapsedDays นับรวมวันปัจจุบัน (วันที่ 16 = ผ่านมา 16 วัน)
 */
function TSP_pace_(mtd, elapsedDays, daysInMonth) {
  var e = Number(elapsedDays);
  if (!isFinite(e) || e <= 0) return 0;
  return Math.round((Number(mtd) / e) * Number(daysInMonth));
}

/** สถานะ pace: G >= 100%, Y 90-99.9%, R < 90% */
function TSP_paceStatus_(forecast, target) {
  var p = TSP_pct_(forecast, target);
  if (p === null) return 'N';
  if (p >= 100) return 'G';
  if (p >= 90) return 'Y';
  return 'R';
}

/** milestone ล่าสุดที่ผ่านแล้ว จากรายการ เช่น [50,80,100,110] */
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
 * ถ้าไม่เจอ fallback = R0 x seasonal_index
 */
function TSP_findTarget_(targetRows, branchRows, seasonalRows, unit, period, monthNo) {
  var found = null;
  for (var i = 0; i < targetRows.length; i++) {
    var r = targetRows[i];
    if (String(r.period).trim() === period && TSP_normBranch_(r.branch_code) === TSP_normBranch_(unit)) found = r;
  }
  if (found) {
    var t = Number(found.target);
    if (isFinite(t) && t > 0) return { target: t, src: 'Target' };
  }
  var r0 = null;
  for (var j = 0; j < (branchRows || []).length; j++) {
    if (TSP_normBranch_(branchRows[j].branch_code) === TSP_normBranch_(unit)) r0 = Number(branchRows[j].R0);
  }
  var idx = null;
  for (var k = 0; k < (seasonalRows || []).length; k++) {
    if (Number(seasonalRows[k].month_no) === Number(monthNo)) idx = Number(seasonalRows[k].seasonal_index);
  }
  if (isFinite(r0) && r0 > 0 && isFinite(idx) && idx > 0) {
    return { target: Math.round(r0 * idx), src: 'R0xIndex' };
  }
  return { target: 0, src: 'none' };
}

/* ------------------------------ hash ------------------------------ */

/** FNV-1a 32-bit -> hex 8 ตัว ใช้ให้จอเช็คว่าข้อมูลเปลี่ยนหรือยัง (ทำงานได้ทั้งใน GAS และ node) */
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
 *   logs:[], machines:[], branches:[],
 *   targetRows:[], targetBranchRows:[], seasonalRows:[],
 *   milestones:[50,80,100,110], schedule: null|object
 * }
 */
function TSP_buildPayload_(ctx) {
  var now = ctx.now || new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var order = ctx.dateOrder || 'MDY';

  var coinPrice = {};
  var machineCount = 0;
  for (var i = 0; i < (ctx.machines || []).length; i++) {
    var m = ctx.machines[i];
    if (!m.Machine_ID) continue;
    coinPrice[m.Machine_ID] = Number(m.Coin_Price) || 0;
  }

  var isAll = !ctx.branchId || String(ctx.branchId).toUpperCase() === 'ALL';
  var unitInfo = isAll ? { unit: 'ALL', members: null, shared: false } : TSP_targetUnit_(ctx.branchId);
  var members = unitInfo.members;

  for (var mi = 0; mi < (ctx.machines || []).length; mi++) {
    var mm = ctx.machines[mi];
    var bidm = TSP_normBranch_(mm.Branch_ID);
    if (String(mm.Status || '').trim().toLowerCase() !== 'active') continue;
    if (!members || members.indexOf(bidm) >= 0) machineCount++;
  }

  var agg = TSP_aggregate_(ctx.logs || [], coinPrice, members, order);

  // ---- วันนี้ ----
  var kToday = TSP_dkey_(today);
  var revToday = agg.byDate[kToday] || 0;
  var logsToday = 0;
  for (var a = 0; a < agg.rows.length; a++) if (agg.rows[a].k === kToday) logsToday++;

  // ---- สัปดาห์: 7 วันจบที่วันนี้ + สัปดาห์ก่อนหน้า ----
  var week = [], wSum = 0, pSum = 0;
  for (var w = 6; w >= 0; w--) {
    var dk = TSP_dkey_(TSP_addDays_(today, -w));
    var v = agg.byDate[dk] || 0;
    week.push(Math.round(v));
    wSum += v;
  }
  for (var p = 13; p >= 7; p--) pSum += agg.byDate[TSP_dkey_(TSP_addDays_(today, -p))] || 0;

  // ---- เดือน (MTD) ----
  var mtd = 0, mtdCash = 0;
  var mStart = new Date(today.getFullYear(), today.getMonth(), 1);
  for (var dd = 0; dd < TSP_daysInMonth_(today); dd++) {
    var dkey = TSP_dkey_(TSP_addDays_(mStart, dd));
    mtd += agg.byDate[dkey] || 0;
    mtdCash += agg.byDateCash[dkey] || 0;
  }
  var dim = TSP_daysInMonth_(today);
  var elapsed = today.getDate();
  var period = TSP_period_(today);

  var tgt = { target: 0, src: 'none' };
  if (isAll) {
    var units = TSP_allUnits_(), sum = 0, any = false;
    for (var u = 0; u < units.length; u++) {
      var t1 = TSP_findTarget_(ctx.targetRows || [], ctx.targetBranchRows || [], ctx.seasonalRows || [], units[u], period, today.getMonth() + 1);
      if (t1.target > 0) { sum += t1.target; any = true; }
    }
    tgt = { target: sum, src: any ? 'sum' : 'none' };
  } else {
    tgt = TSP_findTarget_(ctx.targetRows || [], ctx.targetBranchRows || [], ctx.seasonalRows || [], unitInfo.unit, period, today.getMonth() + 1);
  }

  var tgtDay = tgt.target > 0 ? Math.round(tgt.target / dim) : 0;
  var forecast = TSP_pace_(mtd, elapsed, dim);
  var pctMonth = TSP_pct_(mtd, tgt.target);

  // ---- เช็คพอยต์ของวันนี้ ----
  var todayRows = agg.rows.filter(function (r) { return r.k === kToday; });
  var cps = TSP_clusterCheckpoints_(todayRows, 45);
  var cashToday = agg.byDateCash[kToday] || 0;
  var diffCash = Math.round(cashToday - revToday);
  var cashFlag = (revToday > 0 && cashToday > 0 && Math.abs(diffCash) / revToday > 0.03) ? 1 : 0;

  // ---- ชื่อสาขา ----
  var bname = '';
  for (var bi = 0; bi < (ctx.branches || []).length; bi++) {
    if (TSP_normBranch_(ctx.branches[bi].Branch_ID) === TSP_normBranch_(ctx.branchId)) bname = String(ctx.branches[bi].Branch_Name || '');
  }
  if (isAll) bname = 'ทุกสาขา';

  var warns = [];
  for (var wk in agg.warns) if (agg.warns.hasOwnProperty(wk)) warns.push(wk + ':' + agg.warns[wk]);
  if (tgt.src === 'none') warns.push('no-target');
  if (unitInfo.shared) warns.push('target-shared:' + unitInfo.unit);

  var payload = {
    v: 1,
    b: isAll ? 'ALL' : TSP_normBranch_(ctx.branchId),
    bn: bname,
    nm: machineCount,
    d: { r: Math.round(revToday), t: tgtDay, p: TSP_pct_(revToday, tgtDay), n: logsToday },
    c: { l: cps.slice(0, 6), cash: Math.round(cashToday), df: diffCash, fl: cashFlag },
    w: {
      v: week,
      s: Math.round(wSum),
      pv: Math.round(pSum),
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
    TSP_normBranch_: TSP_normBranch_, TSP_targetUnit_: TSP_targetUnit_, TSP_allUnits_: TSP_allUnits_,
    TSP_toDate_: TSP_toDate_, TSP_dkey_: TSP_dkey_, TSP_addDays_: TSP_addDays_,
    TSP_daysInMonth_: TSP_daysInMonth_, TSP_period_: TSP_period_,
    TSP_rowRevenue_: TSP_rowRevenue_, TSP_isTrue_: TSP_isTrue_, TSP_aggregate_: TSP_aggregate_,
    TSP_clusterCheckpoints_: TSP_clusterCheckpoints_, TSP_pct_: TSP_pct_, TSP_pace_: TSP_pace_,
    TSP_paceStatus_: TSP_paceStatus_, TSP_milestone_: TSP_milestone_, TSP_findTarget_: TSP_findTarget_,
    TSP_hash_: TSP_hash_, TSP_buildPayload_: TSP_buildPayload_
  };
}
