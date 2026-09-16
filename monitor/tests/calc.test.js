/**
 * ทดสอบสูตรใน TSP_Monitor_Calc.gs ด้วย node (ไม่ต้องมี Apps Script)
 *   รัน:  node monitor/tests/calc.test.js
 * ใช้ตัวเลขจริงจาก TSP_TargetBot_Config เป็นกรณีอ้างอิง
 */
const path = require('path');
const C = require(path.join(__dirname, '..', 'apps-script', 'TSP_Monitor_Calc.gs'));

let pass = 0, fail = 0;
function eq(label, got, want) {
  const g = JSON.stringify(got), w = JSON.stringify(want);
  if (g === w) { pass++; console.log('  ✓ ' + label); }
  else { fail++; console.log('  ✗ ' + label + '\n      ได้  : ' + g + '\n      ควรได้: ' + w); }
}

console.log('\n[1] รหัสสาขา / หน่วยเป้า');
eq('TPS-03 -> TPS-03', C.TSP_normBranch_('TPS-03'), 'TPS-03');
eq('TSP-03 -> TPS-03', C.TSP_normBranch_('TSP-03'), 'TPS-03');
eq('tps 06a -> TPS-06A', C.TSP_normBranch_(' tps06a '), 'TPS-06A');
eq('TPS-05 อยู่หน่วย TSP-57 (แชร์เป้า)', C.TSP_targetUnit_('TPS-05'), { unit: 'TSP-57', members: ['TPS-05', 'TPS-07'], shared: true });
eq('TPS-06B อยู่หน่วย TSP-06', C.TSP_targetUnit_('TPS-06B').unit, 'TSP-06');
eq('TPS-01 ไม่แชร์เป้า', C.TSP_targetUnit_('TPS-01').shared, false);

console.log('\n[2] วันที่');
eq('07/28/2026 (AppSheet MDY)', C.TSP_dkey_(C.TSP_toDate_('07/28/2026')), '2026-07-28');
eq('28/07/2026 (วันเกิน 12 = DMY อัตโนมัติ)', C.TSP_dkey_(C.TSP_toDate_('28/07/2026')), '2026-07-28');
eq('2/8/2026 ภายใต้ MDY = 8 ก.พ.', C.TSP_dkey_(C.TSP_toDate_('2/8/2026', 'MDY')), '2026-02-08');
eq('2/8/2026 ภายใต้ DMY = 2 ส.ค.', C.TSP_dkey_(C.TSP_toDate_('2/8/2026', 'DMY')), '2026-08-02');
eq('งวด TargetBot ของ ก.ย. 2026', C.TSP_period_(new Date(2026, 8, 16)), '2569-09');
eq('จำนวนวันเดือน ก.ย.', C.TSP_daysInMonth_(new Date(2026, 8, 16)), 30);

console.log('\n[3] รายได้ต่อแถว');
const price = { 'RCH-0001': 10, 'RCH-0002': 20 };
eq('เชื่อคอลัมน์ Revenue เมื่อมีค่า', C.TSP_rowRevenue_({ Machine_ID: 'RCH-0001', Revenue: 1250, Meter_Diff: 5 }, price, {}).rev, 1250);
eq('fallback = Meter_Diff x Coin_Price', C.TSP_rowRevenue_({ Machine_ID: 'RCH-0001', Revenue: 0, Meter_Diff: 12, Meter_Previous: 100 }, price, {}).rev, 120);
eq('อ่านมิเตอร์ครั้งแรก (prev=0) ไม่นับรายได้', C.TSP_rowRevenue_({ Machine_ID: 'RCH-0002', Revenue: 0, Meter_Current: 4593, Meter_Previous: 0, Meter_Diff: 0 }, price, {}).rev, 0);
eq('  ...และติด warning first-reading', C.TSP_rowRevenue_({ Machine_ID: 'RCH-0002', Revenue: 0, Meter_Previous: 0 }, price, {}).warn, 'first-reading:RCH-0002');
eq('มิเตอร์รีเซ็ต (diff ติดลบ) = 0', C.TSP_rowRevenue_({ Machine_ID: 'RCH-0001', Revenue: 0, Meter_Diff: -800, Meter_Previous: 900 }, price, {}).rev, 0);
eq('ตู้ไม่มีราคาเหรียญ = 0 + warning', C.TSP_rowRevenue_({ Machine_ID: 'XX-0001', Revenue: 0, Meter_Diff: 5, Meter_Previous: 10 }, price, {}).warn, 'no-coin-price:XX-0001');

console.log('\n[4] รวมยอด + เช็คพอยต์');
const machines = [
  { Machine_ID: 'RCH-0001', Branch_ID: 'TPS-03', Coin_Price: 10, Status: 'Active' },
  { Machine_ID: 'RCH-0002', Branch_ID: 'TPS-03', Coin_Price: 20, Status: 'Active' },
  { Machine_ID: 'GOC-0001', Branch_ID: 'TPS-01', Coin_Price: 10, Status: 'Active' }
];
const logs = [
  { Log_ID: 'a', Log_Date: '09/16/2026', Branch_ID: 'TPS-03', Machine_ID: 'RCH-0001', Meter_Previous: 100, Meter_Current: 200, Meter_Diff: 100, Revenue: 1000, Is_Collection: 'TRUE', Cash_Collected: 980, Timestamp: new Date(2026, 8, 16, 11, 20) },
  { Log_ID: 'b', Log_Date: '09/16/2026', Branch_ID: 'TPS-03', Machine_ID: 'RCH-0002', Meter_Previous: 50, Meter_Current: 70, Meter_Diff: 20, Revenue: 0, Is_Collection: 'TRUE', Cash_Collected: 400, Timestamp: new Date(2026, 8, 16, 11, 35) },
  { Log_ID: 'c', Log_Date: '09/16/2026', Branch_ID: 'TPS-03', Machine_ID: 'RCH-0001', Meter_Previous: 200, Meter_Current: 230, Meter_Diff: 30, Revenue: 300, Is_Collection: 'TRUE', Cash_Collected: 300, Timestamp: new Date(2026, 8, 16, 18, 5) },
  { Log_ID: 'd', Log_Date: '09/15/2026', Branch_ID: 'TPS-03', Machine_ID: 'RCH-0001', Meter_Previous: 80, Meter_Current: 100, Meter_Diff: 20, Revenue: 200 },
  { Log_ID: 'e', Log_Date: '09/16/2026', Branch_ID: 'TPS-01', Machine_ID: 'GOC-0001', Meter_Previous: 10, Meter_Current: 60, Meter_Diff: 50, Revenue: 500 }
];
const coin = { 'RCH-0001': 10, 'RCH-0002': 20, 'GOC-0001': 10 };
const agg = C.TSP_aggregate_(logs, coin, ['TPS-03'], 'MDY');
eq('ยอด 16/9 ของ TPS-03 = 1000 + (20x20) + 300', agg.byDate['2026-09-16'], 1700);
eq('ยอด 15/9', agg.byDate['2026-09-15'], 200);
eq('ไม่ปนสาขาอื่น (TPS-01 ถูกกรองออก)', Object.keys(agg.byDate).length, 2);
eq('เงินสดที่เก็บได้ 16/9', agg.byDateCash['2026-09-16'], 1680);
const cps = C.TSP_clusterCheckpoints_(agg.rows.filter(r => r.k === '2026-09-16'), 45);
eq('จับกลุ่มได้ 2 รอบ (11:20 กับ 18:05)', cps.length, 2);
eq('รอบแรก 11:20 รวม 1380 จาก 2 ตู้', cps[0], { t: '11:20', a: 1380, n: 2 });
eq('รอบสอง 18:05 รวม 300', cps[1], { t: '18:05', a: 300, n: 1 });

console.log('\n[5] เป้า / pace / milestone');
const targetRows = [
  { period: '2569-08', branch_code: 'TSP-01', target: 38500 },
  { period: '2569-08', branch_code: 'TSP-01', target: 38500 },
  { period: '2569-09', branch_code: 'TSP-01', target: 35800 },
  { period: '2569-09', branch_code: 'TSP-03', target: 91000 }
];
const tbBranch = [{ branch_code: 'TSP-01', R0: 55000, active: 'TRUE' }, { branch_code: 'TSP-02', R0: 63000, active: 'TRUE' }];
const seasonal = [{ month_no: 9, seasonal_index: 0.65 }, { month_no: 8, seasonal_index: 0.7 }];
eq('เป้า TSP-03 ก.ย. จากแท็บ Target', C.TSP_findTarget_(targetRows, tbBranch, seasonal, 'TSP-03', '2569-09', 9), { target: 91000, src: 'Target' });
eq('แถวซ้ำไม่ทำให้ยอดบวกซ้ำ', C.TSP_findTarget_(targetRows, tbBranch, seasonal, 'TSP-01', '2569-08', 8).target, 38500);
eq('ไม่มีในแท็บ Target -> R0 x index (63000 x 0.65)', C.TSP_findTarget_(targetRows, tbBranch, seasonal, 'TSP-02', '2569-09', 9), { target: 40950, src: 'R0xIndex' });
eq('ไม่มีข้อมูลเลย -> 0', C.TSP_findTarget_([], [], [], 'TSP-99', '2569-09', 9).target, 0);
eq('pace: 45,500 ใน 16 วัน ของเดือน 30 วัน', C.TSP_pace_(45500, 16, 30), 85313);
eq('pct 45,500 / 91,000', C.TSP_pct_(45500, 91000), 50);
eq('หารศูนย์ไม่ระเบิด', C.TSP_pct_(100, 0), null);
eq('pace status เขียว', C.TSP_paceStatus_(95000, 91000), 'G');
eq('pace status เหลือง (93.4%)', C.TSP_paceStatus_(85000, 91000), 'Y');
eq('pace status แดง', C.TSP_paceStatus_(60000, 91000), 'R');
eq('milestone ที่ผ่านแล้วของ 50%', C.TSP_milestone_(50, [50, 80, 100, 110]), 50);
eq('milestone ของ 99.9%', C.TSP_milestone_(99.9, [50, 80, 100, 110]), 80);
eq('milestone ของ 112%', C.TSP_milestone_(112, [50, 80, 100, 110]), 110);

console.log('\n[6] payload ทั้งก้อน');
const monthLogs = [];
for (let d = 1; d <= 16; d++) {
  monthLogs.push({
    Log_ID: 'm' + d, Log_Date: '09/' + String(d).padStart(2, '0') + '/2026', Branch_ID: 'TPS-03',
    Machine_ID: 'RCH-0001', Meter_Previous: 100 * d, Meter_Current: 100 * d + 100, Meter_Diff: 100, Revenue: 2000
  });
}
const p = C.TSP_buildPayload_({
  branchId: 'TPS-03', now: new Date(2026, 8, 16, 14, 32), dateOrder: 'MDY',
  logs: monthLogs, machines: machines, branches: [{ Branch_ID: 'TPS-03', Branch_Name: 'โรบินสันฉลอง', Machine_Prefix: 'RCH' }],
  targetRows: targetRows, targetBranchRows: tbBranch, seasonalRows: seasonal, milestones: [50, 80, 100, 110]
});
eq('ยอดวันนี้', p.d.r, 2000);
eq('เป้ารายวัน = 91,000 / 30', p.d.t, 3033);
eq('ยอดสัปดาห์ (7 วัน x 2,000)', p.w.s, 14000);
eq('สัปดาห์ก่อนหน้า', p.w.pv, 14000);
eq('MTD 16 วัน', p.m.r, 32000);
eq('% ของเป้าเดือน', p.m.p, 35.2);
eq('คาดการณ์ปิดเดือน', p.m.f, 60000);
eq('สถานะ pace = แดง', p.m.fs, 'R');
eq('จำนวนตู้ Active ของสาขา', p.nm, 2);
eq('ชื่อสาขา', p.bn, 'โรบินสันฉลอง');
eq('มีค่า hash', typeof p.h === 'string' && p.h.length === 8, true);
const p2 = C.TSP_buildPayload_({
  branchId: 'TPS-03', now: new Date(2026, 8, 16, 19, 59), dateOrder: 'MDY',
  logs: monthLogs, machines: machines, branches: [{ Branch_ID: 'TPS-03', Branch_Name: 'โรบินสันฉลอง' }],
  targetRows: targetRows, targetBranchRows: tbBranch, seasonalRows: seasonal
});
eq('hash ไม่เปลี่ยนเมื่อข้อมูลเท่าเดิม (เวลาเปลี่ยนไม่นับ)', p2.h, p.h);
const pShared = C.TSP_buildPayload_({
  branchId: 'TPS-05', now: new Date(2026, 8, 16, 14, 0), dateOrder: 'MDY',
  logs: [], machines: [], branches: [], targetRows: [], targetBranchRows: [], seasonalRows: []
});
eq('สาขาแชร์เป้าติดธง sh=1', pShared.m.sh, 1);
eq('และมี warning target-shared', pShared.wn.indexOf('target-shared:TSP-57') >= 0, true);

console.log('\n────────────────────────');
console.log(`ผ่าน ${pass} / ไม่ผ่าน ${fail}`);
process.exit(fail ? 1 : 0);
