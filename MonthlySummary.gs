/**
 * ระบบงานรวมรายเดือน: สรุปรายบุคคล + รายโครงการ
 * อ่านจากชีต "แผนรายเดือน" (มอบหมายต้นเดือน) เทียบกับ "ส่งงานรายวัน" และ "แจ้งเตือน"
 * ใช้ร่วมกับ Code.gs (CFG, normMonth, parseDate, thaiDate ฯลฯ)
 */

const SHEET_MONTHLY = 'สรุปรายเดือน';

/** คำนวณสรุปของเดือน month (รูปแบบ yyyy-MM ค.ศ. เช่น 2026-10) */
function getMonthlySummary(month) {
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName(CFG.SHEET_PLAN).getDataRange().getValues().slice(1)
    .filter(r => normMonth(r[0]) === month);
  const daily = ss.getSheetByName(CFG.SHEET_DAILY).getDataRange().getValues().slice(1)
    .filter(r => monthKey(parseDate(r[0])) === month);
  const alerts = ss.getSheetByName(CFG.SHEET_ALERT).getDataRange().getValues().slice(1)
    .filter(r => monthKey(parseDate(r[0])) === month);

  const persons = {}, projects = {};
  const newRow = () => ({ planned: 0, done: 0, late: 0, pending: 0, overdue: 0 });

  plan.forEach(p => {
    const person = p[1], project = p[2], ea = String(p[3]);
    const rows = daily.filter(d => d[1] === person && String(d[3]) === ea);
    const done = rows.length > 0 ? 1 : 0;                  // EA นี้มีการส่งงานอย่างน้อย 1 ครั้ง
    const late = rows.filter(d => d[6] === 'ส่งช้า').length;
    const pending = rows.filter(d => d[6] === 'นอกแผน-รออนุมัติ').length;
    const overdue = alerts.filter(a => a[1] === person && String(a[2]) === ea && a[4] === 'ค้างอยู่').length;

    [[persons, person], [projects, project]].forEach(([bag, key]) => {
      const o = bag[key] || (bag[key] = newRow());
      o.planned += 1; o.done += done; o.late += late; o.pending += pending; o.overdue += overdue;
    });
  });

  const finish = bag => Object.keys(bag).map(k => {
    const o = bag[k];
    return Object.assign({ name: k, remaining: o.planned - o.done,
      pct: o.planned ? Math.round(o.done * 100 / o.planned) : 0 }, o);
  }).sort((a, b) => a.name.localeCompare(b.name, 'th'));

  const people = finish(persons), projs = finish(projects);
  const total = people.reduce((t, o) => ({
    planned: t.planned + o.planned, done: t.done + o.done,
    late: t.late + o.late, overdue: t.overdue + o.overdue }), { planned: 0, done: 0, late: 0, overdue: 0 });
  total.remaining = total.planned - total.done;
  total.pct = total.planned ? Math.round(total.done * 100 / total.planned) : 0;

  return { month, label: thaiMonthLabel(month), total, persons: people, projects: projs };
}

/** "2026-10" -> "ตุลาคม 2569" */
function thaiMonthLabel(month) {
  const [y, m] = month.split('-').map(Number);
  return THAI_MONTHS[m - 1] + ' ' + (y + 543);
}

/** บันทึกภาพรวมของเดือนลงชีต "สรุปรายเดือน" (รันได้เองหรือผ่าน Trigger วันที่ 1) */
function snapshotMonth(month) {
  month = month || monthKey(new Date());
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_MONTHLY) || ss.insertSheet(SHEET_MONTHLY);
  if (sh.getLastRow() === 0)
    sh.appendRow(['เดือน','ประเภท','ชื่อ','EA ตามแผน','ส่งแล้ว','คงเหลือ','ร้อยละ','ส่งช้า (ครั้ง)','ค้างส่ง (วัน)']);
  const s = getMonthlySummary(month);
  const rows = [];
  s.persons.forEach(o => rows.push([s.label, 'รายบุคคล', o.name, o.planned, o.done, o.remaining, o.pct, o.late, o.overdue]));
  s.projects.forEach(o => rows.push([s.label, 'รายโครงการ', o.name, o.planned, o.done, o.remaining, o.pct, o.late, o.overdue]));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/** เก็บสรุปของเดือนที่แล้วอัตโนมัติ ทุกวันที่ 1 ของเดือน (รันครั้งเดียวเพื่อตั้ง Trigger) */
function createMonthlyTrigger() {
  ScriptApp.newTrigger('snapshotLastMonth').timeBased()
    .onMonthDay(1).atHour(6).inTimezone(CFG.TZ).create();
}

function snapshotLastMonth() {
  const d = new Date(); d.setMonth(d.getMonth() - 1);
  snapshotMonth(monthKey(d));
}
