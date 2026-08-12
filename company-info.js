// ==================== COMPANY INFO (วันเกิดพนักงาน / วันหยุดบริษัท) ====================
// โมดูลสำหรับ "ตอบคำถาม" เท่านั้น
// ⚠️ ไม่ยุ่งกับ flow การลง timesheet เลย — ตัวกันไม่ให้ไปทับ timesheet อยู่ในฝั่ง bot
//    (bot จะเช็ค CLEAR_TIMESHEET_PATTERN ก่อน ถ้าเข้า pattern timesheet จะไม่เรียกไฟล์นี้)

const holidays = require('./holidays.json');
const birthdays = require('./birthdays.json');

// วันหยุดชุดนี้เป็นของ พ.ศ. 2569 = ค.ศ. 2026 (ประกาศ HR Announcement 2025-002)
const HOLIDAY_YEAR = 2026;

const TH_MONTHS = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม',
];
const TH_DAYS = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

function startOfDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function daysBetween(from, to) {
  return Math.round((startOfDay(to) - startOfDay(from)) / 86400000);
}

// "วันพฤหัสบดีที่ 1 มกราคม 2026"
function formatThaiDate(date) {
  return `วัน${TH_DAYS[date.getDay()]}ที่ ${date.getDate()} ${TH_MONTHS[date.getMonth()]} ${date.getFullYear()}`;
}

// ==================== วันหยุด ====================

// วันหยุดบริษัททั้งหมดของ HOLIDAY_YEAR พร้อม Date object เรียงตามวัน
function getHolidayList() {
  return holidays
    .map(h => ({ ...h, date: new Date(HOLIDAY_YEAR, h.month - 1, h.day) }))
    .sort((a, b) => a.date - b.date);
}

// วันหยุดบริษัทถัดไปนับจาก from (null ถ้าหมดปีแล้ว)
function getNextHoliday(from = new Date()) {
  const today = startOfDay(from);
  return getHolidayList().find(h => h.date >= today) || null;
}

// ==================== วันเกิด ====================

// ชื่อที่เก็บเป็น "?" = อ่านจากปฏิทินไม่ออก ยังไม่ได้ยืนยันชื่อ
function displayName(b) {
  return b.name && b.name !== '?' ? b.name : '(ยังไม่ทราบชื่อ)';
}

// วันเกิดของวันที่ระบุ
function getBirthdaysOn(date = new Date()) {
  const d = date.getDate();
  const m = date.getMonth() + 1;
  return birthdays.filter(b => b.day === d && b.month === m);
}

// วันเกิดในเดือนที่ระบุ (1-12) เรียงตามวันที่
function getBirthdaysInMonth(month) {
  return birthdays.filter(b => b.month === month).sort((a, b) => a.day - b.day);
}

// วันเกิดที่ใกล้จะถึง count คนถัดไป (วนข้ามปีให้อัตโนมัติ)
function getUpcomingBirthdays(from = new Date(), count = 5) {
  const today = startOfDay(from);
  return birthdays
    .map(b => {
      let next = new Date(today.getFullYear(), b.month - 1, b.day);
      if (next < today) next = new Date(today.getFullYear() + 1, b.month - 1, b.day);
      return { ...b, date: next, daysLeft: daysBetween(today, next) };
    })
    .sort((a, b) => a.date - b.date)
    .slice(0, count);
}

// ==================== KNOWLEDGE CONTEXT ====================

// รวมข้อมูลทั้งหมด + ข้อมูลที่คำนวณไว้แล้ว ส่งให้ AI ใช้ตอบ
// (คำนวณวันที่/จำนวนวันในโค้ดเอง ไม่ปล่อยให้ AI คิดเลข เพราะ AI คิดวันที่ผิดบ่อย)
function buildKnowledgeContext(now = new Date()) {
  const today = startOfDay(now);
  const lines = [];

  lines.push(`วันนี้: ${formatThaiDate(today)} (พ.ศ. ${today.getFullYear() + 543})`);
  lines.push('');

  // --- วันหยุด ---
  const list = getHolidayList();
  const remaining = list.filter(h => h.date >= today);
  lines.push(`=== วันหยุดบริษัทประจำปี พ.ศ. ${HOLIDAY_YEAR + 543} (ค.ศ. ${HOLIDAY_YEAR}) ทั้งหมด ${list.length} วัน ===`);
  lines.push('(ประกาศจาก HR สำหรับพนักงานบริษัท ศาลาแดง, เพลินจิต, สามย่าน, วาเนสซ่า หวาง — Back Office)');
  list.forEach((h, i) => {
    const diff = daysBetween(today, h.date);
    let when;
    if (diff === 0) when = '← วันนี้';
    else if (diff < 0) when = 'ผ่านไปแล้ว';
    else when = `อีก ${diff} วัน`;
    lines.push(`${i + 1}. ${formatThaiDate(h.date)} — ${h.name} (${when})`);
  });

  const next = getNextHoliday(today);
  lines.push('');
  if (!next) {
    lines.push(`วันหยุดถัดไป: หมดแล้วสำหรับปี ${HOLIDAY_YEAR} (ยังไม่มีประกาศของปีถัดไป)`);
  } else if (daysBetween(today, next.date) === 0) {
    // วันนี้เป็นวันหยุดพอดี → บอกทั้งวันนี้และวันหยุดถัดไปจริง ๆ
    lines.push(`วันนี้เป็นวันหยุดบริษัท: ${next.name}`);
    const after = remaining.find(h => daysBetween(today, h.date) > 0);
    lines.push(after
      ? `วันหยุดถัดไป (หลังจากวันนี้): ${formatThaiDate(after.date)} — ${after.name} (อีก ${daysBetween(today, after.date)} วัน)`
      : `วันหยุดถัดไป: ไม่มีแล้วสำหรับปี ${HOLIDAY_YEAR}`);
  } else {
    lines.push(`วันหยุดถัดไป: ${formatThaiDate(next.date)} — ${next.name} (อีก ${daysBetween(today, next.date)} วัน)`);
  }
  lines.push(`วันหยุดที่เหลือของปี ${HOLIDAY_YEAR}: ${remaining.length} วัน`);
  lines.push('หมายเหตุ: นอกจากนี้บริษัทหยุดเสาร์-อาทิตย์ตามปกติ');
  lines.push('');

  // --- วันเกิด ---
  lines.push(`=== วันเกิดพนักงาน (Birthday Staff) ทั้งหมด ${birthdays.length} คน ===`);
  lines.push('(วันเกิดวนซ้ำทุกปี ข้อมูลนี้ไม่ผูกกับปี ค.ศ. ใด ๆ)');
  for (let m = 1; m <= 12; m++) {
    const inMonth = getBirthdaysInMonth(m);
    if (inMonth.length === 0) continue;
    lines.push(`${TH_MONTHS[m - 1]}: ${inMonth.map(b => `${b.day} = ${displayName(b)}`).join(', ')}`);
  }
  lines.push('');

  const todayBdays = getBirthdaysOn(today);
  lines.push(todayBdays.length > 0
    ? `วันเกิดวันนี้: ${todayBdays.map(displayName).join(', ')}`
    : 'วันเกิดวันนี้: ไม่มี');

  const upcoming = getUpcomingBirthdays(today, 5);
  lines.push('วันเกิดที่ใกล้จะถึง:');
  upcoming.forEach(b => {
    lines.push(`- ${displayName(b)}: ${b.day} ${TH_MONTHS[b.month - 1]} (${b.daysLeft === 0 ? 'วันนี้' : `อีก ${b.daysLeft} วัน`})`);
  });

  // เตือนเรื่องชื่อที่ยังไม่ยืนยันเฉพาะเมื่อมีจริง (ปกติไม่มี)
  if (birthdays.some(b => !b.name || b.name === '?')) {
    lines.push('');
    lines.push('⚠️ รายการที่เขียนว่า "(ยังไม่ทราบชื่อ)" คือมีคนเกิดวันนั้นจริง แต่ยังไม่ได้ยืนยันชื่อ — ให้ตอบตามนั้น ห้ามเดาชื่อเอง');
  }

  return lines.join('\n');
}

// ==================== ตรวจจับคำถาม ====================

// หัวข้อที่โมดูลนี้ตอบได้ — ต้องเจาะจงพอที่จะไม่ไปโดนข้อความ timesheet
const INFO_TOPIC = /วันเกิด|เกิดวันไหน|เกิดเดือนไหน|birth\s*day|bday|วันหยุด|หยุดยาว|นักขัตฤกษ์|วันนักขัตฤกษ์|holiday|day\s*off|ปฏิทิน(วันเกิด|วันหยุด|บริษัท)|หยุด\s*(ไหม|มั้ย|รึเปล่า|หรือเปล่า|วันไหน|กี่วัน|เมื่อไหร่|เมื่อไร)/i;

// ต้องมีลักษณะเป็นคำถาม/คำขอ ไม่ใช่แค่พูดถึงลอย ๆ
const INFO_QUESTION = /\?|ไหน|เมื่อไหร่|เมื่อไร|กี่|ใคร|อะไร|บ้าง|ไหม|มั้ย|รึเปล่า|หรือเปล่า|หน่อย|เช็ค|ขอดู|ถาม|เหลือ|ถัดไป|ต่อไป|ครั้งหน้า|เดือนนี้|เดือนหน้า|ปีนี้|วันนี้|พรุ่งนี้/i;

/**
 * เป็นคำถามเรื่องวันเกิด/วันหยุดหรือไม่
 * ⚠️ ฟังก์ชันนี้ไม่ได้เช็ค timesheet — ฝั่งเรียกใช้ต้องกรอง timesheet ออกก่อนเสมอ
 */
function isInfoQuestion(text, { botMentioned = false } = {}) {
  if (!text || !INFO_TOPIC.test(text)) return false;
  if (botMentioned) return true;            // tag bot มาถามตรง ๆ = ถามแน่นอน
  return INFO_QUESTION.test(text);
}

// ==================== ตอบคำถาม ====================

// คำตอบสำรองกรณี AI ล่ม — ตอบจากข้อมูลดิบตรง ๆ
function fallbackAnswer(now = new Date()) {
  const today = startOfDay(now);
  const next = getNextHoliday(today);
  const upcoming = getUpcomingBirthdays(today, 3);
  const parts = [];
  if (!next) {
    parts.push(`🏖️ วันหยุดปี ${HOLIDAY_YEAR} หมดแล้ว`);
  } else {
    const left = daysBetween(today, next.date);
    parts.push(`🏖️ วันหยุด${left === 0 ? 'วันนี้' : 'ถัดไป'}: ${formatThaiDate(next.date)} — ${next.name}${left === 0 ? '' : ` (อีก ${left} วัน)`}`);
  }
  parts.push(`🎂 วันเกิดใกล้ ๆ นี้: ${upcoming.map(b => `${displayName(b)} (${b.day} ${TH_MONTHS[b.month - 1]})`).join(', ')}`);
  return parts.join('\n');
}

/**
 * ตอบคำถามเรื่องวันเกิด/วันหยุด
 * @param {string} text ข้อความคำถาม (ตัด mention ออกแล้ว)
 * @param {(prompt: string) => Promise<string>} geminiText ฟังก์ชันเรียก Gemini จากฝั่ง bot
 */
async function answerInfoQuestion(text, geminiText) {
  const context = buildKnowledgeContext();
  const prompt = `คุณคือ bot ในแชนแนล Discord ของบริษัท มีหน้าที่ตอบคำถามเรื่อง "วันเกิดพนักงาน" กับ "วันหยุดบริษัท"

ข้อมูลที่มี (ถูกต้อง ใช้ตอบได้เลย คำนวณวันมาให้แล้ว):
${context}

คำถาม: "${text}"

กฎการตอบ:
- ตอบจากข้อมูลข้างบนเท่านั้น ห้ามแต่งวันหรือชื่อขึ้นมาเอง
- ถ้าข้อมูลไม่มี ให้บอกตรง ๆ ว่าไม่มีในปฏิทิน อย่าเดา
- ห้ามคำนวณวันเอง ใช้ตัวเลข "อีกกี่วัน" ที่ให้มาแล้วเท่านั้น
- ห้ามใช้คำว่า "วันนี้" กับรายการที่ไม่ใช่ของวันนี้จริง (กันคนอ่านเข้าใจผิด)
- ตอบสั้น กระชับ ภาษาไทย เป็นกันเอง กวน ๆ ได้ (แชนแนลนี้คุยกันสบาย ๆ)
- ถ้าต้องลิสต์หลายรายการ ใช้ bullet สั้น ๆ ไม่เกิน 15 บรรทัด
- ใส่ emoji ได้ 1-2 ตัว (🎂 วันเกิด / 🏖️ วันหยุด)
- ตอบเฉพาะคำตอบ ไม่ต้องเกริ่น ไม่ต้องอธิบายว่าดูจากข้อมูลอะไร`;

  try {
    const answer = await geminiText(prompt);
    return answer && answer.trim() ? answer.trim() : fallbackAnswer();
  } catch (err) {
    console.error('❌ ตอบคำถามวันเกิด/วันหยุดไม่ได้:', err.message);
    return fallbackAnswer();
  }
}

module.exports = {
  HOLIDAY_YEAR,
  isInfoQuestion,
  answerInfoQuestion,
  buildKnowledgeContext,
  getHolidayList,
  getNextHoliday,
  getBirthdaysOn,
  getBirthdaysInMonth,
  getUpcomingBirthdays,
  fallbackAnswer,
};
