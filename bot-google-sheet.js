require('dotenv').config();
process.env.TZ = 'Asia/Bangkok';

const { Client, GatewayIntentBits, Partials } = require('discord.js');
const { google } = require('googleapis');
const OpenAI = require('openai');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ==================== CONFIG ====================

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent
  ],
  partials: [Partials.Message]
});

// Google Sheets Config
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_NAME = 'ลงเวลางาน';
const MSG_ID_COL = 'Z'; // คอลัมน์ซ่อน message ID (ไกลมากไม่มีใครเห็น)

// Discord Config
const DISCORD_TOKEN = process.env.DISCORD_TOKEN_SHEET;
//const CHANNEL_ID = '1481565840477388841';
//const CHANNEL_ID = '1483304085783318598';
const CHANNEL_ID = '1488904036999499910';
//const CHANNEL_ID = 'test';

// Reminder Config
const USERS_SHEET = 'Users';       // ชื่อ Sheet ที่เก็บรายชื่อ
const REMINDER_HOUR = 18;          // แจ้งเตือนตอน 15:00
const REMINDER_MINUTE = 30;

// ==================== GOOGLE SHEETS AUTH ====================

// รองรับทั้ง keyFile (local) และ env variable (cloud)
let auth;
if (process.env.GOOGLE_CREDENTIALS) {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
} else {
  auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, 'google-credentials.json'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

const sheets = google.sheets({ version: 'v4', auth });

// ==================== GOOGLE SHEETS FUNCTIONS ====================

// อ่าน headers (แถวแรก) ของ Sheet
async function getHeaders() {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!1:1`,
    });
    return res.data.values?.[0] || [];
  } catch (err) {
    console.error('❌ อ่าน headers ไม่ได้:', err.message);
    return [];
  }
}

// อ่านข้อมูลจาก Sheet
async function readSheet(range = `${SHEET_NAME}!A:Z`) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
    });
    return res.data.values || [];
  } catch (err) {
    console.error('❌ อ่าน Sheet ไม่ได้:', err.message);
    return [];
  }
}

// เขียนข้อมูลต่อท้าย Sheet (append) - return row number ที่เขียน
async function appendToSheet(headers, values) {
  try {
    const lastCol = String.fromCharCode(64 + headers.length); // A=1, B=2, ... D=4
    const res = await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:${lastCol}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });
    const updatedRange = res.data.updates.updatedRange;
    const rowNumber = parseInt(updatedRange.match(/\d+$/)[0]);
    console.log(`✅ เขียนลง Sheet แถวที่ ${rowNumber}`);
    return rowNumber;
  } catch (err) {
    console.error('❌ เขียน Sheet ไม่ได้:', err.message);
    return null;
  }
}

// อัพเดทข้อมูลใน cell/range ที่ระบุ
async function updateSheet(range, values) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!${range}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });
    console.log('✅ อัพเดท Sheet แล้ว');
    return true;
  } catch (err) {
    console.error('❌ อัพเดท Sheet ไม่ได้:', err.message);
    return false;
  }
}

// เขียน message ID ลงคอลัมน์ซ่อน
async function saveMessageId(rowNumber, messageId) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!${MSG_ID_COL}${rowNumber}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[messageId]] },
    });
  } catch (err) {
    console.error('❌ เก็บ message ID ไม่ได้:', err.message);
  }
}

// หาทุกแถวที่มี message ID นี้ (สำหรับ multi-line)
async function findAllRowsByMessageId(messageId) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!${MSG_ID_COL}:${MSG_ID_COL}`,
    });
    const data = res.data.values || [];
    const rows = [];
    data.forEach((row, idx) => {
      if (row[0] === messageId) rows.push(idx + 1);
    });
    return rows;
  } catch (err) {
    console.error('❌ หา message IDs ไม่ได้:', err.message);
    return [];
  }
}

// ลบหลายแถว (ต้องลบจากล่างขึ้นบนเพื่อไม่ให้ index เลื่อน)
async function deleteRows(rowNumbers) {
  try {
    const sheetMeta = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      fields: 'sheets.properties',
    });
    const sheet = sheetMeta.data.sheets.find(s => s.properties.title === SHEET_NAME);
    const sheetId = sheet.properties.sheetId;

    // ลบจากล่างขึ้นบน
    const sorted = [...rowNumbers].sort((a, b) => b - a);
    const requests = sorted.map(row => ({
      deleteDimension: {
        range: {
          sheetId,
          dimension: 'ROWS',
          startIndex: row - 1,
          endIndex: row,
        },
      },
    }));

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests },
    });

    console.log(`🗑️ ลบ ${rowNumbers.length} แถวจาก Sheet แล้ว`);
  } catch (err) {
    console.error('❌ ลบแถวไม่ได้:', err.message);
  }
}

// ==================== DROPDOWN FUNCTIONS ====================

// อ่านค่า dropdown (data validation) จาก Sheet (อ่านจากแถวข้อมูลแถว 2)
async function getDropdownValues() {
  try {
    const res = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      ranges: [`${SHEET_NAME}!2:2`],
      fields: 'sheets.data.rowData.values.dataValidation',
    });

    const dropdowns = {};
    const rowData = res.data.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values || [];

    const headers = await getHeaders();

    rowData.forEach((cell, idx) => {
      if (cell.dataValidation?.condition?.type === 'ONE_OF_LIST') {
        const values = cell.dataValidation.condition.values.map(v => v.userEnteredValue);
        if (headers[idx]) dropdowns[headers[idx]] = values;
      }
    });

    return dropdowns;
  } catch (err) {
    console.error('❌ อ่าน dropdown ไม่ได้:', err.message);
    return {};
  }
}

// ==================== AI FUNCTION ====================

// ให้ AI วิเคราะห์ข้อความแล้ว map ลง column ที่เหมาะสม (รองรับ multi-line)
async function aiMapToColumns(headers, username, messageText, dropdowns = {}) {
  // สร้าง dropdown info สำหรับ prompt
  let dropdownInfo = '';
  if (Object.keys(dropdowns).length > 0) {
    dropdownInfo = '\n\nคอลัมน์ที่มี dropdown (ต้องเลือกจากค่าเหล่านี้เท่านั้น):';
    for (const [col, values] of Object.entries(dropdowns)) {
      dropdownInfo += `\n- "${col}": [${values.map(v => `"${v}"`).join(', ')}]`;
    }
    dropdownInfo += '\n⚠️ สำหรับคอลัมน์ dropdown: ถ้าผู้ใช้พิมพ์ชื่อผิด/พิมพ์ภาษาไทย/พิมพ์ย่อ ให้ fuzzy match หาค่าที่ใกล้เคียงที่สุดจาก dropdown เช่น "เอนค่อนฟัน" → "Enconfund", "enconfun" → "Enconfund"';
  }

  const prompt = `คุณคือระบบจัดเรียงข้อมูลลง Google Sheet

Sheet มี columns ดังนี้: ${JSON.stringify(headers)}${dropdownInfo}

ข้อมูลที่ได้รับ:
- ผู้ส่ง: ${username}
- ข้อความ:
${messageText}
- วันที่วันนี้: ${new Date().toLocaleDateString('th-TH', { timeZone: 'Asia/Bangkok', day: '2-digit', month: '2-digit', year: 'numeric' })}
- วันที่เมื่อวาน (⚠️ ใช้เฉพาะเมื่อผู้ใช้พูดถึง "เมื่อวาน"/"วานนี้" อย่างชัดเจนเท่านั้น): ${new Date(Date.now() - 86400000).toLocaleDateString('th-TH', { timeZone: 'Asia/Bangkok', day: '2-digit', month: '2-digit', year: 'numeric' })}
- เวลาปัจจุบัน: ${new Date().toLocaleTimeString('th-TH', { timeZone: 'Asia/Bangkok' })}

⚠️ สำคัญมาก: สร้าง row ตามจำนวนบรรทัดข้อมูลเท่านั้น ห้ามสร้าง row เพิ่มเอง ถ้าผู้ใช้พิมพ์มา 1 บรรทัด = 1 row, 3 บรรทัด = 3 rows ถ้าไม่ได้พูดถึง "เมื่อวาน" ให้ใช้วันที่วันนี้ทั้งหมด

กฎ:
1. ถ้าข้อความมีหลายบรรทัด ให้แยกเป็นหลาย row (แต่ละบรรทัด = 1 row)
2. ตอบเป็น JSON object ที่มี key "rows" เป็น array ของ objects
3. แต่ละ object มี key คือชื่อ column (ตรงกับ headers ทุกประการ)
4. ถ้า column ไหนไม่เกี่ยวข้อง ให้ใส่ค่าว่าง ""
5. สำหรับคอลัมน์ที่มี dropdown ต้องใช้ค่าจาก dropdown เท่านั้น (fuzzy match ถ้าพิมพ์ผิด)
6. ตอบเป็น JSON เท่านั้น ไม่ต้องอธิบาย
7. ⚠️ สำคัญมาก: แยกข้อมูลออกจากกันให้ชัดเจน เช่น "ประชุม cfarm 3 ชม." → Detail="ประชุม", Hour="3 ชม.", Project="Cfarm" ห้ามเอาชั่วโมงหรือชื่อ project ไปรวมใน Detail
8. ⚠️ สำคัญ: ถ้าบรรทัดไหนไม่ได้พูดถึง project → ใส่ "" เท่านั้น ห้ามเดาหรือใส่ project เอง เช่น "เดินทาง 2 ชม." ไม่ได้พูดถึง project → Project=""
9. ⚠️ ถ้ามีบรรทัด "Project Cfarm" แยกไว้ท้ายสุด → ใช้ project นั้นกับทุกบรรทัดข้างบนที่ยังไม่มี project กำกับ ห้ามสร้าง row ใหม่สำหรับบรรทัด project (ใช้เฉพาะเมื่อมีบรรทัด Project แยกชัดเจน)
10. ⚠️ Hour: ใส่เฉพาะเมื่อระบุชั่วโมง/นาทีชัดเจน เช่น "3 ชม.", "30 นาที" ถ้าเป็นหน่วยอื่น เช่น "5 รอบ" ไม่ใช่ชั่วโมง → ให้เอาไว้ใน Detail แทน
10. ถ้าบรรทัดมีเวลานำหน้า เช่น "09:00 Assign งาน" ให้แยกเวลาใส่คอลัมน์ Time และเอาส่วนที่เหลือใส่ Detail
11. ⚠️ วันที่: ถ้าไม่ได้ระบุวันที่ → ใช้วันที่วันนี้, ถ้าระบุวันที่มาด้วย → ใช้วันที่ที่ระบุ
    - รองรับหลายรูปแบบ: "เมื่อวาน", "วานนี้", "yesterday" → ใช้วันที่เมื่อวาน
    - "อันนี้ของเมื่อวาน", "ของวันที่ 10/03/2026", "วันที่ 10 มีนา" → ใช้วันที่ที่ระบุ
    - "เมื่อวาน" หรือ "ของเมื่อวาน" ท้ายข้อความ → ใช้กับทุก row (เหมือน Project ท้ายสุด)
    - ถ้าระบุวันที่เป็น พ.ศ. ให้แปลงเป็น ค.ศ. ก่อน (เช่น 2569 → 2026)
    - format วันที่เป็น dd/mm/yyyy เสมอ (ค.ศ.)

ตัวอย่าง input 1 (วันนี้):
"09:00 Assign งาน Abacus ให้ฝั่ง Dev
10:00 ทำคู่มือ WA
Project Cfarm"

ตัวอย่าง output 1:
{"rows": [
  {"dd/mm/yyyy": "15/03/2026", "Name": "user1", "Time": "09:00", "Detail": "Assign งาน Abacus ให้ฝั่ง Dev", "Hour": "", "Project": "Cfarm"},
  {"dd/mm/yyyy": "15/03/2026", "Name": "user1", "Time": "10:00", "Detail": "ทำคู่มือ WA", "Hour": "", "Project": "Cfarm"}
]}

ตัวอย่าง input 2 (ย้อนหลัง):
"09:00 Assign งาน
10:00 ประชุม
Project zg อันนี้ของเมื่อวาน"

ตัวอย่าง output 2:
{"rows": [
  {"dd/mm/yyyy": "14/03/2026", "Name": "user1", "Time": "09:00", "Detail": "Assign งาน", "Hour": "", "Project": "ZG"},
  {"dd/mm/yyyy": "14/03/2026", "Name": "user1", "Time": "10:00", "Detail": "ประชุม", "Hour": "", "Project": "ZG"}
]}

ตัวอย่าง input 3 (ระบุวันที่ + Project ท้ายสุด):
"ประชุม 3 ชม.
เดินทาง 2 ชม.
Project cfarm วันที่ 10/03/2026"

ตัวอย่าง output 3:
{"rows": [
  {"dd/mm/yyyy": "10/03/2026", "Name": "user1", "Detail": "ประชุม", "Hour": "3 ชม.", "Project": "Cfarm"},
  {"dd/mm/yyyy": "10/03/2026", "Name": "user1", "Detail": "เดินทาง", "Hour": "2 ชม.", "Project": "Cfarm"}
]}

ตัวอย่าง input 4 (บางบรรทัดมี project บางบรรทัดไม่มี):
"ประชุม cfarm 3 ชม.
เดินทาง 2 ชม.
เทรนนิ่ง 3 ชม.
ทำscript zg 5 ชม.
แอบหัวหน้านอน 30 นาที
เดินเข้าห้องน้ำ 5 รอบ"

ตัวอย่าง output 4:
{"rows": [
  {"dd/mm/yyyy": "01/04/2026", "Name": "user1", "Detail": "ประชุม", "Hour": "3 ชม.", "Project": "Cfarm"},
  {"dd/mm/yyyy": "01/04/2026", "Name": "user1", "Detail": "เดินทาง", "Hour": "2 ชม.", "Project": ""},
  {"dd/mm/yyyy": "01/04/2026", "Name": "user1", "Detail": "เทรนนิ่ง", "Hour": "3 ชม.", "Project": ""},
  {"dd/mm/yyyy": "01/04/2026", "Name": "user1", "Detail": "ทำscript", "Hour": "5 ชม.", "Project": "ZG"},
  {"dd/mm/yyyy": "01/04/2026", "Name": "user1", "Detail": "แอบหัวหน้านอน", "Hour": "30 นาที", "Project": ""},
  {"dd/mm/yyyy": "01/04/2026", "Name": "user1", "Detail": "เดินเข้าห้องน้ำ 5 รอบ", "Hour": "", "Project": ""}
]}`;

  try {
    const res = await openai.chat.completions.create({
      model: 'gpt-5.4-mini-2026-03-17',
      messages: [{ role: 'user', content: prompt }],
      response_format: { type: 'json_object' },
    });

    const result = JSON.parse(res.choices[0].message.content);
    return result.rows || [result]; // fallback ถ้า AI ตอบเป็น object เดียว
  } catch (err) {
    console.error('❌ AI วิเคราะห์ไม่ได้:', err.message);
    return null;
  }
}

// แปลง AI result เป็น row array ตามลำดับ headers
function mapToRow(headers, aiResult) {
  return headers.map(header => aiResult[header] || '');
}

// ==================== EXPORT FUNCTION ====================

// ให้ AI วิเคราะห์ว่าเป็นคำสั่ง export หรือไม่ และต้องการกรองอะไร
async function aiDetectIntent(messageText, mentionedUsers) {
  const prompt = `วิเคราะห์ข้อความนี้ว่าเป็น "export" (ขอสรุป/ดึงข้อมูล/export) หรือ "data" (กรอกข้อมูลปกติ)

ข้อความ: "${messageText}"
ผู้ที่ถูก mention: ${mentionedUsers.length > 0 ? mentionedUsers.map(u => u.username).join(', ') : 'ไม่มี'}

วันที่วันนี้: ${new Date().toLocaleDateString('th-TH', { timeZone: 'Asia/Bangkok', day: '2-digit', month: '2-digit', year: 'numeric' })}

ตอบเป็น JSON:
- ถ้าเป็น export: {"intent": "export", "split_per_person": true/false, "filters": {"name": "ชื่อที่ต้องการกรอง หรือ null", "date_from": "dd/mm/yyyy หรือ null", "date_to": "dd/mm/yyyy หรือ null", "specific_dates": ["dd/mm/yyyy", ...] หรือ null, "project": "ชื่อ project หรือ null"}}
- ถ้าเป็นกรอกข้อมูล: {"intent": "data"}

หมายเหตุ:
- คำเช่น "สรุป", "export", "ดึงข้อมูล", "timesheet" = export
- split_per_person = true เมื่อผู้ใช้ต้องการแยกไฟล์คนละ report เช่น "ทำเป็น 2 report", "แยกคนละไฟล์", "แยกรายคน"
- split_per_person = false เมื่อต้องการรวมในไฟล์เดียว เช่น "สรุปรวม", "รวมเป็นไฟล์เดียว" หรือไม่ได้ระบุ
- specific_dates: ใช้เมื่อผู้ใช้ระบุวันที่เฉพาะเจาะจง เช่น "เอาแค่วันที่ 17 และ 31 ของเดือนมีนา" → ["17/03/2026", "31/03/2026"]
- ถ้าผู้ใช้พูดถึง "เดือนมีนา" / "มีนาคม" / "มี.ค." → เดือน 03, "เมษา" → 04, ฯลฯ ใช้ปี ค.ศ. เสมอ
- ใช้ date_from/date_to สำหรับช่วงวันที่ ใช้ specific_dates สำหรับวันที่เฉพาะเจาะจง`;

  try {
    const res = await openai.chat.completions.create({
      model: 'gpt-5.4-mini-2026-03-17',
      messages: [{ role: 'user', content: prompt }],
      response_format: { type: 'json_object' },
    });
    return JSON.parse(res.choices[0].message.content);
  } catch (err) {
    console.error('❌ AI detect intent ไม่ได้:', err.message);
    return { intent: 'data' };
  }
}

// Export ข้อมูลจาก Sheet เป็นไฟล์ xlsx
async function exportSheet(filters, message) {
  const headers = await getHeaders();
  const data = await readSheet();
  if (data.length <= 1) {
    await message.reply('📭 ไม่มีข้อมูลใน Sheet');
    return;
  }

  const dateColIdx = headers.findIndex(h => h.toLowerCase().includes('dd') || h.toLowerCase().includes('date') || h.includes('/'));
  const nameColIdx = headers.findIndex(h => h.toLowerCase() === 'name' || h.toLowerCase().includes('name'));
  const projectColIdx = headers.findIndex(h => h.toLowerCase() === 'project' || h.toLowerCase().includes('project'));

  // กรองข้อมูล (ข้ามแถว header)
  let filtered = data.slice(1);

  // กรองตามชื่อ (รองรับหลายคน)
  const names = filters.names || (filters.name ? [filters.name] : []);
  if (names.length > 0) {
    filtered = filtered.filter(row => {
      const name = (row[nameColIdx] || '').toLowerCase();
      return names.some(n => name.includes(n.toLowerCase()));
    });
  }

  // กรองตาม project
  if (filters.project) {
    filtered = filtered.filter(row => {
      const project = (row[projectColIdx] || '').toLowerCase();
      return project.includes(filters.project.toLowerCase());
    });
  }

  // กรองตามวันที่เฉพาะ (specific_dates)
  if (filters.specific_dates && filters.specific_dates.length > 0) {
    // normalize ทุกวันที่ที่ต้องการ
    const targetDates = filters.specific_dates.map(ds => {
      const p = ds.split('/');
      if (p.length < 3) return ds;
      return `${parseInt(p[0])}/${parseInt(p[1])}/${parseInt(p[2]) > 2500 ? parseInt(p[2]) - 543 : parseInt(p[2])}`;
    });

    filtered = filtered.filter(row => {
      const dateStr = (row[dateColIdx] || '').trim();
      const parts = dateStr.split('/');
      if (parts.length < 3) return false;
      const normalized = `${parseInt(parts[0])}/${parseInt(parts[1])}/${parseInt(parts[2]) > 2500 ? parseInt(parts[2]) - 543 : parseInt(parts[2])}`;
      return targetDates.includes(normalized);
    });
  }

  // กรองตามช่วงวันที่ (date_from / date_to)
  if (filters.date_from || filters.date_to) {
    filtered = filtered.filter(row => {
      const dateStr = (row[dateColIdx] || '').trim();
      const parts = dateStr.split('/');
      if (parts.length < 3) return false;
      const d = new Date(parseInt(parts[2]) > 2500 ? parseInt(parts[2]) - 543 : parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));

      if (filters.date_from) {
        const fp = filters.date_from.split('/');
        const from = new Date(parseInt(fp[2]) > 2500 ? parseInt(fp[2]) - 543 : parseInt(fp[2]), parseInt(fp[1]) - 1, parseInt(fp[0]));
        if (d < from) return false;
      }
      if (filters.date_to) {
        const tp = filters.date_to.split('/');
        const to = new Date(parseInt(tp[2]) > 2500 ? parseInt(tp[2]) - 543 : parseInt(tp[2]), parseInt(tp[1]) - 1, parseInt(tp[0]));
        if (d > to) return false;
      }
      return true;
    });
  }

  if (filtered.length === 0) {
    await message.reply('📭 ไม่พบข้อมูลตามเงื่อนไขที่ระบุ');
    return;
  }

  // สร้างไฟล์ xlsx (ใช้แค่ header columns ไม่รวมคอลัมน์ซ่อน)
  const exportData = [headers, ...filtered.map(row => headers.map((_, i) => row[i] || ''))];
  const ws = XLSX.utils.aoa_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Timesheet');

  const fileName = `timesheet_${Date.now()}.xlsx`;
  const filePath = path.join(__dirname, fileName);
  XLSX.writeFile(wb, filePath);

  // ส่งไฟล์กลับ Discord
  await message.reply({
    content: `📊 สรุปข้อมูล ${filtered.length} รายการ`,
    files: [{ attachment: filePath, name: fileName }],
  });

  // ลบไฟล์ temp
  fs.unlinkSync(filePath);
  console.log(`📊 Export ${filtered.length} รายการ`);
}

// Export แยกคนละไฟล์
async function exportSheetPerPerson(filters, message) {
  const headers = await getHeaders();
  const data = await readSheet();
  if (data.length <= 1) {
    await message.reply('📭 ไม่มีข้อมูลใน Sheet');
    return;
  }

  const dateColIdx = headers.findIndex(h => h.toLowerCase().includes('dd') || h.toLowerCase().includes('date') || h.includes('/'));
  const nameColIdx = headers.findIndex(h => h.toLowerCase() === 'name' || h.toLowerCase().includes('name'));
  const projectColIdx = headers.findIndex(h => h.toLowerCase() === 'project' || h.toLowerCase().includes('project'));

  const names = filters.names || (filters.name ? [filters.name] : []);
  if (names.length === 0) {
    await exportSheet(filters, message);
    return;
  }

  // กรองพื้นฐาน (วันที่, project)
  let baseFiltered = data.slice(1);
  if (filters.project && projectColIdx !== -1) {
    baseFiltered = baseFiltered.filter(row => {
      const project = (row[projectColIdx] || '').toLowerCase();
      return project.includes(filters.project.toLowerCase());
    });
  }
  // กรองตามวันที่เฉพาะ (specific_dates)
  if (filters.specific_dates && filters.specific_dates.length > 0) {
    const targetDates = filters.specific_dates.map(ds => {
      const p = ds.split('/');
      if (p.length < 3) return ds;
      return `${parseInt(p[0])}/${parseInt(p[1])}/${parseInt(p[2]) > 2500 ? parseInt(p[2]) - 543 : parseInt(p[2])}`;
    });
    baseFiltered = baseFiltered.filter(row => {
      const dateStr = (row[dateColIdx] || '').trim();
      const parts = dateStr.split('/');
      if (parts.length < 3) return false;
      const normalized = `${parseInt(parts[0])}/${parseInt(parts[1])}/${parseInt(parts[2]) > 2500 ? parseInt(parts[2]) - 543 : parseInt(parts[2])}`;
      return targetDates.includes(normalized);
    });
  }
  if (filters.date_from || filters.date_to) {
    baseFiltered = baseFiltered.filter(row => {
      const dateStr = (row[dateColIdx] || '').trim();
      const parts = dateStr.split('/');
      if (parts.length < 3) return false;
      const d = new Date(parseInt(parts[2]) > 2500 ? parseInt(parts[2]) - 543 : parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      if (filters.date_from) {
        const fp = filters.date_from.split('/');
        const from = new Date(parseInt(fp[2]) > 2500 ? parseInt(fp[2]) - 543 : parseInt(fp[2]), parseInt(fp[1]) - 1, parseInt(fp[0]));
        if (d < from) return false;
      }
      if (filters.date_to) {
        const tp = filters.date_to.split('/');
        const to = new Date(parseInt(tp[2]) > 2500 ? parseInt(tp[2]) - 543 : parseInt(tp[2]), parseInt(tp[1]) - 1, parseInt(tp[0]));
        if (d > to) return false;
      }
      return true;
    });
  }

  // สร้างไฟล์แยกแต่ละคน
  const files = [];
  const tempPaths = [];

  for (const name of names) {
    const personData = baseFiltered.filter(row => {
      const rowName = (row[nameColIdx] || '').toLowerCase();
      return rowName.includes(name.toLowerCase());
    });

    if (personData.length === 0) continue;

    const exportData = [headers, ...personData.map(row => headers.map((_, i) => row[i] || ''))];
    const ws = XLSX.utils.aoa_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Timesheet');

    const fileName = `timesheet_${name}_${Date.now()}.xlsx`;
    const filePath = path.join(__dirname, fileName);
    XLSX.writeFile(wb, filePath);

    files.push({ attachment: filePath, name: fileName });
    tempPaths.push(filePath);
  }

  if (files.length === 0) {
    await message.reply('📭 ไม่พบข้อมูลตามเงื่อนไขที่ระบุ');
    return;
  }

  await message.reply({
    content: `📊 สรุปข้อมูลแยกรายคน ${files.length} ไฟล์`,
    files,
  });

  // ลบไฟล์ temp
  tempPaths.forEach(p => fs.unlinkSync(p));
  console.log(`📊 Export แยก ${files.length} ไฟล์`);
}

// ==================== DISCORD BOT ====================

client.once('ready', async () => {
  console.log(`✅ Bot พร้อมใช้งาน: ${client.user.tag}`);
  console.log(`📊 เชื่อมต่อ Google Sheet ID: ${SPREADSHEET_ID}`);

  const headers = await getHeaders();
  console.log(`📋 Columns ใน Sheet: ${headers.join(', ')}`);
});

client.on('messageCreate', async (message) => {
  if (message.author.bot) return;
  if (message.channel.id !== CHANNEL_ID) return;

  // ให้ AI ตรวจจับว่าเป็นคำสั่ง export หรือกรอกข้อมูล
  // ตัด mention ออกจากข้อความเพื่อดู content จริง
  const cleanContent = message.content.replace(/<@!?\d+>/g, '').trim();
  const mentionedUsers = [...message.mentions.users.values()];

  // ถ้าไม่มีคำสั่ง export ชัดเจน (สรุป, export, ดึงข้อมูล, timesheet) → ไม่ต้องถาม AI ให้ลงข้อมูลเลย
  const exportKeywords = /สรุป|export|ดึงข้อมูล|timesheet|รายงาน|report/i;
  let intent = { intent: 'data' };
  if (exportKeywords.test(cleanContent)) {
    intent = await aiDetectIntent(message.content, mentionedUsers);
  }
  console.log('🧠 Intent:', JSON.stringify(intent));

  // ถ้าเป็น export
  if (intent.intent === 'export') {
    // ถ้า mention user มา ใช้ username ของทุกคนที่ถูก mention
    if (mentionedUsers.length > 0) {
      intent.filters.names = mentionedUsers.map(u => u.username);
    }
    if (intent.split_per_person && (intent.filters.names?.length > 1)) {
      await exportSheetPerPerson(intent.filters, message);
    } else {
      await exportSheet(intent.filters, message);
    }
    return;
  }

  // ถ้าเป็นกรอกข้อมูลปกติ
  const headers = await getHeaders();
  if (headers.length === 0) {
    message.react('❌');
    return;
  }

  const dropdowns = await getDropdownValues();

  // ให้ AI วิเคราะห์ข้อความ (รองรับ multi-line → หลายแถว)
  const aiRows = await aiMapToColumns(headers, message.author.username, message.content, dropdowns);
  if (!aiRows || aiRows.length === 0) {
    message.react('❌');
    return;
  }

  console.log(`🤖 AI วิเคราะห์: ${aiRows.length} แถว`, aiRows);

  // เขียนทีละแถว เพื่อเก็บ message ID ทุกแถว
  let allSuccess = true;
  for (const aiResult of aiRows) {
    const row = mapToRow(headers, aiResult);
    const rowNumber = await appendToSheet(headers, [row]);
    if (rowNumber) {
      await saveMessageId(rowNumber, message.id);
    } else {
      allSuccess = false;
    }
  }

  message.react(allSuccess ? '✅' : '❌');
});

// เมื่อมีการแก้ไขข้อความ → ให้ AI วิเคราะห์ใหม่แล้วอัพเดท
client.on('messageUpdate', async (oldMessage, newMessage) => {
  // ถ้าข้อความเป็น partial (หลัง restart) ให้ fetch ข้อมูลเต็มก่อน
  if (newMessage.partial) await newMessage.fetch();
  if (newMessage.author?.bot) return;
  if (newMessage.channel.id !== CHANNEL_ID) return;
  if (!newMessage.content) return;

  // หาทุกแถวที่มี message ID นี้
  const allRows = await findAllRowsByMessageId(newMessage.id);
  if (allRows.length === 0) {
    console.log('⚠️ ไม่เจอ message ID ใน Sheet:', newMessage.id);
    return;
  }
  console.log(`📝 Edit: เจอ ${allRows.length} แถวเดิม:`, allRows);

  const headers = await getHeaders();
  if (headers.length === 0) return;

  const dropdowns = await getDropdownValues();

  // ให้ AI วิเคราะห์ข้อความใหม่
  const aiRows = await aiMapToColumns(headers, newMessage.author.username, newMessage.content, dropdowns);
  if (!aiRows || aiRows.length === 0) return;

  // ถ้าจำนวนแถวเท่ากัน → อัพเดทในที่เดิม (ไม่ต้องลบ+สร้างใหม่)
  if (aiRows.length === allRows.length) {
    const lastCol = String.fromCharCode(64 + headers.length);
    for (let i = 0; i < aiRows.length; i++) {
      const row = mapToRow(headers, aiRows[i]);
      await updateSheet(`A${allRows[i]}:${lastCol}${allRows[i]}`, [row]);
    }
    console.log(`📝 อัพเดท ${aiRows.length} แถวในที่เดิม`);
  } else {
    // จำนวนแถวเปลี่ยน → ลบเก่าแล้วสร้างใหม่
    await deleteRows(allRows);
    for (const aiResult of aiRows) {
      const row = mapToRow(headers, aiResult);
      const newRow = await appendToSheet(headers, [row]);
      if (newRow) await saveMessageId(newRow, newMessage.id);
    }
    console.log(`📝 ลบ ${allRows.length} แถวเก่า → สร้าง ${aiRows.length} แถวใหม่`);
  }

  newMessage.react('📝');
});

// เมื่อมีการลบข้อความ → ลบทุกแถวที่เกี่ยวข้องใน Sheet
client.on('messageDelete', async (message) => {
  if (message.channel.id !== CHANNEL_ID) return;

  const rows = await findAllRowsByMessageId(message.id);
  if (rows.length === 0) {
    console.log('⚠️ ลบ: ไม่เจอ message ID ใน Sheet:', message.id);
    return;
  }

  await deleteRows(rows);
});

// ==================== REMINDER SYSTEM ====================

// อ่านรายชื่อ user จาก Sheet "Users"
async function getRegisteredUsers() {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${USERS_SHEET}!A:B`,
    });
    const rows = res.data.values || [];
    // ข้ามแถว header, return [{name, discordId}]
    return rows.slice(1).map(row => ({
      name: row[0] || '',
      discordId: row[1] || '',
    })).filter(u => u.name && u.discordId);
  } catch (err) {
    console.error('❌ อ่าน Users sheet ไม่ได้:', err.message);
    return [];
  }
}

// เช็คว่าใครยังไม่กรอกข้อมูลวันนี้
async function checkMissingUsers() {
  const headers = await getHeaders();
  if (headers.length === 0) return;

  const users = await getRegisteredUsers();
  if (users.length === 0) return;

  // อ่านข้อมูลทั้งหมดจาก Sheet1
  const data = await readSheet();
  if (data.length <= 1) {
    // มีแค่ header = ยังไม่มีใครกรอกเลย
    return users;
  }

  // หา index ของคอลัมน์ วันที่ และ ชื่อ
  const dateColIdx = headers.findIndex(h => h.toLowerCase().includes('dd') || h.toLowerCase().includes('date') || h.includes('/'));
  const nameColIdx = headers.findIndex(h => h.toLowerCase() === 'name' || h.toLowerCase().includes('name'));

  if (dateColIdx === -1 || nameColIdx === -1) {
    console.log('⚠️ หาคอลัมน์วันที่หรือชื่อไม่เจอ');
    return [];
  }

  // วันนี้ - normalize เป็น d/m/yyyy (ไม่มี 0 นำหน้า) เพื่อเปรียบเทียบ
  const now = new Date();
  const todayDay = now.getDate();
  const todayMonth = now.getMonth() + 1;
  const todayYearAD = now.getFullYear();
  const todayYearBE = todayYearAD + 543;

  // ฟังก์ชัน normalize วันที่ให้เป็น d/m format (ลบ 0 นำหน้า)
  function normalizeDate(dateStr) {
    const parts = dateStr.split('/');
    if (parts.length < 3) return dateStr;
    const d = parseInt(parts[0]);
    const m = parseInt(parts[1]);
    const y = parseInt(parts[2]);
    return `${d}/${m}/${y}`;
  }

  const todayAD = `${todayDay}/${todayMonth}/${todayYearAD}`;
  const todayBE = `${todayDay}/${todayMonth}/${todayYearBE}`;

  // หาว่าใครกรอกวันนี้แล้ว
  const filledNames = new Set();
  for (let i = 1; i < data.length; i++) {
    const rowDate = normalizeDate((data[i][dateColIdx] || '').trim());
    const rowName = (data[i][nameColIdx] || '').trim();
    if (rowDate === todayAD || rowDate === todayBE) {
      filledNames.add(rowName.toLowerCase());
    }
  }

  // หาคนที่ยังไม่กรอก
  return users.filter(u => !filledNames.has(u.name.toLowerCase()));
}

// ส่งแจ้งเตือนใน Discord
async function sendReminder() {
  const missing = await checkMissingUsers();
  if (!missing || missing.length === 0) {
    console.log('✅ ทุกคนกรอกข้อมูลครบแล้ววันนี้');
    return;
  }

  const channel = client.channels.cache.get(CHANNEL_ID);
  if (!channel) return;

  const mentions = missing.map(u => `<@${u.discordId}>`).join(' ');
  await channel.send(`⏰ อย่าลืมกรอกข้อมูลวันนี้นะงับค้าบบบ:\n${mentions}`);
  console.log(`⏰ แจ้งเตือน ${missing.length} คน`);
}

// ตั้งเวลาเช็คทุกนาที ถ้าถึงเวลาที่กำหนดก็แจ้งเตือน
let reminderSentToday = false;
setInterval(() => {
  const now = new Date();
  if (now.getHours() === REMINDER_HOUR && now.getMinutes() === REMINDER_MINUTE) {
    if (!reminderSentToday) {
      reminderSentToday = true;
      sendReminder();
    }
  }
  // reset flag ตอนเที่ยงคืน
  if (now.getHours() === 0 && now.getMinutes() === 0) {
    reminderSentToday = false;
  }
}, 10 * 1000);

// ==================== START BOT ====================

client.login(DISCORD_TOKEN);
