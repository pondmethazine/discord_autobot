require('dotenv').config();
process.env.TZ = 'Asia/Bangkok';

const { Client, GatewayIntentBits, Partials, ActionRowBuilder, ButtonBuilder, ButtonStyle } = require('discord.js');
const { google } = require('googleapis');
const { GoogleGenAI } = require('@google/genai');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const holidays = require('./holidays.json');
const { isInfoQuestion, answerInfoQuestion } = require('./company-info');

// เช็คว่าวันที่ระบุเป็นวันหยุดหรือไม่ (เสาร์-อาทิตย์ + วันหยุดนักขัตฤกษ์)
function isHoliday(date = new Date()) {
  const dayOfWeek = date.getDay(); // 0=อาทิตย์, 6=เสาร์
  if (dayOfWeek === 0 || dayOfWeek === 6) return true;
  const d = date.getDate();
  const m = date.getMonth() + 1;
  return holidays.some(h => h.day === d && h.month === m);
}

// หา "วันทำงานล่าสุด" ย้อนหลังจากวันที่ระบุ (ข้ามเสาร์-อาทิตย์ + วันหยุด)
function getLastWorkingDay(from = new Date()) {
  const d = new Date(from);
  d.setDate(d.getDate() - 1);
  while (isHoliday(d)) {
    d.setDate(d.getDate() - 1);
  }
  return d;
}

// ==================== CONFIG ====================

// Pattern ที่ถือว่า "เป็น timesheet แน่นอน" — ใช้ที่เดียวกันทุกจุด อย่าแยก copy
// - มีตัวเลข + หน่วยเวลา (ชม, ชั่วโมง, hr, นาที, min)
// - ทั้งวัน / ครึ่งวัน / การลา
// - มีตัวเลขต่อท้ายบรรทัด (เช่น "Testcase power 5")
const CLEAR_TIMESHEET_PATTERN = /\d+\.?\d*\s*(ชม|ชั่วโมง|hr|h\b|นาที|min)|ทั้งวัน|ครึ่งวัน|^ลา|ลาป่วย|ลากิจ|ลาพักร้อน|ลางาน|\S+\s+\d+\.?\d*\s*$/im;

const genai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

const GEMINI_MODEL = 'gemini-2.5-flash';

// ปิด safety filter — แชนแนลนี้คุยกันกวนๆ ถ้าไม่ปิด Gemini จะบล็อกแล้ว .text เป็น undefined
const SAFETY_OFF = [
  { category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_NONE' },
  { category: 'HARM_CATEGORY_HATE_SPEECH', threshold: 'BLOCK_NONE' },
  { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_NONE' },
  { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_NONE' },
];

// ปิด thinking สำหรับงานเบา — เร็วขึ้น ~5 เท่า โดยคำตอบไม่ต่างกัน
const NO_THINKING = { thinkingBudget: 0 };

// ==================== OPENAI FALLBACK ====================
// Gemini เป็นตัวหลักเหมือนเดิม ถ้าเรียกไม่ผ่าน (คีย์หมดอายุ / quota / โดน filter)
// จะสลับไป OpenAI ให้อัตโนมัติ เพื่อไม่ให้ทั้งบอทหยุดทำงานเพราะคีย์เดียว
const OpenAI = require('openai');
const OPENAI_MODEL = 'gpt-4.1-mini';
const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

// เมื่อ Gemini พัง จะพักไม่เรียกซ้ำ 10 นาที (ไม่งั้นทุกข้อความต้องรอ Gemini timeout ก่อนเสมอ)
// พอครบ 10 นาทีจะลอง Gemini ใหม่เอง — ใส่คีย์ใหม่ใน .env แล้วรีสตาร์ทได้ทันทีเช่นกัน
const GEMINI_COOLDOWN_MS = 10 * 60 * 1000;
let geminiDownUntil = 0;

function isGeminiDown() {
  if (!process.env.GEMINI_API_KEY) return true;
  return Date.now() < geminiDownUntil;
}

function markGeminiDown(err) {
  const first = geminiDownUntil === 0 || Date.now() >= geminiDownUntil;
  geminiDownUntil = Date.now() + GEMINI_COOLDOWN_MS;
  if (first) {
    console.error(`⚠️ Gemini ใช้ไม่ได้ → สลับไป OpenAI (${OPENAI_MODEL}) 10 นาที: ${err.message}`);
  }
}

// ดึง JSON ออกจากคำตอบ เผื่อโมเดลห่อด้วย ```json ... ```
function parseJsonLoose(text) {
  const clean = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/```$/, '').trim();
  return JSON.parse(clean);
}

async function openaiJson(prompt) {
  if (!openai) throw new Error('ไม่มีทั้ง GEMINI_API_KEY และ OPENAI_API_KEY ที่ใช้ได้');
  const res = await openai.chat.completions.create({
    model: OPENAI_MODEL,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: 'ตอบกลับเป็น JSON object เท่านั้น ห้ามมีข้อความอื่นนอก JSON' },
      { role: 'user', content: prompt },
    ],
  });
  const text = res.choices?.[0]?.message?.content;
  if (!text) throw new Error('OpenAI ไม่ตอบกลับ');
  return parseJsonLoose(text);
}

async function openaiText(prompt) {
  if (!openai) throw new Error('ไม่มีทั้ง GEMINI_API_KEY และ OPENAI_API_KEY ที่ใช้ได้');
  const res = await openai.chat.completions.create({
    model: OPENAI_MODEL,
    messages: [{ role: 'user', content: prompt }],
  });
  const text = res.choices?.[0]?.message?.content;
  if (!text) throw new Error('OpenAI ไม่ตอบกลับ');
  return text.trim();
}

// เรียก Gemini ให้ตอบเป็น JSON object (fast = งานเบา ไม่ต้องคิดเยอะ)
// ถ้า Gemini พัง (คีย์หมดอายุ / โดน filter / quota) จะสลับไป OpenAI ให้อัตโนมัติ
async function geminiJson(prompt, { fast = false } = {}) {
  if (!isGeminiDown()) {
    try {
      const res = await genai.models.generateContent({
        model: GEMINI_MODEL,
        contents: prompt,
        config: {
          responseMimeType: 'application/json',
          safetySettings: SAFETY_OFF,
          ...(fast && { thinkingConfig: NO_THINKING }),
        },
      });
      if (!res.text) throw new Error('Gemini ไม่ตอบกลับ (อาจโดน filter บล็อก)');
      return JSON.parse(res.text);
    } catch (err) {
      markGeminiDown(err);
    }
  }
  return openaiJson(prompt);
}

// เรียก Gemini ให้ตอบเป็นข้อความล้วน (ใช้กับข้อความสั้นๆ ตอบเล่น)
async function geminiText(prompt) {
  if (!isGeminiDown()) {
    try {
      const res = await genai.models.generateContent({
        model: GEMINI_MODEL,
        contents: prompt,
        config: { safetySettings: SAFETY_OFF, thinkingConfig: NO_THINKING },
      });
      if (!res.text) throw new Error('Gemini ไม่ตอบกลับ (อาจโดน filter บล็อก)');
      return res.text.trim();
    } catch (err) {
      markGeminiDown(err);
    }
  }
  return openaiText(prompt);
}

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.DirectMessages
  ],
  partials: [Partials.Message, Partials.Channel]
});

// Google Sheets Config
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_NAME = 'ลงเวลางาน';
const MSG_ID_COL = 'Z'; // คอลัมน์ซ่อน message ID (ไกลมากไม่มีใครเห็น)
const AUTHOR_ID_COL = 'Y'; // คอลัมน์ซ่อน Discord User ID ของคนลง timesheet

// Discord Config
const DISCORD_TOKEN = process.env.DISCORD_TOKEN_SHEET;
//const CHANNEL_ID = '1481565840477388841';
//const CHANNEL_ID = '1483304085783318598';
const CHANNEL_ID = '1488904036999499910';
//const CHANNEL_ID = 'test';

// ==================== STANDBY MODE ====================
// บอทเริ่มมาในโหมด "หยุดทำงาน" — ไม่ลง Sheet ไม่ตอบ ไม่ react ไม่เตือน ไม่ส่ง DM ตามเวลา
// จะกลับมาทำงานเมื่อ "สมาชิก" (ไม่ใช่ admin) พิมพ์คำปลุก แล้วบอทตอบ "กลับมาทำงานแล้ว"
// admin พิมพ์คำปลุกเอง = ไม่ปลุก (บอทไม่ action อะไรเหมือนเดิม)
const WAKE_WORD = 'สาดไปสมาชิก';
const WAKE_REPLY = 'กลับมาทำงานแล้ว';
let botActive = false;

// DM Control Map (userId → channelId)
const dmControlMap = new Map();

// เก็บข้อมูลข้อความที่รอการตัดสินใจ (export หรือ timesheet)
// key = messageId, value = { message, mentionedUsers, content }
const pendingExportChoice = new Map();

// Reminder Config
const USERS_SHEET = 'Users';       // ชื่อ Sheet ที่เก็บรายชื่อ
// เวลาแจ้งเตือน - แจ้งเตือนซ้ำจนกว่าจะกรอกข้อมูล
// type: 'today' = เช็ควันนี้, 'yesterday' = เช็คเมื่อวาน
const REMINDER_TIMES = [
  { hour: 10, minute: 0,  type: 'yesterday' },  // 10:00 แจ้งเตือนคนที่ยังไม่กรอกเมื่อวาน
  { hour: 18, minute: 30, type: 'today' },
  { hour: 19, minute: 0,  type: 'today' },
  { hour: 20, minute: 0,  type: 'today' },
  { hour: 21, minute: 0,  type: 'today' },
  { hour: 22, minute: 0,  type: 'today' },
  { hour: 23, minute: 0,  type: 'today' },       // 5 ทุ่ม = รอบสุดท้ายของวันนี้
];

// ==================== SCHEDULED DM ====================

// ข้อความตั้งเวลาส่ง DM หาคนใดคนหนึ่ง (แยกจากระบบ timesheet ทั้งหมด)
// - เวลาเป็น Asia/Bangkok (ตั้ง TZ ไว้บรรทัดบนสุดของไฟล์แล้ว)
// - once: true = ส่งครั้งเดียวแล้วเลิก, false = ส่งทุกวันเวลานี้
// - skipHoliday: true = ไม่ส่งวันเสาร์-อาทิตย์และวันหยุดนักขัตฤกษ์
// - ถ้า message ยังเป็น placeholder จะไม่ส่ง (กันเผลอยิงข้อความเปล่าใส่คนจริง)
const SCHEDULED_DM_PLACEHOLDER = 'ใส่ข้อความที่จะส่งตรงนี้';
const SCHEDULED_DMS = [
  {
    hour: 23,
    minute: 10,
    userId: '1369270229481422908',   // natthasitart
    message: SCHEDULED_DM_PLACEHOLDER,
    once: true,
    skipHoliday: false,
  },
];

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

// เขียน message ID + Discord user ID ลงคอลัมน์ซ่อน
async function saveMessageId(rowNumber, messageId, authorId = '') {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!${AUTHOR_ID_COL}${rowNumber}:${MSG_ID_COL}${rowNumber}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[authorId, messageId]] },
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
//hello//

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
9.1 ⚠️ ถ้าบรรทัดเดียวมี project หลายอัน เช่น "เรียนทำฟอร์ม MSIG + SDA 8 ชั่วโมง" → แยกเป็นหลาย row ตามจำนวน project (1 project = 1 row)
    - Detail เดียวกันทุก row (เช่น "เรียนทำฟอร์ม" ทั้งคู่)
    - ชั่วโมงหารเท่าๆ กันตามจำนวน project เช่น 8 ชม. / 2 project = 4 ชม. ต่อ row
    - ตัวอย่าง: "เรียนทำฟอร์ม MSIG + SDA 8 ชม." → [{Detail:"เรียนทำฟอร์ม", Hour:"4", Project:"MSIG"}, {Detail:"เรียนทำฟอร์ม", Hour:"4", Project:"SDA"}]
10. ⚠️ Hour: ใส่เฉพาะเมื่อระบุชั่วโมง/นาทีชัดเจน เช่น "3 ชม.", "30 นาที" ถ้าเป็นหน่วยอื่น เช่น "5 รอบ" ไม่ใช่ชั่วโมง → ให้เอาไว้ใน Detail แทน
    - ⚠️ "ทั้งวัน" / "เต็มวัน" / "full day" = 8.5 ชม. (เวลาทำงาน 1 วันของบริษัท)
    - ⚠️ "ครึ่ง" ให้ใช้ .5 เสมอ เช่น "1 ชั่วโมงครึ่ง" = 1.5, "2 ชม.ครึ่ง" = 2.5, "ครึ่งวัน" = 4.25
    - ⚠️ format Hour เป็นตัวเลขเท่านั้น ไม่ต้องมีหน่วย เช่น "3 ชม." → "3", "1 ชั่วโมงครึ่ง" → "1.5", "ทั้งวัน" → "8.5" ห้ามใช้ 1.30 หรือ 1:30
    - ⚠️ ถ้าระบุเป็นนาที → แปลงเป็นชั่วโมงทศนิยม (หาร 60) เช่น "30 นาที" → "0.5", "15 นาที" → "0.25", "45 นาที" → "0.75", "10 นาที" → "0.17"
    - ⚠️ ถ้าตัวเลขลอยๆ ท้ายบรรทัดไม่มีหน่วย → ถือว่าเป็นชั่วโมง เช่น "Testcase power 5" → Detail="Testcase power", Hour="5", "Enconform 3" → Detail="Enconform", Hour="3"
    - ⚠️ ถ้าผสมกัน → บวกเข้าด้วยกัน เช่น "1 ชม. 30 นาที" → "1.5", "2 ชม. 15 นาที" → "2.25", "3 ชม. 45 นาที" → "3.75"
    - ⚠️ ถ้าผู้ใช้ระบุว่า "ลา" (ลาป่วย, ลากิจ, ลาพักร้อน, ลางาน, ไม่ว่าเหตุผลอะไร) → Hour = "" (ว่าง) และ Detail ใส่เหตุผลการลา เช่น "ลาป่วย" → Detail="ลาป่วย", Hour=""
    - ⚠️ ถ้าบรรทัดไหนไม่ได้ระบุชั่วโมงชัดเจน แต่สื่อความหมายว่า "เวลาที่เหลือของวัน" (เช่น "ที่เหลือ", "นอกนั้น", "เวลาที่เหลือ", "ช่วงที่เหลือ", "the rest" ฯลฯ) → ให้คำนวณจาก 8.5 ชม. (เวลาทำงานต่อวัน) ลบชั่วโมงที่ระบุไปแล้วในข้อความเดียวกัน เช่น ถ้ามี 2 + 5 = 7 ชม. → ที่เหลือ = 8.5 - 7 = 1.5
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
    const result = await geminiJson(prompt);
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
    return await geminiJson(prompt);
  } catch (err) {
    console.error('❌ AI detect intent ไม่ได้:', err.message);
    return { intent: 'data' };
  }
}

// ==================== SUPPORT OWNER / SUPPORT BOT ====================

// AI วิเคราะห์ข้อความจาก owner / support bot แล้วตอบกลับหรือแจ้ง admin
// AI วิเคราะห์ว่าข้อความเป็น timesheet จริงไหม (ละเอียด ไม่ใช่แค่ regex)
async function aiIsTimesheetMessage(messageText) {
  try {
    const prompt = `วิเคราะห์ข้อความนี้ว่าเป็นการลง timesheet/บันทึกเวลาทำงานหรือไม่

ข้อความ: "${messageText}"

ข้อความที่เป็น timesheet:
- มีชั่วโมง/นาที/ระยะเวลา เช่น "2 ชม.", "30 นาที", "ทั้งวัน"
- ระบุงาน+เวลา เช่น "ประชุม 1 ชม.", "ทำ report 3hr"
- ระบุงาน+ตัวเลขลอยๆ (ไม่มีหน่วย) เช่น "Testcase power 5", "Enconform 3" → ตัวเลข = ชั่วโมง
- หลายบรรทัดที่แต่ละบรรทัดลงท้ายด้วยตัวเลข เช่น "Testcase power 5\\nEnconform 3"
- การลา เช่น "ลาป่วย", "ลากิจ", "ลาพักร้อน"
- Onsite / เดินทาง / ประชุม + ระยะเวลา

ข้อความที่ไม่ใช่ timesheet:
- ทักทาย, คุยเล่น, เหน็บแนม เช่น "Testcase 8", "หิวข้าว", "555"
- การสนทนาทั่วไป
- คำถาม/คำขอ

ตอบเป็น JSON: {"is_timesheet": true/false}`;

    const result = await geminiJson(prompt, { fast: true });
    return result.is_timesheet === true;
  } catch (err) {
    console.error('❌ AI เช็ค timesheet ไม่ได้:', err.message);
    return false;
  }
}

async function aiAnalyzeOwnerMessage(messageText, authorRole, mentionedUsers, botMentioned = false) {
  const prompt = `คุณเป็น bot ที่คอยช่วย support ${authorRole} ในแชนแนล Discord ของบริษัท
โดยเน้นให้บรรยากาศสนุกๆ กวนๆ ไม่ต้องสุภาพ
ในแชนแนลนี้ทุกคนจะลง timesheet กัน บางทีคนในแชนแนลจะเหน็บแนมคนที่ยังไม่ลง

ข้อความที่ได้รับจาก ${authorRole}: "${messageText}"
คนที่ถูก mention: ${mentionedUsers.length > 0 ? mentionedUsers.map(u => u.username).join(', ') : 'ไม่มี'}
ถูก mention bot เราหรือไม่: ${botMentioned ? 'ใช่ (owner กำลังคุยกับ bot โดยตรง)' : 'ไม่'}

วิเคราะห์แล้วตอบเป็น JSON:
{
  "action": "timesheet" | "reply" | "notify_admin" | "ignore",
  "reply_text": "ข้อความตอบกลับ (ถ้า action=reply) ตอบแบบกวนๆ ไม่สุภาพ สั้นๆ มีอารมณ์ขัน",
  "notify_reason": "เหตุผลที่แจ้ง admin (ถ้า action=notify_admin)"
}

⚠️ กฎสำคัญ: ต้องวิเคราะห์ให้ละเอียดดีๆ ว่าข้อความนั้นพูดถึงเรื่องอะไร อย่า reply มั่วๆ

⚠️⚠️⚠️ กฎลำดับที่ 1 (ต้องเช็คก่อนเสมอ):
ถ้าข้อความมี pattern ของการลง timesheet → action=timesheet ทันที ห้ามตอบ reply หรือ ignore
Pattern ที่ชัดเจน:
- มีตัวเลข + หน่วยเวลา เช่น "8 ชม.", "2 ชม", "30 นาที", "3 hr", "1 ชั่วโมง" → timesheet
- "ทั้งวัน", "ครึ่งวัน", "full day" → timesheet
- "ลาพักร้อน", "ลาป่วย", "ลากิจ", "ลางาน" → timesheet
- ชื่องาน + เวลา เช่น "ประชุม 8 ชม.", "onsite 3 ชม.", "เรียน Netsuite 8 ชม." → timesheet

ตัวอย่าง (ต้อง action=timesheet ทุกเคส):
- "ประชุม 8 ชม." → timesheet ✅ (ห้าม reply!)
- "onsite 5 ชม." → timesheet ✅
- "ลาพักร้อน" → timesheet ✅

กฎอื่นๆ:
- ⚠️ ถ้าเป็น owner แล้ว mention bot เรา (ถูก mention bot=ใช่) → action=reply ตอบคำถาม/สนทนากับ owner ได้เลย สั้นๆ เป็นกันเอง กวนได้นิดหน่อย
- ⚠️ ถ้าเป็น owner แล้ว mention คนอื่น (ที่ไม่ใช่ bot) → action=reply ตอบกลับกวนๆ สนับสนุนสิ่งที่ owner พูดถึง เสริมดราม่าโดยใช้ข้อมูลที่ owner พูด
- ⚠️ ถ้าเป็น owner แล้วไม่ได้ mention ใครเลย (เช่น คุยเล่น, ทักทาย, บ่น, พูดเรื่องทั่วไป) → action=notify_admin เสมอ (ส่งไป admin วิเคราะห์)
- ⚠️ ถ้าเป็นการเหน็บแนม/ทวง/บ่นคนที่ยังไม่ลง timesheet (จาก support bot) → action=reply สนับสนุนกวนๆ เสริมดราม่า
- ⚠️ ถ้าเป็น support bot พิมพ์เรื่องทั่วไปที่ไม่เกี่ยว timesheet → action=ignore (ไม่ต้องตอบ bot อีกตัว)
- reply_text: ใช้ภาษาวัยรุ่น สั้นๆ เป็นกันเอง อาจมี emoji ได้ กวนๆ ได้`;

  try {
    return await geminiJson(prompt, { fast: true });
  } catch (err) {
    console.error('❌ AI วิเคราะห์ owner message ไม่ได้:', err.message);
    return { action: 'ignore' };
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

  // ส่งไฟล์ทาง DM ให้คนขอ (ไม่โชว์ในแชนแนล)
  await message.author.send({
    content: `📊 สรุปข้อมูล ${filtered.length} รายการ`,
    files: [{ attachment: filePath, name: fileName }],
  });
  await message.react('📨');

  // ลบไฟล์ temp
  fs.unlinkSync(filePath);
  console.log(`📊 Export ${filtered.length} รายการ → DM ${message.author.username}`);
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

  // ส่งทาง DM
  await message.author.send({
    content: `📊 สรุปข้อมูลแยกรายคน ${files.length} ไฟล์`,
    files,
  });
  await message.react('📨');

  // ลบไฟล์ temp
  tempPaths.forEach(p => fs.unlinkSync(p));
  console.log(`📊 Export แยก ${files.length} ไฟล์`);
}

// ==================== DISCORD BOT ====================

client.once('ready', async () => {
  console.log(`✅ Bot พร้อมใช้งาน: ${client.user.tag}`);
  console.log(`⏸️ โหมดหยุดทำงาน — รอสมาชิกพิมพ์ "${WAKE_WORD}" (admin พิมพ์ไม่ปลุก)`);
  console.log(`📊 เชื่อมต่อ Google Sheet ID: ${SPREADSHEET_ID}`);

  const headers = await getHeaders();
  console.log(`📋 Columns ใน Sheet: ${headers.join(', ')}`);
});

client.on('messageCreate', async (message) => {
  if (message.author.bot) return;

  // ==================== แกล้งเพื่อน ====================
  // เฉพาะ user ID: 1308030557338341477 เท่านั้นที่ใช้ได้
  const ADMIN_ID = '1308030557338341477';

  if (!message.guild && message.author.id === ADMIN_ID) {
    if (message.content.startsWith('!chat ')) {
      const channelId = message.content.match(/<#(\d+)>/)?.[1] || message.content.split(' ')[1];
      const target = client.channels.cache.get(channelId);
      if (target) {
        dmControlMap.set(message.author.id, channelId);
        message.reply(`🎭 เปิดโหมดควบคุม bot ในแชนแนล #${target.name}\nพิมพ์อะไรก็ได้ bot จะพูดตาม\nพิมพ์ \`!stop\` เพื่อหยุด`);
      } else {
        message.reply('❌ หาแชนแนลไม่เจอ ลองใส่ Channel ID ตรงๆ');
      }
      return;
    }
    if (message.content === '!stop') {
      dmControlMap.delete(message.author.id);
      message.reply('🛑 หยุดโหมดควบคุมแล้ว');
      return;
    }
    if (message.content.startsWith('!reply ')) {
      const match = message.content.match(/!reply\s+https:\/\/discord\.com\/channels\/(\d+)\/(\d+)\/(\d+)\s+([\s\S]+)/);
      if (!match) {
        message.reply('❌ ใช้: `!reply MESSAGE_LINK ข้อความ`\nวิธีเอา link: คลิกขวาข้อความ → Copy Message Link');
        return;
      }
      const [, , channelId, messageId, text] = match;
      try {
        const channel = client.channels.cache.get(channelId);
        const targetMsg = await channel.messages.fetch(messageId);
        await targetMsg.reply(text);
        message.react('✅');
      } catch (e) {
        message.reply('❌ reply ไม่ได้: ' + e.message);
      }
      return;
    }
    const targetChannelId = dmControlMap.get(message.author.id);
    if (targetChannelId) {
      const target = client.channels.cache.get(targetChannelId);
      if (target) {
        await target.send(message.content);
        message.react('✅');
      }
      return;
    }
    return;
  }
  // !say และ !reply ในแชนแนล (เฉพาะ admin)
  if (message.author.id === ADMIN_ID) {
    if (message.content.startsWith('!say ')) {
      const args = message.content.slice(5);
      const channelMention = args.match(/^<#(\d+)>\s*/);
      let targetChannel, text;
      if (channelMention) {
        targetChannel = client.channels.cache.get(channelMention[1]);
        text = args.replace(/^<#\d+>\s*/, '');
      } else {
        targetChannel = message.channel;
        text = args;
      }
      if (!text || !targetChannel) return;
      try { await message.delete(); } catch (e) {}
      await targetChannel.send(text);
      return;
    }
    if (message.content.startsWith('!reply ') && message.reference?.messageId) {
      const text = message.content.slice(7);
      if (!text) return;
      try {
        const targetMsg = await message.channel.messages.fetch(message.reference.messageId);
        try { await message.delete(); } catch (e) {}
        await targetMsg.reply(text);
      } catch (e) {
        console.error('❌ reply ไม่ได้:', e.message);
      }
      return;
    }
  }

  // ==================== ประตูโหมดหยุดทำงาน ====================
  // ทุกอย่างหลังบรรทัดนี้คือ "งาน" ของบอท — ถ้ายังไม่ถูกปลุกจะไม่ทำอะไรเลย
  if (!botActive) {
    // admin พิมพ์เอง = ไม่ปลุก (เหมือนเดิม: บอทไม่ action อะไรกับข้อความ admin)
    if (message.author.id === ADMIN_ID) return;
    if (message.content.replace(/\s+/g, '').includes(WAKE_WORD)) {
      botActive = true;
      console.log(`▶️ ถูกปลุกโดย ${message.author.username} (${message.author.id})`);
      try {
        await message.reply(WAKE_REPLY);
      } catch (err) {
        console.error('❌ ตอบข้อความปลุกไม่ได้:', err.message);
      }
    }
    return;   // ข้อความที่ปลุกก็จบแค่นี้ ไม่ลง Sheet
  }

  if (message.channel.id !== CHANNEL_ID) return;

  // ==================== ถาม-ตอบ: วันเกิดพนักงาน / วันหยุดบริษัท ====================
  // ⚠️ ห้ามกระทบ timesheet: เข้าเงื่อนไขนี้ได้เฉพาะข้อความที่
  //    1) ไม่เข้า CLEAR_TIMESHEET_PATTERN เลย (ไม่มีชั่วโมง/ไม่ใช่การลา/ไม่มีเลขท้ายบรรทัด) และ
  //    2) เป็นคำถามเรื่องวันเกิด/วันหยุดชัดเจน (หรือ tag bot มาถาม)
  //    ข้อความที่หลุดเงื่อนไขนี้เดิมก็จะโดน AI ตีเป็น "ไม่ใช่ timesheet" แล้วถูกทิ้งอยู่ดี
  if (!CLEAR_TIMESHEET_PATTERN.test(message.content)) {
    const askText = message.content.replace(/<@!?\d+>/g, '').trim();
    const askedBot = message.mentions.has(client.user);
    if (isInfoQuestion(askText, { botMentioned: askedBot })) {
      console.log('📅 คำถามวันเกิด/วันหยุด:', askText);
      try {
        await message.reply(await answerInfoQuestion(askText, geminiText));
      } catch (err) {
        console.error('❌ ตอบคำถามวันเกิด/วันหยุดไม่ได้:', err.message);
      }
      return; // จบที่นี่ ไม่ลง Sheet
    }
  }

  // ==================== SUPPORT OWNER / SUPPORT BOT / FAT GUY ====================
  const OWNER_ID = '1131175352018944050';
  const SUPPORT_BOT_ID = '1480443439706275922';
  const FAT_GUY_ID = '1369270229481422908';
  if (message.author.id === OWNER_ID || message.author.id === SUPPORT_BOT_ID) {
    const role = message.author.id === OWNER_ID ? 'owner ของ server' : 'bot แจ้งเตือน timesheet';
    const ownerMentionedUsers = [...message.mentions.users.values()];
    const botMentioned = message.mentions.has(client.user);

    // Fast path: ถ้ามี pattern timesheet ชัดเจน → ปล่อยผ่านไปเลย ไม่ต้องถาม AI
    const isOwnerTimesheet = CLEAR_TIMESHEET_PATTERN.test(message.content);

    if (isOwnerTimesheet) {
      console.log(`🎯 Owner timesheet fast path → ปล่อยให้ลง Sheet`);
      // ปล่อยผ่านไปทำ timesheet logic ปกติ (ไม่เรียก AI)
    } else {
      const result = await aiAnalyzeOwnerMessage(message.content, role, ownerMentionedUsers, botMentioned);
      console.log(`🎯 Owner/SupportBot analysis:`, result);

      if (result.action === 'timesheet') {
        // ปล่อยให้ระบบปกติลง timesheet ต่อ
      } else {
        if (result.action === 'reply' && result.reply_text) {
          await message.reply(result.reply_text);
        } else if (result.action === 'notify_admin') {
          try {
            const admin = await client.users.fetch(ADMIN_ID);
            await admin.send(`📬 มีคำขอ/คำถามจาก ${message.author.username}\n💬 "${message.content}"\n🤔 ${result.notify_reason || ''}\n🔗 ${message.url}`);
          } catch (e) {}
        }
        return; // ไม่บันทึกลง Sheet
      }
    }
  }

  // ==================== FAT GUY: ตอบเฉพาะข้อความที่ไม่ใช่ timesheet ด้วยคำว่า "อ้วน" ====================
  if (message.author.id === FAT_GUY_ID) {
    const looksLikeTimesheet = /ชม\.?|ชั่วโมง|นาที|ทั้งวัน|ครึ่งวัน|ลา|ประชุม|เดินทาง|onsite|\d+\s*hr/i.test(message.content);
    if (!looksLikeTimesheet) {
      try {
        const prompt = `ผู้ใช้พิมพ์ข้อความ: "${message.content}"

งานของคุณ: ตอบกลับข้อความนี้ 1 ประโยค
- ตอบกลับแบบไม่สุภาพ กวนๆ
- ⚠️ ต้องลงท้ายด้วยคำที่เกี่ยวกับ "อ้วน" เสมอ เช่น "อ้วน", "อ้วง", "ไอ้อ้วน", "ไอ้อ้วง", "เจ้าอ้วน"
- ตอบสั้นๆ 1 ประโยค ไม่ต้องอธิบาย`;
        const reply = await geminiText(prompt);
        await message.reply(reply);
      } catch (err) {
        console.error('❌ Fat guy reply ไม่ได้:', err.message);
      }
      return; // ไม่ลง Sheet
    }
    // ถ้าเป็น timesheet → ปล่อยผ่านให้ระบบลง Sheet ตามปกติ (ไม่ต้อง reply)
  }

  // ==================== ตรวจจับคำว่า "อ้วน" ====================
  if (message.content.match(/อ้วน|อ้วง/i)) {
    try {
      const admin = await client.users.fetch(ADMIN_ID);
      await admin.send(`👤 ${message.author.username}\n💬 "${message.content}"\n🔗 ${message.url}`);
    } catch (e) {}
  }

  // ถ้าเป็นการ reply ข้อความคนอื่น → ไม่บันทึกลง Sheet
  if (message.reference?.messageId) return;

  // ให้ AI ตรวจจับว่าเป็นคำสั่ง export หรือกรอกข้อมูล
  // ตัด mention ออกจากข้อความเพื่อดู content จริง
  const cleanContent = message.content.replace(/<@!?\d+>/g, '').trim();
  const mentionedUsers = [...message.mentions.users.values()];

  // ถ้า admin/owner พิมพ์ข้อความที่มี keyword ของ export → ถามก่อนว่าจะ export หรือลง timesheet
  const canExport = message.author.id === ADMIN_ID || message.author.id === OWNER_ID;
  const exportKeywords = /สรุป|export|ดึงข้อมูล|timesheet|รายงาน|report/i;

  if (canExport && exportKeywords.test(cleanContent)) {
    // เก็บข้อมูลไว้ใน pendingExportChoice
    pendingExportChoice.set(message.id, { message, mentionedUsers });

    // สร้างปุ่มให้เลือก
    const row = new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(`choice_export_${message.id}`)
        .setLabel('📊 ออก Report')
        .setStyle(ButtonStyle.Primary),
      new ButtonBuilder()
        .setCustomId(`choice_timesheet_${message.id}`)
        .setLabel('📝 ลง Timesheet')
        .setStyle(ButtonStyle.Secondary),
    );

    await message.reply({
      content: '🤔 ต้องการออก report หรือลง timesheet?',
      components: [row],
    });

    // เคลียร์ pending หลัง 5 นาที (กันข้อมูลค้าง)
    setTimeout(() => pendingExportChoice.delete(message.id), 5 * 60 * 1000);
    return;
  }

  // ข้อ 5: ข้อความของ admin ไม่ต้องลง timesheet (แต่ admin ยังใช้ export buttons ได้ข้างบน)
  if (message.author.id === ADMIN_ID) return;

  // ข้อ 4: วิเคราะห์ให้ละเอียดว่าข้อความเป็นเรื่อง timesheet จริงๆ ไหม
  // Fast path: ถ้ามี pattern เวลาชัดเจน → ถือว่าเป็น timesheet ทันที ไม่ต้องถาม AI
  // - มีหน่วยเวลาชัดๆ (ชม, hr, นาที)
  // - ทั้งวัน / ครึ่งวัน / ลา
  // - มีตัวเลขต่อท้ายคำ/บรรทัด (เช่น "Testcase power 5", "Enconform 3") → น่าจะเป็นชั่วโมง
  let isTimesheet = CLEAR_TIMESHEET_PATTERN.test(message.content);

  // ถ้าไม่ match pattern ชัดเจน ค่อยให้ AI วิเคราะห์
  if (!isTimesheet) {
    isTimesheet = await aiIsTimesheetMessage(message.content);
  }

  if (!isTimesheet) {
    console.log('⚠️ ไม่ใช่ข้อความ timesheet:', message.content);
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
      await saveMessageId(rowNumber, message.id, message.author.id);
    } else {
      allSuccess = false;
    }
  }

  message.react(allSuccess ? '✅' : '❌');
});

// เมื่อมีการแก้ไขข้อความ → ให้ AI วิเคราะห์ใหม่แล้วอัพเดท
client.on('messageUpdate', async (oldMessage, newMessage) => {
  if (!botActive) return;
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
      if (newRow) await saveMessageId(newRow, newMessage.id, newMessage.author.id);
    }
    console.log(`📝 ลบ ${allRows.length} แถวเก่า → สร้าง ${aiRows.length} แถวใหม่`);
  }

  newMessage.react('📝');
});

// เมื่อมีการลบข้อความ → ลบทุกแถวที่เกี่ยวข้องใน Sheet
client.on('messageDelete', async (message) => {
  if (!botActive) return;
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

// เช็คว่าใครยังไม่กรอกข้อมูล (รับ targetDate เพื่อเช็ควันไหนก็ได้)
async function checkMissingUsers(targetDate = new Date()) {
  const headers = await getHeaders();
  if (headers.length === 0) return;

  const users = await getRegisteredUsers();
  if (users.length === 0) return;

  const data = await readSheet();
  if (data.length <= 1) return users;

  const dateColIdx = headers.findIndex(h => h.toLowerCase().includes('dd') || h.toLowerCase().includes('date') || h.includes('/'));
  const nameColIdx = headers.findIndex(h => h.toLowerCase() === 'name' || h.toLowerCase().includes('name'));

  if (dateColIdx === -1 || nameColIdx === -1) {
    console.log('⚠️ หาคอลัมน์วันที่หรือชื่อไม่เจอ');
    return [];
  }

  const day = targetDate.getDate();
  const month = targetDate.getMonth() + 1;
  const yearAD = targetDate.getFullYear();
  const yearBE = yearAD + 543;

  function normalizeDate(dateStr) {
    const parts = dateStr.split('/');
    if (parts.length < 3) return dateStr;
    return `${parseInt(parts[0])}/${parseInt(parts[1])}/${parseInt(parts[2])}`;
  }

  const checkAD = `${day}/${month}/${yearAD}`;
  const checkBE = `${day}/${month}/${yearBE}`;

  // index ของคอลัมน์ AUTHOR_ID_COL (Y = index 24)
  const authorColIdx = AUTHOR_ID_COL.charCodeAt(0) - 65;

  // เก็บ Discord ID ของคนที่กรอกวันนี้ (match จาก ID ที่ซ่อนในคอลัมน์ Y)
  const filledIds = new Set();
  for (let i = 1; i < data.length; i++) {
    const rowDate = normalizeDate((data[i][dateColIdx] || '').trim());
    const rowAuthorId = (data[i][authorColIdx] || '').trim();
    if ((rowDate === checkAD || rowDate === checkBE) && rowAuthorId) {
      filledIds.add(rowAuthorId);
    }
  }

  // match จาก Discord ID (แม่นยำกว่า username เพราะ ID ไม่เปลี่ยน)
  return users.filter(u => !filledIds.has(u.discordId));
}

// ส่งแจ้งเตือนใน Discord
async function sendReminder(type = 'today') {
  if (!botActive) return;   // ยังไม่ถูกปลุก → ไม่แจ้งเตือน
  // ถ้าวันนี้เป็นวันหยุด → ไม่แจ้งเตือนอะไรเลย (รวมการแจ้งเตือนย้อนหลังด้วย)
  if (isHoliday()) return;

  const channel = client.channels.cache.get(CHANNEL_ID);
  if (!channel) return;

  if (type === 'yesterday') {
    // หาวันทำงานล่าสุด (จันทร์ → ศุกร์, พุธหลังหยุด → อังคาร ฯลฯ)
    const lastWorkingDay = getLastWorkingDay();
    const missing = await checkMissingUsers(lastWorkingDay);
    const dateStr = `${lastWorkingDay.getDate()}/${lastWorkingDay.getMonth() + 1}/${lastWorkingDay.getFullYear()}`;
    const dayNames = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    const lastDayName = dayNames[lastWorkingDay.getDay()];

    // เช็คว่า "วันทำงานล่าสุด" = เมื่อวาน (1 วันก่อน) หรือข้ามวันหยุดมา
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    const lastWd = new Date(lastWorkingDay);
    lastWd.setHours(0, 0, 0, 0);
    const isYesterday = yesterday.getTime() === lastWd.getTime();

    // คำที่ใช้เรียก: "เมื่อวาน" ถ้าวันถัดไปเป็น 1 วัน, ไม่งั้นใช้ "วันXXX"
    const dayLabel = isYesterday ? `เมื่อวาน (${dateStr})` : `วัน${lastDayName}ที่ ${dateStr}`;

    if (!missing || missing.length === 0) {
      // ทุกคนกรอกครบ → ส่งข้อความขอบคุณ (ให้ AI สร้างข้อความ random)
      try {
        const prompt = `สร้างข้อความขอบคุณทุกคนที่ลง timesheet ของ${dayLabel} ครบแล้ว 1 ข้อความ
- ใช้ภาษาไทย สบายๆ เป็นกันเอง กวนๆ ได้
- 1-2 ประโยค สั้นๆ
- ใส่ emoji ได้ 1-2 ตัว
- หลากหลาย ไม่ซ้ำเดิม
- ถ้าเป็น "เมื่อวาน" ให้ใช้คำว่า "เมื่อวาน" ไม่ต้องระบุชื่อวัน
- ตอบเฉพาะข้อความ ไม่ต้องอธิบาย`;
        await channel.send(await geminiText(prompt));
      } catch (e) {
        await channel.send(`🎉 ทุกคนกรอก timesheet ของ${dayLabel} ครบแล้ว ขอบคุณครับ!`);
      }
      console.log(`✅ ทุกคนกรอก ${dateStr} ครบ → ส่งข้อความขอบคุณ`);
      return;
    }
    const mentions = missing.map(u => `<@${u.discordId}>`).join(' ');
    await channel.send(`⏰ ยังไม่ได้กรอกข้อมูลของ${dayLabel}:\n${mentions}\nกรุณากรอกย้อนหลังด้วยนะครับ!`);
    console.log(`⏰ แจ้งเตือนย้อนหลัง (${dateStr}) ${missing.length} คน`);
  } else {
    // เช็ควันนี้
    const missing = await checkMissingUsers();
    if (!missing || missing.length === 0) {
      console.log('✅ ทุกคนกรอกข้อมูลครบแล้ววันนี้');
      return;
    }
    const mentions = missing.map(u => `<@${u.discordId}>`).join(' ');
    await channel.send(`⏰ อย่าลืมกรอกข้อมูลวันนี้นะค้าบบบ:\n${mentions}`);
    console.log(`⏰ แจ้งเตือนวันนี้ ${missing.length} คน`);
  }
}

// ตั้งเวลาเช็คทุก 10 วินาที ถ้าถึงเวลาที่กำหนดก็แจ้งเตือน
// เก็บ flag ว่าเวลาไหนส่งไปแล้ว (reset ทุกวัน)
const reminderSentFlags = new Set();
setInterval(() => {
  if (!botActive) return;   // ยังไม่ถูกปลุก → ไม่แตะ flag ด้วย จะได้ยิงรอบที่เหลือได้หลังถูกปลุก
  const now = new Date();
  const h = now.getHours();
  const m = now.getMinutes();

  // reset flags ตอนเที่ยงคืน (00:01)
  if (h === 0 && m === 1) {
    reminderSentFlags.clear();
  }

  // เช็คทุกเวลาที่ตั้งไว้
  for (const time of REMINDER_TIMES) {
    if (h === time.hour && m === time.minute) {
      const key = `${time.hour}:${time.minute}`;
      if (!reminderSentFlags.has(key)) {
        reminderSentFlags.add(key);
        sendReminder(time.type || 'today');
      }
    }
  }
}, 10 * 1000);

// ==================== SCHEDULED DM RUNNER ====================

// ส่ง DM 1 job — แยกออกมาเป็นฟังก์ชันเพื่อให้เทสได้โดยไม่ต้องรอเวลาจริง
async function sendScheduledDm(job) {
  const label = `${job.hour}:${String(job.minute).padStart(2, '0')}`;
  if (!job.message || job.message === SCHEDULED_DM_PLACEHOLDER) {
    console.warn(`⚠️ ข้าม scheduled DM ${label} — ยังไม่ได้ใส่ข้อความ`);
    return false;
  }
  if (job.skipHoliday && isHoliday()) {
    console.log(`⏭️ ข้าม scheduled DM ${label} — วันหยุด`);
    return false;
  }
  try {
    const user = await client.users.fetch(job.userId);
    await user.send(job.message);
    console.log(`📨 ส่ง DM ${label} ให้ ${user.username} (${job.userId}) แล้ว`);
    return true;
  } catch (err) {
    // ที่เจอบ่อยสุดคือ 50007 = ปลายทางปิดรับ DM จากคนในเซิร์ฟเวอร์
    console.error(`❌ ส่ง DM ให้ ${job.userId} ไม่สำเร็จ: ${err.message}`);
    return false;
  }
}

// job แบบ once ต้องจำข้ามการ restart ไม่งั้นบอทรีสตาร์ททีก็ยิงซ้ำอีกวัน → เก็บลงไฟล์
const SCHEDULED_DM_STATE_FILE = path.join(__dirname, 'scheduled-dm-sent.json');

function jobKey(job) {
  return `${job.userId}@${job.hour}:${String(job.minute).padStart(2, '0')}`;
}

function loadScheduledDmDone() {
  try {
    return new Set(JSON.parse(fs.readFileSync(SCHEDULED_DM_STATE_FILE, 'utf8')));
  } catch {
    return new Set();   // ไม่มีไฟล์ = ยังไม่เคยส่ง
  }
}

function markScheduledDmDone(job) {
  scheduledDmDone.add(jobKey(job));
  try {
    fs.writeFileSync(SCHEDULED_DM_STATE_FILE, JSON.stringify([...scheduledDmDone], null, 2), 'utf8');
  } catch (err) {
    console.error(`❌ เขียน ${SCHEDULED_DM_STATE_FILE} ไม่ได้: ${err.message}`);
  }
}

// เช็คทุก 10 วินาทีเหมือน reminder ด้านบน — flag กันส่งซ้ำในนาทีเดียวกัน
const scheduledDmSentFlags = new Set();
const scheduledDmDone = loadScheduledDmDone();   // สำหรับ job ที่ once: true

setInterval(() => {
  if (!botActive) return;   // ยังไม่ถูกปลุก → ไม่ส่ง DM ตามเวลา
  const now = new Date();
  const h = now.getHours();
  const m = now.getMinutes();

  // reset flag ตอนเที่ยงคืน เพื่อให้ job รายวันส่งได้อีกในวันถัดไป
  if (h === 0 && m === 1) scheduledDmSentFlags.clear();

  SCHEDULED_DMS.forEach((job) => {
    if (h !== job.hour || m !== job.minute) return;

    const key = jobKey(job);
    if (job.once && scheduledDmDone.has(key)) return;
    if (scheduledDmSentFlags.has(key)) return;
    scheduledDmSentFlags.add(key);

    // mark เฉพาะตอนส่งสำเร็จจริง ถ้าส่งไม่ผ่านจะได้ลองใหม่พรุ่งนี้
    sendScheduledDm(job).then((sent) => {
      if (sent && job.once) markScheduledDmDone(job);
    });
  });
}, 10 * 1000);

// ==================== MORNING MESSAGE ====================

// ให้ AI สร้างข้อความให้กำลังใจตอนเช้า
// async function generateMorningMessage() {
//   try {
//     return await geminiText(`สร้างข้อความให้กำลังใจเพื่อนร่วมงานตอนเช้า 1 ข้อความ สั้นๆ กระชับ 1-2 ประโยค
// - ใช้ภาษาไทย สบายๆ เป็นกันเอง
// - ใส่ emoji ได้
// - ห้ามซ้ำกับคำว่า "สู้ๆ" ตรงๆ ให้หลากหลาย
// - อาจเป็นมุกตลก คำคม หรือให้กำลังใจ สลับกันไป
// - ตอบแค่ข้อความเท่านั้น ไม่ต้องอธิบาย`);
//   } catch (err) {
//     console.error('❌ สร้างข้อความเช้าไม่ได้:', err.message);
//     return '☀️ เช้าวันใหม่ ขอให้ทุกคนมีวันที่ดีนะครับ!';
//   }
// }

// ส่งข้อความให้กำลังใจตอนเช้า
// let morningSentToday = false;
// setInterval(async () => {
//   const now = new Date();
//   if (now.getHours() === 9 && now.getMinutes() === 0) {
//     if (!morningSentToday && !isHoliday()) {
//       morningSentToday = true;
//       const channel = client.channels.cache.get(CHANNEL_ID);
//       if (channel) {
//         const msg = await generateMorningMessage();
//         await channel.send(msg);
//         console.log('☀️ ส่งข้อความเช้า:', msg);
//       }
//     }
//   }
//   if (now.getHours() === 0 && now.getMinutes() === 1) {
//     morningSentToday = false;
//   }
// }, 10 * 1000);

// ==================== START BOT ====================

// ==================== BUTTON INTERACTION ====================

client.on('interactionCreate', async (interaction) => {
  if (!botActive) return;
  if (!interaction.isButton()) return;
  if (!interaction.customId.startsWith('choice_')) return;

  // แยก action และ messageId
  const [, action, originalMessageId] = interaction.customId.split('_');
  const pending = pendingExportChoice.get(originalMessageId);

  if (!pending) {
    await interaction.reply({ content: '⏰ หมดเวลา หรือเคยเลือกไปแล้ว', ephemeral: true });
    return;
  }

  // จำกัดให้คนเดิมที่ขอเท่านั้นถึงจะกดได้
  if (interaction.user.id !== pending.message.author.id) {
    await interaction.reply({ content: '❌ เฉพาะคนที่พิมพ์คำสั่งเท่านั้นที่กดได้', ephemeral: true });
    return;
  }

  const { message, mentionedUsers } = pending;
  pendingExportChoice.delete(originalMessageId);

  if (action === 'export') {
    await interaction.update({ content: '📊 กำลังออก report...', components: [] });

    // วิเคราะห์ intent ด้วย AI เพื่อดึง filters
    const intent = await aiDetectIntent(message.content, mentionedUsers);
    if (mentionedUsers.length > 0) {
      intent.filters = intent.filters || {};
      intent.filters.names = mentionedUsers.map(u => u.username);
    }
    if (intent.split_per_person && intent.filters?.names?.length > 1) {
      await exportSheetPerPerson(intent.filters, message);
    } else {
      await exportSheet(intent.filters || {}, message);
    }
  } else if (action === 'timesheet') {
    await interaction.update({ content: '📝 กำลังลง timesheet...', components: [] });

    // ลง timesheet ตามปกติ
    const headers = await getHeaders();
    if (headers.length === 0) return;
    const dropdowns = await getDropdownValues();
    const aiRows = await aiMapToColumns(headers, message.author.username, message.content, dropdowns);
    if (!aiRows || aiRows.length === 0) {
      await message.react('❌');
      return;
    }
    let allSuccess = true;
    for (const aiResult of aiRows) {
      const row = mapToRow(headers, aiResult);
      const rowNumber = await appendToSheet(headers, [row]);
      if (rowNumber) {
        await saveMessageId(rowNumber, message.id, message.author.id);
      } else {
        allSuccess = false;
      }
    }
    await message.react(allSuccess ? '✅' : '❌');
  }
});

client.login(DISCORD_TOKEN);
