require('dotenv').config();
process.env.TZ = 'Asia/Bangkok';


const { Client, GatewayIntentBits } = require('discord.js');
const OpenAI = require('openai');

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent
  ]
});

const CHANNEL_ID = '1168883889289302028';   //Consult Tong

const schedules = [
  //{ hour: 04, minute: 00, message: '@everyone ถึงเวลาสอบถามคุณครูดูแลน้องๆ ค้าบบบ' }
  { hour: 4, minute: 0, second: 10, message: '@everyone อ้วงงงงงงง ไปกินข้าวกันค้าบบบ' },
  { hour: 4, minute: 0, second: 10, message: '<@1131175352018944050> อ้วงงงงงงง ทำไรอยู่' }

  // { hour: 18, minute: 30, message: '@everyone ถึงเวลาเลิกงานแล้วน้อง กลับบ้านๆ' }
];

const SYSTEM_PROMPT = `
คุณคือ Discord Bot ชื่อ Bot PON
คุณคือผู้ชาย
หน้าที่:
- ตอบเป็นภาษาไทย
- ตอบคำถามเกี่ยวกับการทำงานและชีวิตประจำวัน
- ตอบแบบสั้น ไม่เกิน 5 บรรทัด
- ใช้ emoji เล็กน้อย
- ถ้ามีการถามเล่น ให้ตอบเล่นได้
- ถ้ามีคำถามไม่เหมาะสม ให้ปฏิเสธสุภาพ
`;

// ฟังก์ชันส่งข้อความ
async function sendMessage(text) {
  try {
    const channel = await client.channels.fetch(CHANNEL_ID);
    await channel.send({
      content: text,
      allowedMentions: { parse: ['everyone'] }
    });
    console.log('✅ ส่งแล้ว:', text);
  } catch (err) {
    console.error('❌ ส่งไม่สำเร็จ:', err.message);
  }
}

// คำนวณเวลาถึง 17:17
function scheduleDaily(hour, minute, second = 0, message) {
  const now = new Date();
  const target = new Date();

  target.setHours(hour, minute, second, 0);

  if (target <= now) {
    target.setDate(target.getDate() + 1);
  }

  const delay = target - now;

  console.log(
    `⏳ [${hour}:${minute}:${second}] จะส่งใน ${(delay / 1000).toFixed(0)} วินาที`
  );

  setTimeout(() => {
    sendMessage(message);

    setInterval(() => {
      sendMessage(message);
    }, 24 * 60 * 60 * 1000);
  }, delay);
}


client.once('ready', () => {
  console.log(`🤖 Bot online as ${client.user.tag}`);

  for (const s of schedules) {
    scheduleDaily(s.hour, s.minute, s.second, s.message);
  }
});

client.on('messageCreate', async (message) => {
  // ❌ ไม่ตอบ bot ด้วยกันเอง
  if (message.author.bot) return;

  // ❌ ต้อง @ bot เท่านั้น
  if (!message.mentions.has(client.user)) return;

  try {
    // เอา @bot ออกก่อนส่งไปให้ AI
    const userQuestion = message.content
      .replace(`<@${client.user.id}>`, '')
      .replace(`<@!${client.user.id}>`, '')
      .trim();

    if (!userQuestion) {
      return message.reply('ถามอะไรหน่อยสิ 👀');
    }

    // แสดงสถานะกำลังพิมพ์
    await message.channel.sendTyping();

    const response = await openai.chat.completions.create({
      model: 'gpt-5-mini',
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: userQuestion }
      ]
    });
    
    

    const answer = response.choices[0].message.content;

    await message.reply(answer);

  } catch (err) {
    console.error(err);
    message.reply('❌ ตอนนี้ AI มีปัญหา ลองใหม่อีกครั้งนะ');
  }
});


client.login(process.env.DISCORD_TOKEN);