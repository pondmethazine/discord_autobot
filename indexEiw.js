require('dotenv').config();
process.env.TZ = 'Asia/Bangkok';


const { Client, GatewayIntentBits } = require('discord.js');
const OpenAI = require('openai');
const fs = require('fs');

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.GuildVoiceStates // ⭐ ต้องมี
  ]
});


const CHANNEL_ID = '1464119403644846153';  //Eiw Discord

const schedules = [
  // { hour: 12, minute: 30, message: '@everyone พักเที่ยงแล้วววว ไปกินข้าวกันเถอะ ใครลงไปช้าจะโกรธแล้วนะ 😤😤' },
  // { hour: 12, minute: 30, second: 10, message: '<@1155698781673771079> อ้วงงงงงงง ไปกินข้าวกันค้าบบบ' },
  // { hour: 17, minute: 21, second: 10, message: '<@1155698781673771079> ไอ้อ้วงงงงงงง ทำไรอยู่' },
  // { hour: 17, minute: 21, second: 15, message: '<@1155698781673771079> ทำไรอยู่ค้าบบบ' },
  { hour: 18, minute: 30, message: '@everyone ถึงเวลาเลิกงานแล้วน้อง กลับบ้านๆ' },
  // { hour: 10, minute: 38, message: '<@1155698781673771079> คิดถึงพี่อิ๋วจุงเบบยย' }
  
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

const {
  joinVoiceChannel,
  createAudioPlayer,
  createAudioResource,
  AudioPlayerStatus,
  NoSubscriberBehavior
} = require('@discordjs/voice');

const ytdl = require('@distube/ytdl-core');

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

function cleanYouTubeUrl(url) {
  try {
    const u = new URL(url);
    const videoId = u.searchParams.get('v');
    if (!videoId) return null;
    return `https://www.youtube.com/watch?v=${videoId}`;
  } catch {
    return null;
  }
}

client.once('ready', () => {
  console.log(`🤖 Bot online as ${client.user.tag}`);

  for (const s of schedules) {
    scheduleDaily(s.hour, s.minute, s.second, s.message);
  }
});

// client.on('messageCreate', async (message) => {
//   // ❌ ไม่ตอบ bot ด้วยกันเอง
//   if (message.author.bot) return;

//   // ❌ ต้อง @ bot เท่านั้น
//   if (!message.mentions.has(client.user)) return;

//   try {
//     // เอา @bot ออกก่อนส่งไปให้ AI
//     const userQuestion = message.content
//       .replace(`<@${client.user.id}>`, '')
//       .replace(`<@!${client.user.id}>`, '')
//       .trim();

//     if (!userQuestion) {
//       return message.reply('ถามอะไรหน่อยสิ 👀');
//     }

//     // แสดงสถานะกำลังพิมพ์
//     await message.channel.sendTyping();

//     const response = await openai.chat.completions.create({
//       model: 'gpt-5-mini',
//       messages: [
//         { role: 'system', content: SYSTEM_PROMPT },
//         { role: 'user', content: userQuestion }
//       ]
//     });
    
    

//     const answer = response.choices[0].message.content;

//     await message.reply(answer);

//   } catch (err) {
//     console.error(err);
//     message.reply('❌ ตอนนี้ AI มีปัญหา ลองใหม่อีกครั้งนะ');
//   }
// });

client.on('messageCreate', async (message) => {
  if (message.author.bot) return;
  if (!message.mentions.has(client.user)) return;

  const content = message.content
    .replace(`<@${client.user.id}>`, '')
    .replace(`<@!${client.user.id}>`, '')
    .trim();

  // ======================
  // 🎵 PLAY MUSIC
  // ======================
  const player = createAudioPlayer({
    behaviors: {
      noSubscriber: NoSubscriberBehavior.Pause,
    },
  });
  
  // ป้องกัน bot crash (สำคัญมาก)
  player.on('error', error => {
    console.error('Audio player error:', error.message);
  });
  
  if (content.startsWith('เล่น')) {
    const urlMatch = content.match(
      /(https?:\/\/(?:www\.)?(?:youtube\.com\/watch\?v=|youtu\.be\/)[^\s]+)/i
    );
  
    if (!urlMatch) {
      return message.reply('❌ ใส่ลิงก์ YouTube มาด้วยนะ 🎵');
    }
  
    const url = urlMatch[1];
  
    if (!ytdl.validateURL(url)) {
      return message.reply('❌ ลิงก์ YouTube ไม่ถูกต้อง');
    }
  
    const voiceChannel = message.member.voice.channel;
    if (!voiceChannel) {
      return message.reply('🎧 ต้องเข้าห้องเสียงก่อนนะ');
    }
  
    const connection = joinVoiceChannel({
      channelId: voiceChannel.id,
      guildId: voiceChannel.guild.id,
      adapterCreator: voiceChannel.guild.voiceAdapterCreator,
    });
  
    // const agent = ytdl.createAgent({
    //   cookies: fs.readFileSync('./cookies.txt', 'utf8')
    // });
    
    const stream = ytdl(url, {
      filter: 'audioonly',
      quality: 'highestaudio',
      highWaterMark: 1 << 25,
    });
  
    const resource = createAudioResource(stream);
    player.play(resource);
    connection.subscribe(player);
  
    message.reply('🎶 กำลังเปิดเพลงให้เลย~');
  
    player.on(AudioPlayerStatus.Idle, () => {
      connection.destroy();
    });
  }
  
  
  


  // ======================
  // 🤖 AI CHAT
  // ======================
  if (!content) {
    return message.reply('ถามอะไรหน่อยสิ 👀');
  }

  try {
    await message.channel.sendTyping();

    const response = await openai.chat.completions.create({
      model: 'gpt-5-mini',
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content }
      ]
    });

    await message.reply(response.choices[0].message.content);
  } catch (err) {
    console.error(err);
    message.reply('❌ ตอนนี้ AI มีปัญหา ลองใหม่อีกครั้งนะ');
  }
});


client.login(process.env.DISCORD_TOKEN);