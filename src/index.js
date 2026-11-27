// Discord front-end: one message in the home channel = one task thread.
// Reply in the thread to continue, !stop to abort, !close to archive.
import fs from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  Client, GatewayIntentBits, Partials, ActionRowBuilder, ButtonBuilder, ButtonStyle, ChannelType,
} from 'discord.js';
import { config, assertDiscordConfig } from './config.js';
import { Store } from './store.js';
import { runTurn } from './agent.js';
import { createWorktree, listRepos, parseRepoPrefix } from './repos.js';
import { chunk, codeBlock, fmtDuration } from './text.js';

assertDiscordConfig();
const store = new Store(path.join(config.dataDir, 'state.json'));

const client = new Client({
  intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent],
  partials: [Partials.Channel],
});

// ---------- run queue (global concurrency limit) ----------
const running = new Map(); // threadId → { abort, startedAt }
const pendingInput = new Map(); // threadId → [text] (messages sent while a turn was running)
const waitQueue = []; // [{ threadId, thread, prompt }]

function schedule(job) {
  if (running.size < config.maxConcurrentRuns) startRun(job);
  else {
    waitQueue.push(job);
    job.thread.send(`คิวเต็ม (${running.size}/${config.maxConcurrentRuns}) รอคิวที่ ${waitQueue.length} คับ`).catch(() => {});
  }
}

function next() {
  while (running.size < config.maxConcurrentRuns && waitQueue.length) startRun(waitQueue.shift());
}

// ---------- approvals ----------
const approvals = new Map(); // nonce → resolve(bool)

async function askApproval(thread, { tool, input, reason, signal }) {
  const nonce = randomUUID().slice(0, 8);
  const detail = tool === 'Bash' ? String(input.command || '') : JSON.stringify(input, null, 2);
  const row = new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`appr:${nonce}:yes`).setLabel('อนุมัติ').setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId(`appr:${nonce}:no`).setLabel('ไม่อนุมัติ').setStyle(ButtonStyle.Danger),
  );
  const msg = await thread.send({
    content: `<@${config.ownerId}> ขออนุมัติ: **${reason}**\n${codeBlock(detail, 1500)}`,
    components: [row],
    allowedMentions: { users: [config.ownerId] },
  });

  const decision = await new Promise((resolve) => {
    const timer = setTimeout(() => done(false, 'หมดเวลา'), config.approvalTimeoutMin * 60_000);
    const onAbort = () => done(false, 'งานถูกหยุด');
    signal?.addEventListener('abort', onAbort, { once: true });
    function done(ok, why) {
      if (!approvals.has(nonce)) return;
      approvals.delete(nonce);
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
      resolve({ ok, why });
    }
    approvals.set(nonce, (ok) => done(ok, ok ? 'อนุมัติแล้ว' : 'ไม่อนุมัติ'));
  });

  await msg.edit({ content: `${msg.content}\n→ ${decision.why}`, components: [] }).catch(() => {});
  return decision.ok;
}

// ---------- live status message ----------
function makeStatus(thread) {
  const lines = [];
  let msg = null;
  let timer = null;
  let dirty = false;
  const started = Date.now();

  const render = () => `**กำลังทำ** (${fmtDuration(Date.now() - started)})\n${lines.slice(-8).map((l) => `- ${l}`).join('\n')}`.slice(0, 1990);
  async function flush() {
    timer = null;
    if (!dirty) return;
    dirty = false;
    try {
      if (!msg) msg = await thread.send(render());
      else await msg.edit(render());
    } catch (err) {
      console.error('status update failed:', err.message);
    }
  }
  return {
    push(line) {
      lines.push(line.replace(/\s+/g, ' ').slice(0, 220));
      dirty = true;
      if (!timer) timer = setTimeout(flush, 2500);
    },
    async close(finalLine) {
      if (timer) clearTimeout(timer);
      if (finalLine) lines.push(finalLine);
      dirty = true;
      await flush();
    },
  };
}

// ---------- one turn ----------
async function startRun({ threadId, thread, prompt }) {
  const abort = new AbortController();
  running.set(threadId, { abort, startedAt: Date.now() });
  const state = { ...(store.get(threadId) || {}) };
  const save = (patch) => Object.assign(state, store.update(threadId, patch));
  save({ status: 'running' });

  const status = makeStatus(thread);
  const typing = setInterval(() => thread.sendTyping().catch(() => {}), 8000);
  thread.sendTyping().catch(() => {});

  const ui = {
    status: (t) => status.push(t),
    text: (t) => status.push(`> ${t.split('\n')[0]}`),
    approve: (req) => askApproval(thread, req),
    handBack: async ({ blocker, tried, owner_action, prepared }) => {
      const body = [
        `<@${config.ownerId}> **ส่งคืนให้มึง** — งานนี้ต้องให้คนทำต่อ`,
        `**ติดที่:** ${blocker}`,
        `**ลองไปแล้ว:** ${tried}`,
        `**มึงต้องทำ:** ${owner_action}`,
        prepared ? `**เตรียมไว้แล้ว:** ${prepared}` : '',
        'ทำเสร็จแล้วตอบในเธรดนี้ บอทจะทำต่อจากเดิมคับ',
      ].filter(Boolean).join('\n');
      for (const part of chunk(body)) await thread.send({ content: part, allowedMentions: { users: [config.ownerId] } });
    },
  };

  let res;
  try {
    res = await runTurn({ taskId: threadId, prompt, state, save, ui, abort });
  } catch (err) {
    res = { isError: true, subtype: abort.signal.aborted ? 'aborted' : 'crash', result: '', costUsd: 0, turns: 0, durationMs: 0, error: err };
    if (!abort.signal.aborted) console.error(`[${threadId}] run crashed:`, err);
  } finally {
    clearInterval(typing);
    running.delete(threadId);
  }

  const total = (state.costUsd || 0) + (res.costUsd || 0);
  const nextStatus = res.handedBack ? 'handed_back' : res.subtype === 'aborted' ? 'stopped' : res.isError ? 'error' : 'done';
  save({ costUsd: total, status: nextStatus, lastRunAt: new Date().toISOString() });
  await status.close(nextStatus === 'done' ? 'เสร็จ' : nextStatus);

  const labels = {
    error_max_budget_usd: `ชนเพดานงบต่อรอบ ($${config.maxBudgetUsdPerRun}) ตอบในเธรดเพื่อให้ทำต่อ`,
    error_max_turns: `ชนเพดานจำนวนรอบ (${config.maxTurnsPerRun}) ตอบในเธรดเพื่อให้ทำต่อ`,
    aborted: 'หยุดแล้วตามคำสั่ง',
    crash: `บอทพังระหว่างทำ: ${res.error?.message || ''}`.slice(0, 500),
  };
  const body = res.result || labels[res.subtype] || `จบแบบ ${res.subtype}`;
  const footer = `-# ${res.subtype} · ${res.turns} รอบ · ${fmtDuration(res.durationMs)} · $${(res.costUsd || 0).toFixed(2)} (รวมเธรดนี้ $${total.toFixed(2)})`;
  const mention = res.handedBack ? '' : `<@${config.ownerId}> `;
  const parts = chunk(`${mention}${body}\n${footer}`);
  for (const part of parts) await thread.send({ content: part, allowedMentions: { users: [config.ownerId] } }).catch(() => {});

  const queued = pendingInput.get(threadId);
  if (queued?.length && res.subtype !== 'aborted') {
    pendingInput.delete(threadId);
    schedule({ threadId, thread, prompt: queued.join('\n\n') });
  } else pendingInput.delete(threadId);
  next();
}

// ---------- attachments ----------
async function saveAttachments(message, threadId) {
  if (!message.attachments.size) return [];
  const dir = path.join(config.dataDir, 'attachments', threadId);
  fs.mkdirSync(dir, { recursive: true });
  const saved = [];
  for (const a of message.attachments.values()) {
    const name = `${message.id}-${a.name.replace(/[^\w.\-ก-๙]/g, '_')}`;
    const res = await fetch(a.url);
    if (!res.ok) continue;
    fs.writeFileSync(path.join(dir, name), Buffer.from(await res.arrayBuffer()));
    saved.push(path.join(dir, name));
  }
  return saved;
}

function withAttachments(text, files) {
  if (!files.length) return text;
  return `${text}\n\nไฟล์แนบ (อ่านด้วย Read ได้เลย):\n${files.map((f) => `- ${f}`).join('\n')}`;
}

// ---------- commands ----------
function statusReport() {
  const rows = store.all().sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || '')).slice(0, 15);
  if (!rows.length) return 'ยังไม่มีงานคับ';
  const lines = rows.map((t) => {
    const live = running.has(t.threadId) ? 'running' : t.status;
    return `- <#${t.threadId}> ${live} · ${t.repo || '-'} · $${(t.costUsd || 0).toFixed(2)}`;
  });
  return `**งานล่าสุด** (กำลังรัน ${running.size}, รอคิว ${waitQueue.length})\n${lines.join('\n')}`;
}

const HELP = [
  '**วิธีใช้**',
  '- พิมพ์งานในห้องนี้ = เปิดเธรดใหม่หนึ่งงาน เช่น `tttc-ocr-lcmd: แก้ DEF-032 ...` (ใส่ชื่อ repo นำหน้าเพื่อเปิด worktree ทันที, ไม่ใส่ก็ได้)',
  '- แนบรูป/PDF ได้ บอทอ่านเอง',
  '- ตอบในเธรด = สั่งต่องานเดิม · `!stop` หยุด · `!close` ปิดเธรด',
  '- `!status` งานทั้งหมด · `!repos` รายชื่อ repo · `!help`',
].join('\n');

// ---------- events ----------
client.on('messageCreate', async (message) => {
  try {
    if (message.author.bot) return;
    if (message.author.id !== config.ownerId) return; // only the owner can drive this bot
    const ch = message.channel;
    const inHome = ch.id === config.homeChannelId;
    const inTaskThread = ch.isThread?.() && ch.parentId === config.homeChannelId;
    if (!inHome && !inTaskThread) return;
    const content = message.content.trim();

    if (content === '!help') return void (await message.reply(HELP));
    if (content === '!status') return void (await message.reply(statusReport()));
    if (content === '!repos') {
      const repos = listRepos();
      for (const part of chunk(`${repos.length} repo:\n${repos.join(', ')}`)) await message.reply(part);
      return;
    }

    if (inTaskThread) {
      const threadId = ch.id;
      if (content === '!stop') {
        const r = running.get(threadId);
        if (r) { r.abort.abort(); await message.react('🛑'); } else await message.reply('ไม่มีงานรันอยู่ในเธรดนี้คับ');
        return;
      }
      if (content === '!close') {
        running.get(threadId)?.abort.abort();
        store.update(threadId, { status: 'closed' });
        const wts = (store.get(threadId)?.worktrees || []).map((w) => `- ${w.dir} (${w.branch})`).join('\n');
        await message.reply(`ปิดงานแล้วคับ${wts ? `\nworktree ยังอยู่ ลบเองเมื่อไม่ใช้:\n${wts}` : ''}`);
        await ch.setArchived(true).catch(() => {});
        return;
      }
      if (!store.get(threadId)) store.update(threadId, { repo: null, title: ch.name });
      const files = await saveAttachments(message, threadId);
      const prompt = withAttachments(content || '(ดูไฟล์แนบ)', files);
      if (running.has(threadId)) {
        pendingInput.set(threadId, [...(pendingInput.get(threadId) || []), prompt]);
        await message.react('⏳');
        return;
      }
      if (waitQueue.some((j) => j.threadId === threadId)) {
        pendingInput.set(threadId, [...(pendingInput.get(threadId) || []), prompt]);
        await message.react('⏳');
        return;
      }
      schedule({ threadId, thread: ch, prompt });
      return;
    }

    // New task in the home channel.
    if (!content && !message.attachments.size) return;
    const { repo, text } = parseRepoPrefix(content);
    const title = (text || content || 'งานใหม่').split('\n')[0].slice(0, 90);
    const thread = await message.startThread({ name: repo ? `[${repo}] ${title}`.slice(0, 100) : title });
    const record = { repo, title, status: 'queued' };
    if (repo) {
      try {
        const wt = await createWorktree(repo, thread.id);
        Object.assign(record, { worktrees: [{ repo, dir: wt.dir, branch: wt.branch, base: wt.base }], cwd: wt.dir });
        await thread.send(`-# worktree ${wt.dir} · ${wt.branch} จาก ${wt.base}${wt.envCopied.length ? ` · copy ${wt.envCopied.join(', ')}` : ''}`);
      } catch (err) {
        await thread.send(`เปิด worktree ของ ${repo} ไม่ได้: ${err.message.split('\n')[0]}\nจะเริ่มจาก ${config.reposRoot} แทนคับ`);
      }
    }
    store.update(thread.id, record);
    const files = await saveAttachments(message, thread.id);
    schedule({ threadId: thread.id, thread, prompt: withAttachments(text || '(ดูไฟล์แนบ)', files) });
  } catch (err) {
    console.error('messageCreate failed:', err);
    message.reply(`พังตอนรับข้อความ: ${err.message}`.slice(0, 1900)).catch(() => {});
  }
});

client.on('interactionCreate', async (interaction) => {
  if (!interaction.isButton()) return;
  const m = interaction.customId.match(/^appr:(\w+):(yes|no)$/);
  if (!m) return;
  if (interaction.user.id !== config.ownerId) {
    return void (await interaction.reply({ content: 'เฉพาะเจ้าของเท่านั้นคับ', ephemeral: true }).catch(() => {}));
  }
  const resolve = approvals.get(m[1]);
  await interaction.deferUpdate().catch(() => {});
  if (resolve) resolve(m[2] === 'yes');
  else await interaction.followUp({ content: 'คำขอนี้หมดอายุแล้วคับ', ephemeral: true }).catch(() => {});
});

client.once('clientReady', async () => {
  console.log(`logged in as ${client.user.tag}; home channel ${config.homeChannelId}; repos root ${config.reposRoot}`);
  const home = await client.channels.fetch(config.homeChannelId).catch(() => null);
  if (!home || home.type !== ChannelType.GuildText) {
    console.error(`HOME_CHANNEL_ID ${config.homeChannelId} is not a text channel the bot can see`);
    return;
  }
  // Turns that were running when the process died cannot be resumed mid-flight; say so in each thread.
  for (const t of store.all().filter((x) => x.status === 'running' || x.status === 'queued')) {
    store.update(t.threadId, { status: 'interrupted' });
    const th = await client.channels.fetch(t.threadId).catch(() => null);
    th?.send(`<@${config.ownerId}> บอทรีสตาร์ทระหว่างทำงานนี้ งานหยุดกลางทาง ตอบในเธรดนี้เพื่อให้ทำต่อจากเดิมคับ`).catch(() => {});
  }
  await home.send(`ออนไลน์แล้วคับ · พิมพ์ \`!help\` ดูวิธีใช้`).catch(() => {});
});

for (const sig of ['SIGINT', 'SIGTERM']) {
  process.on(sig, () => {
    for (const r of running.values()) r.abort.abort();
    client.destroy().finally(() => process.exit(0));
  });
}

client.login(config.discordToken);
