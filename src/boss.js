// Bridge to the "พ่อใหญ่" session in Claude Desktop through plain files (format: $BOSS_DIR/README.md).
//   inbox.jsonl   we append the owner's `!boss ...` commands; พ่อใหญ่ reads them (it owns inbox.cursor)
//   outbox.jsonl  พ่อใหญ่ appends messages for the owner; we post them to Discord (we own outbox.cursor)
// Nothing here runs through the bot's own Agent SDK.
import fs from 'node:fs';
import path from 'node:path';

const PREFIX = /^!boss(?:\s+|$)/i;

// Returns the command text for พ่อใหญ่, or null when the message is not meant for it.
// In the boss channel every message counts; elsewhere only `!boss <text>`.
export function parseBossCommand(content, { inBossChannel = false } = {}) {
  const text = String(content || '').trim();
  if (PREFIX.test(text)) return text.replace(PREFIX, '').trim();
  if (inBossChannel && text && !text.startsWith('!')) return text;
  return null;
}

export function appendInbox(dir, { id, from, text, ts = new Date().toISOString() }) {
  fs.mkdirSync(dir, { recursive: true });
  const line = JSON.stringify({ id: String(id), ts, from: String(from), text: String(text) });
  // One write per line with O_APPEND: lines from concurrent writers never interleave.
  fs.appendFileSync(path.join(dir, 'inbox.jsonl'), `${line}\n`);
  return line;
}

function readCursor(dir) {
  try {
    return fs.readFileSync(path.join(dir, 'outbox.cursor'), 'utf8').trim() || null;
  } catch {
    return null;
  }
}

export function writeCursor(dir, id) {
  const file = path.join(dir, 'outbox.cursor');
  fs.writeFileSync(`${file}.tmp`, `${id}\n`);
  fs.renameSync(`${file}.tmp`, file);
}

// Entries after outbox.cursor. A last line without "\n" may still be mid-write, so it waits for the next poll.
// If the cursor id is no longer in the file (outbox was rotated), everything in it counts as new.
export function readOutboxSince(dir) {
  let raw = '';
  try {
    raw = fs.readFileSync(path.join(dir, 'outbox.jsonl'), 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return { entries: [], skipped: [] };
    throw err;
  }
  const complete = raw.endsWith('\n') ? raw : raw.slice(0, raw.lastIndexOf('\n') + 1);
  const entries = [];
  const skipped = [];
  for (const line of complete.split('\n')) {
    if (!line.trim()) continue;
    try {
      const e = JSON.parse(line);
      if (!e || typeof e !== 'object' || !e.id) throw new Error('no id');
      entries.push({ ...e, id: String(e.id) });
    } catch {
      skipped.push(line);
    }
  }
  const cursor = readCursor(dir);
  const at = cursor ? entries.findIndex((e) => e.id === cursor) : -1;
  return { entries: entries.slice(at + 1), skipped };
}

export function formatOutbox(entry, ownerId) {
  const level = ['info', 'warn', 'alert'].includes(entry.level) ? entry.level : 'info';
  const tag = { info: '', warn: '**[เตือน]** ', alert: '**[ด่วน]** ' }[level];
  const mention = level === 'alert' && ownerId ? `<@${ownerId}> ` : '';
  const reply = entry.reply_to ? `\n-# ตอบคำสั่ง ${entry.reply_to}` : '';
  return {
    content: `${mention}**พ่อใหญ่:** ${tag}${String(entry.text ?? '')}${reply}`,
    allowedMentions: { users: mention ? [ownerId] : [] },
  };
}

// Posts each new outbox entry, advancing the cursor only after Discord accepted it,
// so a failed send is retried next tick instead of being lost.
const reported = new Set(); // unreadable lines are logged once, not every poll
export async function pumpOutbox(dir, post, log = console) {
  const { entries, skipped } = readOutboxSince(dir);
  for (const line of skipped) {
    if (reported.has(line)) continue;
    reported.add(line);
    log.error(`boss outbox: skipped unreadable line: ${line.slice(0, 200)}`);
  }
  let sent = 0;
  for (const entry of entries) {
    await post(entry);
    writeCursor(dir, entry.id);
    sent += 1;
  }
  return sent;
}

export function startOutboxPoller({ dir, intervalMs, post, log = console }) {
  let busy = false;
  const tick = async () => {
    if (busy) return;
    busy = true;
    try {
      await pumpOutbox(dir, post, log);
    } catch (err) {
      log.error(`boss outbox: ${err.message}`);
    } finally {
      busy = false;
    }
  };
  const timer = setInterval(tick, intervalMs);
  timer.unref?.();
  tick();
  return () => clearInterval(timer);
}
