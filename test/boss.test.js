import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseBossCommand, appendInbox, readOutboxSince, formatOutbox, pumpOutbox, writeCursor } from '../src/boss.js';

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'boss-'));
const outLine = (o) => `${JSON.stringify(o)}\n`;
const quiet = { error() {} };

test('parseBossCommand: !boss prefix anywhere, plain text only in the boss channel', () => {
  assert.equal(parseBossCommand('!boss เช็ค session ที่ค้าง'), 'เช็ค session ที่ค้าง');
  assert.equal(parseBossCommand('  !BOSS   hello '), 'hello');
  assert.equal(parseBossCommand('!boss'), '');
  assert.equal(parseBossCommand('!bossy hi'), null);
  assert.equal(parseBossCommand('tttc-ocr-lcmd: แก้บั๊ก'), null);
  assert.equal(parseBossCommand('สรุปงานวันนี้', { inBossChannel: true }), 'สรุปงานวันนี้');
  assert.equal(parseBossCommand('!help', { inBossChannel: true }), null);
  assert.equal(parseBossCommand('', { inBossChannel: true }), null);
});

test('appendInbox writes one JSON line per command in the documented format', () => {
  const dir = tmp();
  appendInbox(dir, { id: '111', from: '42', text: 'หนึ่ง\nสองบรรทัด', ts: '2026-11-27T00:00:00.000Z' });
  appendInbox(dir, { id: '222', from: '42', text: 'สอง' });
  const lines = fs.readFileSync(path.join(dir, 'inbox.jsonl'), 'utf8').split('\n');
  assert.equal(lines.length, 3);
  assert.equal(lines[2], '');
  assert.deepEqual(JSON.parse(lines[0]), { id: '111', ts: '2026-11-27T00:00:00.000Z', from: '42', text: 'หนึ่ง\nสองบรรทัด' });
  assert.equal(JSON.parse(lines[1]).id, '222');
});

test('readOutboxSince: no file, no cursor, after cursor, rotated file, partial last line', () => {
  const dir = tmp();
  assert.deepEqual(readOutboxSince(dir).entries, []);

  const file = path.join(dir, 'outbox.jsonl');
  fs.writeFileSync(file, outLine({ id: 'a', text: '1' }) + outLine({ id: 'b', text: '2' }));
  assert.deepEqual(readOutboxSince(dir).entries.map((e) => e.id), ['a', 'b']);

  writeCursor(dir, 'a');
  assert.deepEqual(readOutboxSince(dir).entries.map((e) => e.id), ['b']);

  fs.appendFileSync(file, '{"id":"c","text":"half');
  assert.deepEqual(readOutboxSince(dir).entries.map((e) => e.id), ['b'], 'half-written line waits');
  fs.appendFileSync(file, '"}\n');
  assert.deepEqual(readOutboxSince(dir).entries.map((e) => e.id), ['b', 'c']);

  writeCursor(dir, 'gone');
  assert.deepEqual(readOutboxSince(dir).entries.map((e) => e.id), ['a', 'b', 'c'], 'unknown cursor = rotated');
});

test('readOutboxSince skips unreadable lines without blocking later ones', () => {
  const dir = tmp();
  fs.writeFileSync(path.join(dir, 'outbox.jsonl'), 'not json\n' + outLine({ text: 'no id' }) + outLine({ id: 'x', text: 'ok' }));
  const { entries, skipped } = readOutboxSince(dir);
  assert.deepEqual(entries.map((e) => e.id), ['x']);
  assert.equal(skipped.length, 2);
});

test('formatOutbox mentions the owner only for alert', () => {
  const info = formatOutbox({ id: '1', level: 'info', text: 'เสร็จแล้ว', reply_to: null }, '42');
  assert.equal(info.content, '**พ่อใหญ่:** เสร็จแล้ว');
  assert.deepEqual(info.allowedMentions.users, []);

  const warn = formatOutbox({ id: '2', level: 'warn', text: 'disk เหลือน้อย', reply_to: '111' }, '42');
  assert.match(warn.content, /^\*\*พ่อใหญ่:\*\* \*\*\[เตือน\]\*\* disk เหลือน้อย\n-# ตอบคำสั่ง 111$/);
  assert.ok(!warn.content.includes('<@42>'));

  const alert = formatOutbox({ id: '3', level: 'alert', text: 'prod ล่ม' }, '42');
  assert.ok(alert.content.startsWith('<@42> **พ่อใหญ่:** **[ด่วน]** prod ล่ม'));
  assert.deepEqual(alert.allowedMentions.users, ['42']);

  assert.equal(formatOutbox({ id: '4', level: 'weird', text: 't' }, '42').content, '**พ่อใหญ่:** t');
});

test('pumpOutbox posts new entries in order and advances the cursor only after a successful post', async () => {
  const dir = tmp();
  const file = path.join(dir, 'outbox.jsonl');
  fs.writeFileSync(file, outLine({ id: 'a', text: '1' }) + outLine({ id: 'b', text: '2' }) + outLine({ id: 'c', text: '3' }));
  const cursor = () => fs.readFileSync(path.join(dir, 'outbox.cursor'), 'utf8').trim();

  const posted = [];
  await assert.rejects(pumpOutbox(dir, async (e) => {
    if (e.id === 'b') throw new Error('discord down');
    posted.push(e.id);
  }, quiet));
  assert.deepEqual(posted, ['a']);
  assert.equal(cursor(), 'a');

  assert.equal(await pumpOutbox(dir, async (e) => posted.push(e.id), quiet), 2);
  assert.deepEqual(posted, ['a', 'b', 'c']);
  assert.equal(cursor(), 'c');

  assert.equal(await pumpOutbox(dir, async (e) => posted.push(e.id), quiet), 0, 'nothing new = nothing posted');
});
