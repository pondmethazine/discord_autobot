# work-assistant-bot

ผู้ช่วยงานผ่าน Discord: สั่งงานจากมือถือ บอทลงมือบน Mac เครื่องนี้ (repo, git, gcloud ของเครื่องนี้) ด้วย Claude Agent SDK

## ใช้ยังไง
- พิมพ์งานในห้อง HOME_CHANNEL_ID → บอทเปิดเธรดหนึ่งงาน เช่น `tttc-ocr-lcmd: แก้ DEF-032 ...`
  ใส่ชื่อ repo นำหน้า = เปิด git worktree + branch `bot/<id>` ให้ทันที (ไม่แตะ working copy ของเรา)
- แนบรูป/PDF ได้ · ตอบในเธรด = สั่งต่อ (resume session เดิม) · `!stop` `!close` `!status` `!repos` `!help`
- งานเสี่ยง (push, deploy, gcloud ที่เปลี่ยนของ, bq query, ต่อ DB, curl POST, แก้ไฟล์นอก worktree) → ปุ่มอนุมัติ
  กฎอยู่ที่ `src/policy.js` และเทสต์ที่ `test/policy.test.js`
- `!boss <ข้อความ>` (เจ้าของเท่านั้น) ส่งคำสั่งให้ session **พ่อใหญ่** ใน Claude Desktop ทางไฟล์ ไม่รันผ่าน Agent SDK ของบอท
  ดูรายละเอียดหัวข้อ "ต่อกับพ่อใหญ่" ด้านล่าง
- งานที่เกินกำลัง (login, ตัดสินใจกับลูกค้า, ของที่ไม่อยู่ในเครื่อง, แก้ 3 รอบไม่ผ่าน) → บอท "ส่งคืน" พร้อมขั้นตอนที่ต้องทำ

## ติดตั้ง
```
npm install
cp .env.example .env      # ใส่ DISCORD_TOKEN, OWNER_ID, HOME_CHANNEL_ID (+ ANTHROPIC_API_KEY)
npm test                  # กฎอนุมัติ
npm run try -- "repo: งาน" # ลองฝั่ง agent จาก terminal โดยไม่ผ่าน Discord
./scripts/install-launchd.sh   # รันค้างเป็น LaunchAgent, log ที่ data/bot.log
```

## ต่อกับพ่อใหญ่
คุยกันผ่านไฟล์ใน `BOSS_DIR` (ค่าเริ่มต้น `~/.claude/boss`, format อยู่ที่ `README.md` ในโฟลเดอร์นั้น)
- **สั่ง:** `!boss <ข้อความ>` ในห้องหลักหรือในเธรด → บอท append `{"id","ts","from","text"}` ลง `inbox.jsonl` แล้วตอบ 📨
  ถ้าตั้ง `BOSS_CHANNEL_ID` ทุกข้อความของเจ้าของในห้องนั้นถือเป็นคำสั่งถึงพ่อใหญ่ ไม่ต้องพิมพ์ `!boss`
- **แจ้งกลับ:** บอทอ่าน `outbox.jsonl` ทุก `BOSS_POLL_SEC` วินาที (ค่าเริ่มต้น 20) โพสต์บรรทัดที่อยู่หลัง `outbox.cursor`
  ลงห้อง `BOSS_CHANNEL_ID` (หรือ `HOME_CHANNEL_ID` ถ้าไม่ตั้ง) แล้วเลื่อน cursor หลังส่งสำเร็จเท่านั้น
  `level: alert` จะ mention เจ้าของ · บรรทัดที่ยังเขียนไม่จบ (ไม่มี `\n` ท้าย) รอรอบถัดไป · บรรทัดที่อ่านไม่ได้ข้ามและ log ไว้
- โค้ดอยู่ที่ `src/boss.js` เทสต์ที่ `test/boss.test.js`

## ไฟล์
- `src/index.js` Discord: เธรด, คิว, ปุ่มอนุมัติ, สถานะสด, ไฟล์แนบ
- `src/agent.js` Agent SDK: system prompt, tools `use_repo` / `hand_back` / `list_repos`, งบต่อรอบ
- `src/policy.js` ตัดสินว่า tool call ไหนทำเองได้ ไหนต้องรออนุมัติ
- `src/boss.js` ทางคุยกับ session พ่อใหญ่ (inbox/outbox)
- `src/repos.js` หา repo และสร้าง worktree (copy `.env*` ให้ด้วย)
- `data/` state.json (เธรด↔session), ไฟล์แนบ, log — ไม่ขึ้น git
