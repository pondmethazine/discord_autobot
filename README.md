# work-assistant-bot

ผู้ช่วยงานผ่าน Discord: สั่งงานจากมือถือ บอทลงมือบน Mac เครื่องนี้ (repo, git, gcloud ของเครื่องนี้) ด้วย Claude Agent SDK

## ใช้ยังไง
- พิมพ์งานในห้อง HOME_CHANNEL_ID → บอทเปิดเธรดหนึ่งงาน เช่น `tttc-ocr-lcmd: แก้ DEF-032 ...`
  ใส่ชื่อ repo นำหน้า = เปิด git worktree + branch `bot/<id>` ให้ทันที (ไม่แตะ working copy ของเรา)
- แนบรูป/PDF ได้ · ตอบในเธรด = สั่งต่อ (resume session เดิม) · `!stop` `!close` `!status` `!repos` `!help`
- งานเสี่ยง (push, deploy, gcloud ที่เปลี่ยนของ, bq query, ต่อ DB, curl POST, แก้ไฟล์นอก worktree) → ปุ่มอนุมัติ
  กฎอยู่ที่ `src/policy.js` และเทสต์ที่ `test/policy.test.js`
- งานที่เกินกำลัง (login, ตัดสินใจกับลูกค้า, ของที่ไม่อยู่ในเครื่อง, แก้ 3 รอบไม่ผ่าน) → บอท "ส่งคืน" พร้อมขั้นตอนที่ต้องทำ

## ติดตั้ง
```
npm install
cp .env.example .env      # ใส่ DISCORD_TOKEN, OWNER_ID, HOME_CHANNEL_ID (+ ANTHROPIC_API_KEY)
npm test                  # กฎอนุมัติ
npm run try -- "repo: งาน" # ลองฝั่ง agent จาก terminal โดยไม่ผ่าน Discord
./scripts/install-launchd.sh   # รันค้างเป็น LaunchAgent, log ที่ data/bot.log
```

## ไฟล์
- `src/index.js` Discord: เธรด, คิว, ปุ่มอนุมัติ, สถานะสด, ไฟล์แนบ
- `src/agent.js` Agent SDK: system prompt, tools `use_repo` / `hand_back` / `list_repos`, งบต่อรอบ
- `src/policy.js` ตัดสินว่า tool call ไหนทำเองได้ ไหนต้องรออนุมัติ
- `src/repos.js` หา repo และสร้าง worktree (copy `.env*` ให้ด้วย)
- `data/` state.json (เธรด↔session), ไฟล์แนบ, log — ไม่ขึ้น git
