// Runs one turn of a task with the Claude Agent SDK and reports back through `ui`.
// A task = one Discord thread = one Claude session; follow-ups resume the same session.
import { query, createSdkMcpServer, tool } from '@anthropic-ai/claude-agent-sdk';
import { z } from 'zod';
import { config } from './config.js';
import { classify } from './policy.js';
import { createWorktree, listRepos } from './repos.js';

export const SYSTEM_APPEND = `
# บทบาทใน Discord
คุณคือผู้ช่วยงานของเจ้าของ (owner) ที่ถูกสั่งงานผ่าน Discord: เจ้าของอ่านจากมือถือและแค่นั่ง monitor
คุยภาษาไทย ลงท้าย คับ (ห้าม ครับ) สั้น กระชับ อ่านบนมือถือได้ ไม่ใช้ emoji

## ทุกข้อความเลือกทางใดทางหนึ่ง
1. ตอบเลย: คำถาม สรุป อธิบาย ถามสถานะ ไม่ต้องแตะโค้ดก็ตอบ
2. ทำให้เสร็จเอง: อ่านโค้ด/log, หาสาเหตุ, แก้โค้ด, รันเทสต์, commit ใน worktree ของบอท
3. ทำแล้วรออนุมัติ: deploy, push, merge เข้า main, gcloud ที่เปลี่ยนของ, แตะข้อมูล prod, ใช้เงิน
   ระบบจะเด้งปุ่มให้เจ้าของกดเอง คุณแค่เรียกคำสั่งตามปกติ ถ้าโดนปฏิเสธ ห้ามหาทางอ้อม
4. ส่งคืนเจ้าของ (เรียก tool hand_back) เมื่อเกินกำลัง เช่น
   - ต้อง login (gcloud auth login, firebase login) หรือกดอะไรใน dashboard ของ vendor
   - ต้องตัดสินใจเรื่องธุรกิจ/ขอบเขตกับลูกค้า หรือต้องส่งข้อความหาลูกค้า/หัวหน้า (ร่างให้ได้ แต่ไม่ส่งเอง)
   - ของอยู่นอกเครื่องนี้ (เครื่อง Windows, เน็ตเวิร์กลูกค้า) หรือข้อมูลที่ต้องมีไม่อยู่ในเครื่อง
   - ลองแก้แล้ว 3 รอบยังไม่ผ่าน: หยุด อย่าวนเผาเงิน
   hand_back ต้องบอก: ลองอะไรไปแล้ว ติดตรงไหน เจ้าของต้องทำอะไรต่อแบบทำตามได้ทันที

## กติกาการแก้โค้ด
- ก่อนแก้โค้ดใน repo ไหน ต้องเรียก use_repo ก่อนเสมอ แล้วทำงานใน path ที่ได้กลับมาเท่านั้น
  ห้ามแก้ไฟล์ใน working copy ของเจ้าของ (${config.reposRoot}/<repo>) โดยตรง เพราะเจ้าของอาจมี session อื่นเปิดอยู่
- commit ได้ใน worktree (author ใช้ค่า git config ที่มีอยู่) แต่ push/deploy ต้องผ่านปุ่มอนุมัติ
- พิสูจน์ด้วยหลักฐานจริง (รันเทสต์ อ่าน log เปิดหน้าเว็บ) ห้ามเดา ห้ามบอกว่าเสร็จถ้ายังไม่ได้ตรวจ

## รายงานตอนจบเทิร์น (ข้อความสุดท้าย)
- ทำอะไรไป / ตรวจอะไรแล้วเห็นอะไร / ค้างอะไร หรือเจ้าของต้องทำอะไร
- ถ้าแก้โค้ด: บอก branch, path ของ worktree และไฟล์ที่แก้
- ไม่เกิน ~15 บรรทัด ถ้ายาวกว่านั้นให้เขียนเป็นไฟล์ใน worktree แล้วบอก path
`;

function summarizeToolUse(name, input = {}) {
  const short = (s, n = 140) => (s && s.length > n ? `${s.slice(0, n)}…` : s || '');
  switch (name) {
    case 'Bash': return `รัน \`${short(input.command, 160)}\``;
    case 'Read': return `อ่าน ${input.file_path}`;
    case 'Edit': case 'MultiEdit': return `แก้ ${input.file_path}`;
    case 'Write': return `เขียน ${input.file_path}`;
    case 'Grep': return `ค้น "${short(input.pattern, 60)}"`;
    case 'Glob': return `หาไฟล์ ${input.pattern}`;
    case 'WebFetch': return `เปิด ${input.url}`;
    case 'WebSearch': return `ค้นเว็บ "${short(input.query, 80)}"`;
    case 'Task': case 'Agent': return `ส่งงานย่อย: ${short(input.description, 80)}`;
    default: return name.replace(/^mcp__/, '');
  }
}

/**
 * @param {object} p
 * @param {string} p.taskId            Discord thread id
 * @param {string} p.prompt
 * @param {object} p.state             current task record from the store (may be empty)
 * @param {(patch:object)=>void} p.save persist task state
 * @param {object} p.ui                { status(text), text(text), approve({tool,input,reason,signal})→Promise<boolean>, handBack(info) }
 * @param {AbortController} p.abort
 */
export async function runTurn({ taskId, prompt, state, save, ui, abort }) {
  const ctx = {
    worktrees: [...(state.worktrees || []).map((w) => w.dir)],
    cwd: state.cwd || config.reposRoot,
    reposRoot: config.reposRoot,
    scratch: [config.dataDir],
  };
  let handedBack = null;

  const assistantTools = createSdkMcpServer({
    name: 'assistant',
    version: '1.0.0',
    tools: [
      tool('list_repos', 'List the git repos the bot can work in.', {}, async () => ({
        content: [{ type: 'text', text: listRepos().join('\n') }],
      })),
      tool(
        'use_repo',
        'Create (or reuse) this task\'s own git worktree + branch for a repo and return its path. Call before editing any code in that repo; then work only inside the returned path.',
        { repo: z.string().describe('folder name under the repos root, e.g. tttc-ocr-lcmd'), base_ref: z.string().optional().describe('git ref to branch from; default origin/HEAD') },
        async ({ repo, base_ref }) => {
          try {
            const existing = (state.worktrees || []).find((w) => w.repo === repo);
            const wt = existing || { repo, ...(await createWorktree(repo, taskId, base_ref)) };
            if (!existing) {
              state.worktrees = [...(state.worktrees || []), { repo, dir: wt.dir, branch: wt.branch, base: wt.base }];
              save({ worktrees: state.worktrees });
              ctx.worktrees.push(wt.dir);
              ui.status(`เปิด worktree ${repo} → ${wt.branch}`);
            }
            const env = wt.envCopied?.length ? `\ncopied env files: ${wt.envCopied.join(', ')}` : '';
            return { content: [{ type: 'text', text: `worktree: ${wt.dir}\nbranch: ${wt.branch}\nbase: ${wt.base || ''}${env}\nWork only inside this path (use absolute paths or cd into it).` }] };
          } catch (err) {
            return { content: [{ type: 'text', text: `use_repo failed: ${err.message}` }], isError: true };
          }
        },
      ),
      tool(
        'hand_back',
        'Stop and hand the task back to the owner because it needs a human or is beyond what you can do. After calling this, end your turn with one short line.',
        {
          blocker: z.string().describe('what blocks you, one or two sentences'),
          tried: z.string().describe('what you already tried / checked'),
          owner_action: z.string().describe('exact steps the owner should do next (commands, where to click)'),
          prepared: z.string().optional().describe('anything you already prepared: branch, draft text, file path'),
        },
        async (args) => {
          handedBack = args;
          await ui.handBack(args);
          return { content: [{ type: 'text', text: 'Handed back to the owner. End your turn now with one short line.' }] };
        },
      ),
    ],
  });

  const canUseTool = async (toolName, input, { signal }) => {
    const { decision, reason } = classify(toolName, input, ctx);
    if (decision === 'allow') return { behavior: 'allow', updatedInput: input };
    if (decision === 'deny') return { behavior: 'deny', message: reason };
    ui.status(`รออนุมัติ: ${reason}`);
    const ok = await ui.approve({ tool: toolName, input, reason, signal });
    return ok
      ? { behavior: 'allow', updatedInput: input }
      : { behavior: 'deny', message: `เจ้าของไม่อนุมัติ (${reason}). อย่าหาทางอ้อม ถ้างานต่อไม่ได้ให้สรุปแล้วจบเทิร์น` };
  };

  const options = {
    cwd: ctx.cwd,
    resume: state.sessionId || undefined,
    model: config.model,
    maxTurns: config.maxTurnsPerRun,
    maxBudgetUsd: config.maxBudgetUsdPerRun,
    permissionMode: 'default',
    canUseTool,
    abortController: abort,
    settingSources: ['user', 'project', 'local'], // your CLAUDE.md, skills, hooks (guard-dangerous etc.) apply too
    systemPrompt: { type: 'preset', preset: 'claude_code', append: SYSTEM_APPEND },
    mcpServers: { assistant: assistantTools },
    additionalDirectories: [config.worktreesRoot, config.dataDir],
    stderr: (line) => { if (/error/i.test(line)) console.error(`[claude:${taskId}] ${line.trim()}`); },
  };

  let sessionId = state.sessionId;
  let final = null;
  for await (const msg of query({ prompt, options })) {
    if (msg.type === 'system' && msg.subtype === 'init') {
      sessionId = msg.session_id;
      save({ sessionId, cwd: ctx.cwd });
    } else if (msg.type === 'assistant' && msg.parent_tool_use_id == null) {
      if (msg.error) ui.status(`API error: ${msg.error}`);
      for (const block of msg.message?.content || []) {
        if (block.type === 'tool_use') ui.status(summarizeToolUse(block.name, block.input));
        else if (block.type === 'text' && block.text.trim()) ui.text(block.text.trim());
      }
    } else if (msg.type === 'result') {
      final = msg;
    }
  }

  return {
    sessionId,
    handedBack,
    subtype: final?.subtype || 'aborted',
    isError: final ? final.is_error || final.subtype !== 'success' : true,
    result: final?.result || '',
    costUsd: final?.total_cost_usd || 0,
    turns: final?.num_turns || 0,
    durationMs: final?.duration_ms || 0,
    denials: final?.permission_denials?.length || 0,
  };
}
