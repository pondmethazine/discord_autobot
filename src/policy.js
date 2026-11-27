// Decides, per tool call, whether the agent may go ahead on its own or must wait for
// the owner's button press. Rule of thumb: anything local and reversible inside the
// task's worktree runs freely; anything that leaves the machine, touches prod,
// spends money, or touches your own working copies waits for approval.
import path from 'node:path';
import os from 'node:os';
import { isInside } from './repos.js';

const READ_ONLY_TOOLS = new Set([
  'Read', 'Glob', 'Grep', 'LS', 'WebFetch', 'WebSearch', 'TodoWrite', 'TodoRead',
  'Task', 'Agent', 'Skill', 'ToolSearch', 'BashOutput', 'KillShell', 'KillBash', 'TaskOutput',
  'TaskStop', 'ListMcpResourcesTool', 'ReadMcpResourceTool', 'ExitPlanMode', 'EnterPlanMode',
  'Monitor',
]);
const EDIT_TOOLS = new Set(['Write', 'Edit', 'MultiEdit', 'NotebookEdit']);
const TRUSTED_MCP_PREFIXES = ['mcp__assistant__', 'mcp__context7__', 'mcp__chrome-devtools__'];

// Each rule: [regex, human reason]. Matched against the whole command string,
// so a risky step hidden in `a && b` or a pipeline still triggers approval.
const BASH_RULES = [
  [/\bgit\s+(?:-C\s+\S+\s+)?push\b/, 'git push (ส่งโค้ดขึ้น remote)'],
  [/\bgit\s+ship\b/, 'git ship (release)'],
  [/\bgit\s+(?:-C\s+\S+\s+)?reset\s+--hard\b/, 'git reset --hard'],
  [/\bgit\s+(?:-C\s+\S+\s+)?clean\s+-\w*[fdx]/, 'git clean'],
  [/\bgit\s+(?:-C\s+\S+\s+)?branch\s+-D\b/, 'ลบ branch'],
  [/\bgit\s+(?:-C\s+\S+\s+)?worktree\s+remove\b/, 'ลบ worktree'],
  [/(?<![\w-])deploy(?![\w\/-])/i, 'deploy'],
  [/\b(?:railway\s+(?:up|redeploy|down)|netlify\s+(?:deploy|sites:delete)|vercel(?:\s|$))/, 'deploy ผ่าน PaaS'],
  [/\bfirebase\s+\S*(?:deploy|delete|:set|:unset|:disable|use\s+--add)/, 'firebase เปลี่ยนของบน cloud'],
  [/\b(?:gsutil|gcloud\s+storage)\s+(?:-m\s+)?(?:rm|mv|rsync|cp\s+\S+\s+gs:|cp\s+.*\sgs:)/, 'เขียน/ลบไฟล์ใน GCS'],
  [/\bbq\s+(?:query|rm|mk|load|insert|update|cp)\b(?![\s\S]*--dry_run)/, 'BigQuery (มีค่าใช้จ่าย/เขียนข้อมูล)'],
  [/\b(?:psql|mysql|sqlcmd|mongosh|mongo|redis-cli)\b/, 'ต่อฐานข้อมูลตรง'],
  [/\bcurl\b[^|;&]*(?:-X\s*(?:POST|PUT|PATCH|DELETE)|--request\s+(?:POST|PUT|PATCH|DELETE)|\s-d\s|--data|-F\s|--form)/i, 'ยิง HTTP ที่เขียนข้อมูล'],
  [/\b(?:wget)\b[^|;&]*--post/i, 'ยิง HTTP ที่เขียนข้อมูล'],
  [/\bnpm\s+(?:publish|unpublish|install\s+-g|i\s+-g)\b/, 'npm publish / ติดตั้ง global'],
  [/\b(?:docker|podman)\s+(?:push|rm|rmi|system\s+prune)\b/, 'docker push/ลบ'],
  [/\b(?:sudo|launchctl|crontab|shutdown|reboot)\b/, 'คำสั่งระดับเครื่อง'],
  [/\b(?:kill|pkill|killall)\b/, 'kill process'],
  [/\bPIPE_SPEND_OK\b/, 'เปิดสวิตช์ใช้เงิน (PIPE_SPEND_OK)'],
  [/>\s*\S*\.env\b/, 'เขียนทับไฟล์ .env'],
  [/\bchmod\s+-R\b|\bchown\b/, 'เปลี่ยนสิทธิ์ไฟล์'],
];

const GCLOUD_READ = /\b(?:list|describe|read|get-value|get-iam-policy|tail|logs?\s+read|config\s+(?:list|get)|auth\s+list|info|version|help)\b/;

const GIT_WRITE = /\bgit\s+(?:-C\s+\S+\s+)?(?:commit|checkout|switch|stash|pull|merge|rebase|cherry-pick|am|apply|restore|rm|mv|add|tag)\b/;

function checkRmR(cmd, allowedRoots) {
  const m = cmd.match(/\brm\s+(-\w*r\w*|-\w*R\w*|--recursive)\b([^;&|]*)/);
  if (!m) return null;
  const targets = m[2].trim().split(/\s+/).filter((a) => a && !a.startsWith('-'));
  const ok = targets.length > 0 && targets.every((t) => {
    if (/[*?$`~]/.test(t) && !t.startsWith('/tmp/')) return false;
    const abs = path.resolve(t);
    return allowedRoots.some((r) => isInside(abs, r)) && !allowedRoots.includes(abs);
  });
  return ok ? null : 'ลบโฟลเดอร์ (rm -r) นอก worktree';
}

/**
 * @param {string} toolName
 * @param {object} input
 * @param {{ worktrees: string[], cwd: string, reposRoot: string, scratch: string[] }} ctx
 * @returns {{ decision: 'allow'|'ask'|'deny', reason: string }}
 */
export function classify(toolName, input = {}, ctx) {
  const writableRoots = [...ctx.worktrees, ...ctx.scratch, os.tmpdir(), '/tmp', '/private/tmp'];

  if (toolName === 'AskUserQuestion') {
    return { decision: 'deny', reason: 'ห้ามใช้ AskUserQuestion ใน Discord: ถ้าต้องการคำตอบ ให้จบเทิร์นด้วยคำถามเป็นข้อความ หรือเรียก hand_back' };
  }
  if (READ_ONLY_TOOLS.has(toolName)) return { decision: 'allow', reason: 'read-only' };
  if (TRUSTED_MCP_PREFIXES.some((p) => toolName.startsWith(p))) return { decision: 'allow', reason: 'trusted tool' };

  if (EDIT_TOOLS.has(toolName)) {
    const p = input.file_path || input.notebook_path || input.path;
    if (p && writableRoots.some((r) => isInside(p, r))) return { decision: 'allow', reason: 'แก้ไฟล์ใน worktree' };
    return { decision: 'ask', reason: `แก้ไฟล์นอก worktree: ${p || '(ไม่ระบุ path)'}` };
  }

  if (toolName === 'Bash') {
    const cmd = String(input.command || '');
    for (const [re, reason] of BASH_RULES) if (re.test(cmd)) return { decision: 'ask', reason };
    if (/\bgcloud\b/.test(cmd) && !GCLOUD_READ.test(cmd)) return { decision: 'ask', reason: 'gcloud ที่อาจเปลี่ยนของบน cloud' };
    if (/\bgcloud\b[\s\S]*\b(?:create|delete|update|deploy|set|add|remove|patch|start|stop|execute|submit|enable|disable|import|export|rollback|cancel|resize|replace|apply)\b/.test(cmd)) {
      return { decision: 'ask', reason: 'gcloud ที่เปลี่ยนของบน cloud' };
    }
    const rm = checkRmR(cmd, writableRoots);
    if (rm) return { decision: 'ask', reason: rm };
    if (GIT_WRITE.test(cmd)) {
      const inWorktree = ctx.worktrees.some((w) => isInside(ctx.cwd, w));
      const mentionsOtherRepo = new RegExp(`${escapeRe(ctx.reposRoot)}/(?!_bot_worktrees)[\\w.-]+`).test(cmd);
      if (!inWorktree || mentionsOtherRepo) {
        return { decision: 'ask', reason: 'คำสั่ง git ที่เขียนลง working copy ของมึงเอง (ไม่ใช่ worktree ของบอท)' };
      }
    }
    return { decision: 'allow', reason: 'local command' };
  }

  if (toolName.startsWith('mcp__')) return { decision: 'ask', reason: `MCP tool ภายนอก: ${toolName}` };
  return { decision: 'ask', reason: `tool ที่ไม่รู้จัก: ${toolName}` };
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
