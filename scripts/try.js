// Run one task turn from the terminal, without Discord. Useful for testing the agent side.
//   npm run try -- "tttc-ocr-lcmd: สรุป 5 commit ล่าสุดให้หน่อย"
//   TRY_APPROVE=1 npm run try -- "..."     # auto-approve risky steps (careful)
//   TRY_TASK=<id> npm run try -- "..."     # continue an earlier terminal task
import path from 'node:path';
import { config } from '../src/config.js';
import { Store } from '../src/store.js';
import { runTurn } from '../src/agent.js';
import { createWorktree, parseRepoPrefix } from '../src/repos.js';

const raw = process.argv.slice(2).join(' ');
if (!raw) {
  console.error('usage: npm run try -- "<repo>: <task>"');
  process.exit(1);
}
const store = new Store(path.join(config.dataDir, 'try-state.json'));
const taskId = process.env.TRY_TASK || `t${Date.now()}`;
const state = { ...(store.get(taskId) || {}) };
const save = (patch) => Object.assign(state, store.update(taskId, patch));

const { repo, text } = parseRepoPrefix(raw);
if (repo && !state.worktrees?.length) {
  const wt = await createWorktree(repo, taskId);
  save({ repo, worktrees: [{ repo, dir: wt.dir, branch: wt.branch, base: wt.base }], cwd: wt.dir });
  console.log(`[worktree] ${wt.dir} (${wt.branch} from ${wt.base})`);
}

const ui = {
  status: (t) => console.log(`[..] ${t}`),
  text: (t) => console.log(`[text] ${t}`),
  approve: async ({ reason, input }) => {
    const ok = process.env.TRY_APPROVE === '1';
    console.log(`[approval ${ok ? 'AUTO-YES' : 'AUTO-NO'}] ${reason}: ${JSON.stringify(input).slice(0, 200)}`);
    return ok;
  },
  handBack: async (info) => console.log('[hand_back]', info),
};

const res = await runTurn({ taskId, prompt: text, state, save, ui, abort: new AbortController() });
save({ costUsd: (state.costUsd || 0) + res.costUsd, status: res.handedBack ? 'handed_back' : res.isError ? 'error' : 'done' });
console.log('\n===== RESULT =====\n' + res.result);
console.log(`\n[${res.subtype}] task=${taskId} turns=${res.turns} cost=$${res.costUsd.toFixed(4)} denials=${res.denials} ${Math.round(res.durationMs / 1000)}s`);
