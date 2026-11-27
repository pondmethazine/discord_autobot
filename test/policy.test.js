import { test } from 'node:test';
import assert from 'node:assert/strict';
import { classify } from '../src/policy.js';
import { parseRepoPrefix } from '../src/repos.js';
import { chunk } from '../src/text.js';

const WT = '/Users/user/rifle/_bot_worktrees/back-office-iq-123456';
const ctx = { worktrees: [WT], cwd: WT, reposRoot: '/Users/user/rifle', scratch: ['/Users/user/rifle/discord-auto-bot/data'] };
const bash = (command, c = ctx) => classify('Bash', { command }, c).decision;

test('read-only and local work runs without asking', () => {
  assert.equal(classify('Read', { file_path: '/etc/hosts' }, ctx).decision, 'allow');
  assert.equal(classify('Grep', { pattern: 'x' }, ctx).decision, 'allow');
  for (const c of ['npm test', 'git status', 'git log --oneline -5', 'git diff', 'ls deploy/', 'cat deploy/notify-discord.sh',
    'gcloud logging read "resource.type=cloud_run_revision" --limit 20', 'gcloud run revisions list --service api',
    'git add -A && git commit -m "fix"', `rm -rf ${WT}/dist`, 'node scripts/check.js']) {
    assert.equal(bash(c), 'allow', c);
  }
});

test('edits inside the worktree are allowed, outside need approval', () => {
  assert.equal(classify('Edit', { file_path: `${WT}/src/a.ts` }, ctx).decision, 'allow');
  assert.equal(classify('Write', { file_path: '/tmp/x.txt' }, ctx).decision, 'allow');
  assert.equal(classify('Edit', { file_path: '/Users/user/rifle/back-office-iq/src/a.ts' }, ctx).decision, 'ask');
  assert.equal(classify('Write', { file_path: '/Users/user/.claude/settings.json' }, ctx).decision, 'ask');
});

test('anything leaving the machine, touching prod or spending money asks', () => {
  for (const c of [
    'git push origin bot/123456', 'git -C /x push', 'npm run build && git push', 'git ship',
    'firebase deploy --only hosting', 'gcloud run deploy api --source .', './deploy.sh', 'bash scripts/deploy.sh prod',
    'npm run deploy', 'gcloud secrets versions add KEY --data-file=-', 'gcloud run services update api --memory 1Gi',
    'gcloud config set project x', 'bq query "select 1"', 'gsutil rm gs://b/x', 'gsutil cp a.pdf gs://bucket/a.pdf',
    'curl -X POST https://api.example.com/x', 'curl -d "a=1" https://x', 'psql "$DATABASE_URL"',
    'git reset --hard origin/main', 'rm -rf /Users/user/rifle/back-office-iq', 'rm -rf ~/x', 'sudo ls',
    'PIPE_SPEND_OK=yes node run.js', 'echo X=1 > .env', 'railway up', 'npm publish',
  ]) {
    assert.equal(bash(c), 'ask', c);
  }
});

test('git writes in the owner\'s own working copy ask', () => {
  const home = { ...ctx, cwd: '/Users/user/rifle' };
  assert.equal(bash('git -C back-office-iq commit -m x', home), 'ask');
  assert.equal(bash('cd /Users/user/rifle/back-office-iq && git stash', ctx), 'ask');
  assert.equal(bash('git status', home), 'allow');
});

test('AskUserQuestion is denied with guidance; unknown MCP tools ask', () => {
  assert.equal(classify('AskUserQuestion', {}, ctx).decision, 'deny');
  assert.equal(classify('mcp__assistant__use_repo', {}, ctx).decision, 'allow');
  assert.equal(classify('mcp__plugin_supabase_supabase__execute_sql', {}, ctx).decision, 'ask');
});

test('repo prefix parsing', () => {
  const repos = ['tttc-ocr-lcmd', 'back-office-iq'];
  assert.deepEqual(parseRepoPrefix('tttc-ocr-lcmd: แก้ DEF-032', repos), { repo: 'tttc-ocr-lcmd', text: 'แก้ DEF-032' });
  assert.deepEqual(parseRepoPrefix('[back-office-iq] deploy', repos), { repo: 'back-office-iq', text: 'deploy' });
  assert.deepEqual(parseRepoPrefix('note: ไม่ใช่ repo', repos), { repo: null, text: 'note: ไม่ใช่ repo' });
});

test('chunk keeps messages under the Discord limit and fences balanced', () => {
  const long = '```\n' + 'x'.repeat(100).concat('\n').repeat(60) + '```';
  const parts = chunk(long, 500);
  assert.ok(parts.length > 1);
  for (const p of parts) {
    assert.ok(p.length <= 510, `len ${p.length}`);
    assert.equal((p.match(/```/g) || []).length % 2, 0, p.slice(0, 40));
  }
});
