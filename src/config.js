import 'dotenv/config';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// An empty ANTHROPIC_API_KEY= line in .env would shadow the Claude Code login; treat it as unset.
if (process.env.ANTHROPIC_API_KEY === '') delete process.env.ANTHROPIC_API_KEY;

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function num(name, fallback) {
  const v = process.env[name];
  if (v === undefined || v === '') return fallback;
  const n = Number(v);
  if (!Number.isFinite(n)) throw new Error(`${name} must be a number, got "${v}"`);
  return n;
}

export const config = {
  root: ROOT,

  // TODO[human]: create a NEW bot in https://discord.com/developers/applications
  //   → Bot → Reset Token → copy into .env as DISCORD_TOKEN=...
  //   → Bot → enable "Message Content Intent"
  //   → OAuth2 → URL Generator: scopes "bot", permissions: View Channels, Send Messages,
  //     Create Public Threads, Send Messages in Threads, Read Message History,
  //     Attach Files, Add Reactions, Manage Threads → open the URL and invite it to your private server.
  discordToken: process.env.DISCORD_TOKEN || '',

  // TODO[human]: Discord → Settings → Advanced → Developer Mode on, then
  //   right-click yourself → Copy User ID → OWNER_ID=...
  //   right-click the private channel the bot should listen in → Copy Channel ID → HOME_CHANNEL_ID=...
  ownerId: process.env.OWNER_ID || '',
  homeChannelId: process.env.HOME_CHANNEL_ID || '',

  // Where your repos live. The bot can work in any git repo directly under this folder.
  reposRoot: process.env.REPOS_ROOT || path.join(os.homedir(), 'rifle'),
  // Where the bot creates its own git worktrees (never inside your working copies).
  worktreesRoot: process.env.WORKTREES_ROOT || path.join(os.homedir(), 'rifle', '_bot_worktrees'),

  dataDir: process.env.DATA_DIR || path.join(ROOT, 'data'),

  model: process.env.MODEL || undefined, // undefined = Claude Code default
  maxBudgetUsdPerRun: num('MAX_BUDGET_USD_PER_RUN', 5),
  maxTurnsPerRun: num('MAX_TURNS_PER_RUN', 150),
  maxConcurrentRuns: num('MAX_CONCURRENT_RUNS', 3),
  // How long an approval button waits before it counts as "deny".
  approvalTimeoutMin: num('APPROVAL_TIMEOUT_MIN', 360),
};

export function assertDiscordConfig() {
  const missing = ['discordToken', 'ownerId', 'homeChannelId'].filter((k) => !config[k]);
  if (missing.length) {
    const env = { discordToken: 'DISCORD_TOKEN', ownerId: 'OWNER_ID', homeChannelId: 'HOME_CHANNEL_ID' };
    throw new Error(`Missing in .env: ${missing.map((k) => env[k]).join(', ')} (see .env.example)`);
  }
}
