// Repo discovery and per-task git worktrees.
// The bot never edits your working copies: every task that changes code gets its own
// worktree + branch under WORKTREES_ROOT, so it can't collide with sessions you have open.
import fs from 'node:fs';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { config } from './config.js';

const run = promisify(execFile);

async function git(cwd, ...args) {
  const { stdout } = await run('git', args, { cwd, maxBuffer: 16 * 1024 * 1024 });
  return stdout.trim();
}

const SKIP = /^(_|\.|node_modules$)/;

export function listRepos() {
  if (!fs.existsSync(config.reposRoot)) return [];
  return fs
    .readdirSync(config.reposRoot, { withFileTypes: true })
    .filter((d) => d.isDirectory() && !SKIP.test(d.name))
    .filter((d) => fs.existsSync(path.join(config.reposRoot, d.name, '.git')))
    .map((d) => d.name)
    .sort();
}

export function repoPath(name) {
  if (!/^[\w.-]+$/.test(name)) throw new Error(`invalid repo name: ${name}`);
  const p = path.join(config.reposRoot, name);
  if (!fs.existsSync(path.join(p, '.git'))) throw new Error(`not a git repo under ${config.reposRoot}: ${name}`);
  return p;
}

// "tttc-ocr-lcmd: fix DEF-032" or "[tttc-ocr-lcmd] fix ..." → { repo, text }
export function parseRepoPrefix(text, repos = listRepos()) {
  const m = text.match(/^\s*(?:\[([\w.-]+)\]|([\w.-]+)\s*:)\s*([\s\S]*)$/);
  const name = m && (m[1] || m[2]);
  if (name && repos.includes(name)) return { repo: name, text: m[3].trim() };
  return { repo: null, text: text.trim() };
}

async function defaultBranch(repo) {
  try {
    const ref = await git(repo, 'symbolic-ref', '--short', 'refs/remotes/origin/HEAD');
    return ref; // e.g. origin/main
  } catch {
    for (const b of ['origin/main', 'origin/master']) {
      try {
        await git(repo, 'rev-parse', '--verify', '--quiet', b);
        return b;
      } catch {}
    }
    return 'HEAD';
  }
}

// Untracked env files are not in the worktree; copy them so tests and scripts can run.
function copyEnvFiles(from, to) {
  const copied = [];
  for (const f of fs.readdirSync(from)) {
    if (/^\.env(\..+)?$/.test(f) && !/\.(example|sample|template)$/.test(f)) {
      const dst = path.join(to, f);
      if (!fs.existsSync(dst)) {
        fs.copyFileSync(path.join(from, f), dst);
        copied.push(f);
      }
    }
  }
  return copied;
}

export async function createWorktree(repoName, taskId, baseRef) {
  const repo = repoPath(repoName);
  const short = String(taskId).slice(-6);
  const branch = `bot/${short}`;
  const dir = path.join(config.worktreesRoot, `${repoName}-${short}`);
  if (fs.existsSync(dir)) return { dir, branch, base: baseRef || null, reused: true, envCopied: [] };

  fs.mkdirSync(config.worktreesRoot, { recursive: true });
  try {
    await git(repo, 'fetch', '--quiet', 'origin');
  } catch (err) {
    console.warn(`fetch failed for ${repoName} (working from local refs): ${err.message.split('\n')[0]}`);
  }
  const base = baseRef || (await defaultBranch(repo));
  await git(repo, 'worktree', 'add', '-b', branch, dir, base);
  const envCopied = copyEnvFiles(repo, dir);
  return { dir, branch, base, reused: false, envCopied };
}

export function isInside(child, parent) {
  const rel = path.relative(path.resolve(parent), path.resolve(child));
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}
