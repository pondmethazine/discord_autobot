// Small JSON-file store for task state (one record per Discord thread).
// Writes are atomic (tmp file + rename) so a crash mid-write never corrupts it.
import fs from 'node:fs';
import path from 'node:path';

export class Store {
  constructor(file) {
    this.file = file;
    fs.mkdirSync(path.dirname(file), { recursive: true });
    this.data = { tasks: {} };
    if (fs.existsSync(file)) {
      try {
        this.data = JSON.parse(fs.readFileSync(file, 'utf8'));
        this.data.tasks ||= {};
      } catch (err) {
        const bad = `${file}.corrupt-${Date.now()}`;
        fs.renameSync(file, bad);
        console.error(`state file unreadable, moved to ${bad}: ${err.message}`);
      }
    }
  }

  get(threadId) {
    return this.data.tasks[threadId];
  }

  all() {
    return Object.entries(this.data.tasks).map(([threadId, t]) => ({ threadId, ...t }));
  }

  update(threadId, patch) {
    const cur = this.data.tasks[threadId] || { createdAt: new Date().toISOString(), costUsd: 0 };
    this.data.tasks[threadId] = { ...cur, ...patch, updatedAt: new Date().toISOString() };
    this.save();
    return this.data.tasks[threadId];
  }

  save() {
    const tmp = `${this.file}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(this.data, null, 2));
    fs.renameSync(tmp, this.file);
  }
}
