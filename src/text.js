// Discord caps messages at 2000 chars. Split on line boundaries, and keep ``` fences balanced.
export function chunk(text, max = 1900) {
  const out = [];
  let cur = '';
  let inFence = false;
  for (const line of String(text).split('\n')) {
    const pieces = line.length > max ? line.match(new RegExp(`.{1,${max}}`, 'gs')) : [line];
    for (const piece of pieces) {
      if (cur.length + piece.length + 1 > max) {
        out.push(inFence ? `${cur}\n\`\`\`` : cur);
        cur = inFence ? '```\n' : '';
      }
      cur += (cur && !cur.endsWith('\n') ? '\n' : '') + piece;
      if (/^```/.test(piece.trim())) inFence = !inFence;
    }
  }
  if (cur.trim()) out.push(cur);
  return out.length ? out : [''];
}

export function codeBlock(s, max = 1500) {
  const body = String(s).replace(/```/g, "'''");
  return '```\n' + (body.length > max ? `${body.slice(0, max)}\n…(ตัดเหลือ ${max} ตัวอักษร)` : body) + '\n```';
}

export function fmtDuration(ms) {
  const s = Math.round((ms || 0) / 1000);
  if (s < 60) return `${s}s`;
  const m = Math.floor(s / 60);
  if (m < 60) return `${m}m${String(s % 60).padStart(2, '0')}s`;
  return `${Math.floor(m / 60)}h${String(m % 60).padStart(2, '0')}m`;
}
