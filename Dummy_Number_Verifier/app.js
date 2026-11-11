// >>> Paste your deployed Google Apps Script Web App URL here <<<
const API_URL = 'https://script.google.com/macros/s/AKfycbzga__srV9aDAIq59mf9oBdY2mwz5dCzFYea4O5-6pBnI-zvJd0xNrUIsIvAm68EQlXow/exec';
const $ = s => document.querySelector(s);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
document.documentElement.dataset.theme = localStorage.getItem('theme') || (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light');
function toggleTheme() {
  const t = document.documentElement.dataset.theme == 'dark' ? 'light' : 'dark';
  document.documentElement.dataset.theme = t; localStorage.setItem('theme', t);
  document.querySelectorAll('.tg').forEach(b => b.textContent = t == 'dark' ? '☀️' : '🌙');
}
function _ui() {
  if (!$('#busy')) document.body.insertAdjacentHTML('beforeend', '<div class="toast" id="toast"></div><div class="busy" id="busy"><i></i></div>');
  document.querySelectorAll('.tg').forEach(b => { b.textContent = document.documentElement.dataset.theme == 'dark' ? '☀️' : '🌙'; b.onclick = toggleTheme; });
}
// app.js is loaded at the end of <body>, so run now (page scripts call the API immediately)
document.body ? _ui() : addEventListener('DOMContentLoaded', _ui);
let _tt; function toast(m, t = '') { const e = $('#toast'); e.textContent = m; e.className = 'toast on ' + t; clearTimeout(_tt); _tt = setTimeout(() => e.className = 'toast', 3200); }
function busy(v) { $('#busy').classList.toggle('on', v); }
async function api(action, p = {}, quiet = false) {
  if (!quiet) busy(true);
  try {
    const r = await fetch(API_URL, { method: 'POST', body: JSON.stringify({ action, token: sessionStorage.getItem('tk'), ...p }) });
    const j = await r.json();
    if (!j.ok) { if (j.error == 'SESSION') { logout(); } throw new Error(j.error); }
    return j;
  } catch (e) { toast(e.message == 'Failed to fetch' ? 'Network / API URL error' : e.message, 'e'); throw e; }
  finally { if (!quiet) busy(false); }
}
function guard(role) {
  const u = JSON.parse(sessionStorage.getItem('user') || 'null');
  if (!u || u.role != role) { sessionStorage.clear(); location.replace('index.html'); return null; } // clear first: prevents index<->page redirect loop
  return u;
}
let _lo = false;
function logout() { if (_lo) return; _lo = true; sessionStorage.clear(); location.replace('index.html'); }
function xl(name, head, rows) {
  const ws = XLSX.utils.aoa_to_sheet([head, ...rows]); ws['!cols'] = head.map(h => ({ wch: Math.max(14, h.length + 2) }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report'); XLSX.writeFile(wb, name + '.xlsx');
}