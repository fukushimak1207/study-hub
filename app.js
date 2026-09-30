// 学びハブ(P1: 見る + 状態の変更)
'use strict';

const STATUSES = ['新着', '検討中', '申込済', '参加済', '見送り'];
const K_CONF = 'hub.conf';
const K_DATA = 'hub.data';
const DEMO = new URLSearchParams(location.search).has('demo');

let data = { seminars: [], papers: [], fetched_at: '' };
let seminarFilter = '予定';

// ---------- 保存(端末内) ----------
function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 保存できなくても表示は続ける */ } }
function getConf() { return lsGet(K_CONF) || { url: '', token: '' }; }

// ---------- 表示の小道具 ----------
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function safeUrl(u) { return /^https?:\/\//i.test(u || '') ? u : ''; }
function todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function isDate(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s || ''); }
function daysUntil(s) {
  if (!isDate(s)) return null;
  const [y, m, d] = s.split('-').map(Number);
  const t = new Date(); t.setHours(0, 0, 0, 0);
  return Math.round((new Date(y, m - 1, d) - t) / 86400000);
}
function fmtDate(s) {
  if (!isDate(s)) return s || '';
  const [y, m, d] = s.split('-').map(Number);
  const w = '日月火水木金土'[new Date(y, m - 1, d).getDay()];
  return m + '/' + d + '(' + w + ')';
}
function whenText(s) {
  const main = s.date ? fmtDate(s.date) : '';
  if (main && s.date_text && s.date_text !== s.date) return main + ' ' + s.date_text;
  return main || s.date_text || '日付未定';
}

function banner(msg) {
  const b = document.getElementById('banner');
  b.textContent = msg || '';
  b.hidden = !msg;
}

// ---------- 勉強会のカード ----------
function deadlineChip(s) {
  const n = daysUntil(s.deadline);
  if (n === null || !['新着', '検討中'].includes(s.status)) return '';
  if (n < 0) return '<span class="chip">締切済み</span>';
  const cls = n <= 3 ? 'chip hot' : 'chip warn';
  return '<span class="' + cls + '">締切 ' + fmtDate(s.deadline) + (n === 0 ? ' 今日' : ' あと' + n + '日') + '</span>';
}
function statusChip(st) {
  const cls = st === '新着' ? 'chip hot' : st === '申込済' ? 'chip ok' : 'chip';
  return '<span class="' + cls + '">' + esc(st) + '</span>';
}

function seminarCard(s) {
  const meta = [s.organizer, s.place, s.format, s.fee ? '参加費 ' + s.fee : '']
    .filter(Boolean).map(esc).map(x => '<span>' + x + '</span>').join('');
  const url = safeUrl(s.url);
  const opts = STATUSES.map(x => '<option' + (x === s.status ? ' selected' : '') + '>' + x + '</option>').join('');
  return '<article class="card">' +
    '<div class="chips">' + statusChip(s.status) + deadlineChip(s) +
      (s.role && s.role !== '参加' ? '<span class="chip">' + esc(s.role) + '</span>' : '') + '</div>' +
    '<div class="item-title">' + esc(s.name) + '</div>' +
    '<div class="meta"><span><b>' + esc(whenText(s)) + '</b></span>' + meta + '</div>' +
    (s.theme ? '<p class="small">' + esc(s.theme) + '</p>' : '') +
    (s.note ? '<p class="small">メモ: ' + esc(s.note) + '</p>' : '') +
    '<div class="row">' +
      (url ? '<a href="' + esc(url) + '" target="_blank" rel="noopener">案内を開く</a>' : '<span class="small">案内リンクなし</span>') +
      '<select aria-label="状態を変える" data-id="' + esc(s.id) + '">' + opts + '</select>' +
    '</div>' +
  '</article>';
}

function sortByDate(a, b) {
  const x = isDate(a.date) ? a.date : '9999', y = isDate(b.date) ? b.date : '9999';
  return x < y ? -1 : x > y ? 1 : 0;
}

// ---------- 論文 ----------
function weeks() {
  return Array.from(new Set(data.papers.map(p => p.week))).sort().reverse();
}
function paperCard(p) {
  const url = safeUrl(p.url);
  return '<article class="card">' +
    '<div class="item-title">' + (url ? '<a href="' + esc(url) + '" target="_blank" rel="noopener">' + esc(p.title) + '</a>' : esc(p.title)) + '</div>' +
    '<div class="meta"><span>' + esc(p.authors) + '</span><span><i>' + esc(p.journal) + '</i></span><span>' + esc(p.year) + '</span></div>' +
    (p.theme ? '<div class="chips"><span class="chip">' + esc(p.theme) + '</span></div>' : '') +
    (p.summary_ja ? '<p class="summary">' + esc(p.summary_ja) + '</p>' : '') +
  '</article>';
}

// ---------- 各タブの描画 ----------
function renderHome() {
  const s = data.seminars;
  const soon = s.filter(x => ['新着', '検討中'].includes(x.status) && (d => d !== null && d >= 0 && d <= 14)(daysUntil(x.deadline)))
    .sort((a, b) => a.deadline < b.deadline ? -1 : 1);
  const next = s.filter(x => x.status === '申込済' && (daysUntil(x.date) === null || daysUntil(x.date) >= 0))
    .sort(sortByDate).slice(0, 3);
  const nNew = s.filter(x => x.status === '新着').length;
  const nConsider = s.filter(x => x.status === '検討中').length;
  const w = weeks()[0];
  const thisWeek = data.papers.filter(p => p.week === w);

  let h = '';
  if (DEMO) h += '<div class="banner">見本データを表示しています(実データではありません)</div>';
  h += '<div class="stats">' +
    '<a class="stat" href="#" data-go="seminars" data-filter="新着"><b>' + nNew + '</b><span>新着の勉強会</span></a>' +
    '<a class="stat" href="#" data-go="seminars" data-filter="検討中"><b>' + nConsider + '</b><span>検討中</span></a>' +
    '<a class="stat" href="#" data-go="papers"><b>' + thisWeek.length + '</b><span>今週の論文</span></a>' +
  '</div>';
  h += '<h3>締切が近い(2週間以内)</h3>';
  h += soon.length ? soon.map(seminarCard).join('') : '<p class="empty">締切が近いものはありません</p>';
  h += '<h3>次の予定(申込済)</h3>';
  h += next.length ? next.map(seminarCard).join('') : '<p class="empty">申込済みの予定はありません</p>';
  if (w) {
    const issue = safeUrl((thisWeek.find(p => p.issue_url) || {}).issue_url);
    h += '<div class="week-head"><h3>今週の論文(' + esc(fmtDate(w)) + ' 収集)</h3>' +
      (issue ? '<a class="small" href="' + esc(issue) + '" target="_blank" rel="noopener">Issue で選ぶ</a>' : '') + '</div>';
    h += thisWeek.slice(0, 3).map(paperCard).join('');
    if (thisWeek.length > 3) h += '<a href="#" data-go="papers">残り ' + (thisWeek.length - 3) + ' 本を見る</a>';
  }
  document.getElementById('tab-home').innerHTML = h;
}

function renderSeminars() {
  const names = ['予定', ...STATUSES, '全部'];
  let list = data.seminars;
  if (seminarFilter === '予定') list = list.filter(x => !['参加済', '見送り'].includes(x.status));
  else if (seminarFilter !== '全部') list = list.filter(x => x.status === seminarFilter);
  list = list.slice().sort(sortByDate);
  if (seminarFilter === '参加済') list.reverse();

  let h = '<div class="filters">' + names.map(n => {
    const c = n === '予定' ? data.seminars.filter(x => !['参加済', '見送り'].includes(x.status)).length
      : n === '全部' ? data.seminars.length : data.seminars.filter(x => x.status === n).length;
    return '<button type="button" data-filter="' + n + '"' + (n === seminarFilter ? ' class="on"' : '') + '>' + n + ' ' + c + '</button>';
  }).join('') + '</div>';
  h += list.length ? list.map(seminarCard).join('') : '<p class="empty">該当する勉強会はありません</p>';
  document.getElementById('tab-seminars').innerHTML = h;
}

function renderPapers() {
  let h = '<p class="small">※要約は AI(Gemini)によるものです。判断は必ず原文で確認してください。</p>';
  const ws = weeks();
  if (!ws.length) h += '<p class="empty">論文はまだありません</p>';
  ws.forEach(w => {
    const ps = data.papers.filter(p => p.week === w);
    const issue = safeUrl((ps.find(p => p.issue_url) || {}).issue_url);
    h += '<div class="week-head"><h2>' + esc(w) + ' 収集(' + ps.length + '本)</h2>' +
      (issue ? '<a class="small" href="' + esc(issue) + '" target="_blank" rel="noopener">Issue で選ぶ</a>' : '') + '</div>';
    h += ps.map(paperCard).join('');
  });
  document.getElementById('tab-papers').innerHTML = h;
}

function renderAll() {
  renderHome();
  renderSeminars();
  renderPapers();
  document.getElementById('fetched-at').textContent = data.fetched_at || 'まだありません';
}

// ---------- 通信 ----------
async function fetchData() {
  if (DEMO) return;
  const c = getConf();
  if (!c.url || !c.token) {
    banner('設定タブでサーバーの URL と合言葉を入れてください');
    return;
  }
  try {
    const r = await fetch(c.url + '?token=' + encodeURIComponent(c.token), { cache: 'no-store' });
    const res = await r.json();
    if (!res.ok) throw new Error(res.error || '取得に失敗しました');
    data = { seminars: res.seminars, papers: res.papers, fetched_at: new Date().toLocaleString('ja-JP') };
    lsSet(K_DATA, data);
    banner('');
    renderAll();
  } catch (e) {
    banner('最新のデータを取れませんでした(' + e.message + ')。前回のデータを表示しています');
  }
}

async function post(payload) {
  const c = getConf();
  const r = await fetch(c.url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ token: c.token }, payload)),
  });
  const res = await r.json();
  if (!res.ok) throw new Error(res.error || '保存に失敗しました');
  return res;
}

async function changeStatus(id, status) {
  const s = data.seminars.find(x => x.id === id);
  if (!s) return;
  const before = s.status;
  s.status = status;
  renderAll();
  if (DEMO) return;
  try {
    await post({ action: 'setStatus', id: id, status: status });
    lsSet(K_DATA, data);
  } catch (e) {
    s.status = before;
    renderAll();
    banner('状態を変えられませんでした(' + e.message + ')。電波の良い所でもう一度試してください');
  }
}

// ---------- 操作 ----------
function showTab(name) {
  document.querySelectorAll('.tab').forEach(t => { t.hidden = t.id !== 'tab-' + name; });
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === name));
  window.scrollTo(0, 0);
}

document.addEventListener('click', e => {
  const tabBtn = e.target.closest('.tabs button');
  if (tabBtn) { showTab(tabBtn.dataset.tab); return; }
  const go = e.target.closest('[data-go]');
  if (go) {
    e.preventDefault();
    if (go.dataset.filter) { seminarFilter = go.dataset.filter; renderSeminars(); }
    showTab(go.dataset.go);
    return;
  }
  const f = e.target.closest('.filters button');
  if (f) { seminarFilter = f.dataset.filter; renderSeminars(); }
});

document.addEventListener('change', e => {
  const sel = e.target.closest('select[data-id]');
  if (sel) changeStatus(sel.dataset.id, sel.value);
});

document.getElementById('reload').addEventListener('click', fetchData);

document.getElementById('conf-form').addEventListener('submit', async e => {
  e.preventDefault();
  const url = document.getElementById('conf-url').value.trim();
  const token = document.getElementById('conf-token').value.trim();
  const msg = document.getElementById('conf-msg');
  if (!/^https:\/\/script\.google\.com\//.test(url)) { msg.textContent = 'URL は https://script.google.com/ で始まるものを入れてください'; return; }
  lsSet(K_CONF, { url: url, token: token });
  msg.textContent = 'つないでいます…';
  await fetchData();
  msg.textContent = data.fetched_at ? '保存しました。データを取得できました' : '保存しましたが、取得できませんでした。URL と合言葉を確かめてください';
});

// ---------- 起動 ----------
(function start() {
  const c = getConf();
  document.getElementById('conf-url').value = c.url || '';
  document.getElementById('conf-token').value = c.token || '';
  if (DEMO) {
    data = demoData();
  } else {
    data = lsGet(K_DATA) || data;
  }
  renderAll();
  fetchData();
  if ('serviceWorker' in navigator && !DEMO) navigator.serviceWorker.register('sw.js').catch(() => {});
})();

// ---------- 見本データ(?demo のときだけ) ----------
function demoData() {
  const add = n => { const d = new Date(); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
  return {
    fetched_at: '見本',
    seminars: [
      { id: 'd1', name: '見本: 微生物検査研修会「血液培養の最近の話題」', organizer: '見本技師会', date: add(20), date_text: '14:00〜16:00', place: 'オンライン', format: 'Zoom', fee: '無料', deadline: add(2), url: 'https://example.com/', status: '新着', role: '', theme: '血液培養陽性時の報告体制' },
      { id: 'd2', name: '見本: 教育セミナー', organizer: '見本学会', date: add(35), date_text: '', place: '横浜', format: '現地', fee: '3,000円', deadline: add(10), url: '', status: '検討中', role: '', theme: '' },
      { id: 'd3', name: '見本: 精度管理セミナー', organizer: '見本学会', date: add(6), date_text: '', place: '', format: 'Zoom', fee: '無料', deadline: '', url: '', status: '申込済', role: '参加', theme: '' },
      { id: 'd4', name: '見本: 同定法の研修', organizer: '見本技師会', date: '', date_text: '11月予定(時間未発表)', place: '神奈川', format: '現地', fee: '', deadline: '', url: '', status: '申込済', role: '実務委員', theme: '' },
    ],
    papers: [
      { pmid: '0', week: add(-3), title: '見本: Rapid detection of carbapenemase producers from positive blood cultures', journal: 'J Clin Microbiol', year: '2026', authors: 'Sample A et al.', theme: '耐性菌', summary_ja: '見本の要約です。陽性血液培養から直接カルバペネマーゼを検出する方法を評価した。', url: 'https://pubmed.ncbi.nlm.nih.gov/', issue_url: '' },
    ],
  };
}
