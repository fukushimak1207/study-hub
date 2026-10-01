// 学びハブ サーバー側(Google Apps Script)
// データは「学びハブ データ」スプレッドシートに保存する。
// 読み書きには合言葉(TOKEN)が必要。合言葉は setup() が作ってスクリプトの保管庫に入れる。

const SEMINAR_COLS = [
  'id', 'name', 'organizer', 'date', 'date_text', 'place', 'format', 'theme',
  'fee', 'deadline', 'url', 'source', 'status', 'role', 'note', 'added_at', 'updated_at'
];
const PAPER_COLS = [
  'pmid', 'week', 'title', 'journal', 'year', 'authors', 'theme', 'summary_ja',
  'url', 'issue_url', 'added_at'
];
const STATUSES = ['新着', '検討中', '申込済', '参加済', '見送り'];
const PAPER_WEEKS_TO_SEND = 8;

// ---------------------------------------------------------------------------
// 初期設定(エディタから1回だけ手で実行する)
// ---------------------------------------------------------------------------
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SS_ID');
  if (!ssId) {
    ssId = SpreadsheetApp.create('学びハブ データ').getId();
    props.setProperty('SS_ID', ssId);
  }
  const ss = SpreadsheetApp.openById(ssId);
  ensureSheet_(ss, 'seminars', SEMINAR_COLS);
  ensureSheet_(ss, 'papers', PAPER_COLS);
  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);

  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
    Logger.log('合言葉を新しく作りました: ' + props.getProperty('TOKEN'));
  } else {
    Logger.log('合言葉は前のまま(変えたいときは TOKEN を消して再実行)');
  }
  Logger.log('データのシート: ' + ss.getUrl());

  // 受付箱(クラウドの巡回が勉強会の JSON を置くフォルダ)
  const inbox = inboxFolder_();
  Logger.log('受付箱: ' + inbox.getUrl());

  // 受付箱を1時間おきに取り込む(二重に作らない)
  const has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'importInbox';
  });
  if (!has) ScriptApp.newTrigger('importInbox').timeBased().everyHours(1).create();
  Logger.log('受付箱の自動取り込み: 1時間おき');
}

// ---------------------------------------------------------------------------
// 受付箱の取り込み(1時間おきに自動実行。手で実行してもよい)
// ---------------------------------------------------------------------------
const INBOX_NAME = '学びハブ受付箱';

function inboxFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('INBOX_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 消されていたら作り直す */ }
  }
  const it = DriveApp.getFoldersByName(INBOX_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(INBOX_NAME);
  props.setProperty('INBOX_ID', f.getId());
  return f;
}

// .json のファイルを読み、勉強会として追記してゴミ箱へ。読めないものは名前に印を付けて残す
function importInbox() {
  const files = inboxFolder_().getFiles();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      if (!/\.json$/i.test(name)) continue;
      let items;
      try {
        const d = JSON.parse(f.getBlob().getDataAsString('utf-8'));
        items = Array.isArray(d) ? d : (d.seminars || []);
      } catch (e) {
        f.setName('[読めない] ' + name.replace(/\.json$/i, '.txt'));
        continue;
      }
      const res = addSeminars_(items);
      Logger.log(name + ': 新規 ' + res.added + ' 件 / 補完 ' + res.filled + ' 件');
      f.setTrashed(true);
    }
  } finally {
    lock.releaseLock();
  }
}

function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  // 日付などを勝手に変換させない(すべて文字列で持つ)
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// 入口
// ---------------------------------------------------------------------------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkToken_(p.token)) return json_({ ok: false, error: '合言葉が違います' });
  return json_(list_());
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '送られてきたデータを読めません' });
  }
  if (!checkToken_(body.token)) return json_({ ok: false, error: '合言葉が違います' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'addSeminars': return json_(addSeminars_(body.items || []));
      case 'addPapers':   return json_(addPapers_(body.items || [], body.week, body.issue_url));
      case 'setStatus':   return json_(setStatus_(body.id, body.status));
      default:            return json_({ ok: false, error: '不明な操作: ' + body.action });
    }
  } finally {
    lock.releaseLock();
  }
}

function checkToken_(t) {
  const real = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !!real && t === real;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// シートの読み書き
// ---------------------------------------------------------------------------
function sheet_(name) {
  const ssId = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return SpreadsheetApp.openById(ssId).getSheetByName(name);
}

function readRows_(sh, cols) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = String(r[j]); });
    return o;
  });
}

function appendRows_(sh, cols, objs) {
  if (!objs.length) return;
  const values = objs.map(function (o) {
    return cols.map(function (c) { return o[c] == null ? '' : String(o[c]); });
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, cols.length).setValues(values);
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
}

// 名称の表記ゆれ(空白・記号・全角半角・「第◯回」)を吸収して照合する
function normName_(s) {
  return String(s || '').normalize('NFKC')
    .replace(/第\s*\d+\s*回/g, '')
    .replace(/[\s「」『』【】()（）・,，、。.:：]/g, '').toLowerCase();
}

// 同じ会か: 日付が同じで、名前の一方がもう一方を含む。日付が無いときは名前の完全一致だけ
function sameSeminar_(a, b) {
  const na = normName_(a.name), nb = normName_(b.name);
  if (!na || !nb) return false;
  if ((a.date || '') !== (b.date || '')) return false;
  if (!a.date) return na === nb;
  return na.indexOf(nb) >= 0 || nb.indexOf(na) >= 0;
}

function list_() {
  const seminars = readRows_(sheet_('seminars'), SEMINAR_COLS);
  let papers = readRows_(sheet_('papers'), PAPER_COLS);
  const weeks = Array.from(new Set(papers.map(function (p) { return p.week; })))
    .sort().reverse().slice(0, PAPER_WEEKS_TO_SEND);
  papers = papers.filter(function (p) { return weeks.indexOf(p.week) >= 0; });
  [seminars, papers].forEach(function (a) { a.forEach(function (o) { delete o._row; }); });
  return { ok: true, seminars: seminars, papers: papers, server_time: now_() };
}

// 同じ会(名称+日付)が既にあれば、空欄だけ埋める。状態は上書きしない。
function addSeminars_(items) {
  const sh = sheet_('seminars');
  const rows = readRows_(sh, SEMINAR_COLS);

  const added = [], filled = [];
  const t = now_();
  items.forEach(function (it) {
    if (!it.name) return;
    const ex = rows.filter(function (r) { return sameSeminar_(r, it); })[0];
    if (ex) {
      if (!ex._row) return; // 同じ送信の中での重複(まだシートに書いていない)は捨てる
      let changed = false;
      SEMINAR_COLS.forEach(function (c, j) {
        if (['id', 'status', 'added_at', 'updated_at'].indexOf(c) >= 0) return;
        if (!ex[c] && it[c]) {
          sh.getRange(ex._row, j + 1).setValue(String(it[c]));
          ex[c] = String(it[c]);
          changed = true;
        }
      });
      if (changed) {
        sh.getRange(ex._row, SEMINAR_COLS.indexOf('updated_at') + 1).setValue(t);
        filled.push(it.name);
      }
      return;
    }
    const o = {};
    SEMINAR_COLS.forEach(function (c) { o[c] = it[c] || ''; });
    o.id = Utilities.getUuid().slice(0, 8);
    o.status = STATUSES.indexOf(it.status) >= 0 ? it.status : '新着';
    o.added_at = t;
    o.updated_at = t;
    added.push(o);
    rows.push(o);
  });
  appendRows_(sh, SEMINAR_COLS, added);
  return { ok: true, added: added.length, filled: filled.length };
}

function addPapers_(items, week, issueUrl) {
  if (!week) return { ok: false, error: 'week(収集日)がありません' };
  const sh = sheet_('papers');
  const have = {};
  readRows_(sh, PAPER_COLS).forEach(function (r) { have[r.pmid] = true; });
  const t = now_();
  const add = [];
  items.forEach(function (it) {
    if (!it.pmid || have[it.pmid]) return;
    have[it.pmid] = true;
    add.push({
      pmid: it.pmid, week: week, title: it.title, journal: it.journal, year: it.year,
      authors: it.authors, theme: it.theme, summary_ja: it.summary_ja,
      url: 'https://pubmed.ncbi.nlm.nih.gov/' + it.pmid + '/',
      issue_url: issueUrl || '', added_at: t
    });
  });
  appendRows_(sh, PAPER_COLS, add);
  return { ok: true, added: add.length };
}

function setStatus_(id, status) {
  if (STATUSES.indexOf(status) < 0) return { ok: false, error: '不明な状態: ' + status };
  const sh = sheet_('seminars');
  const row = readRows_(sh, SEMINAR_COLS).filter(function (r) { return r.id === id; })[0];
  if (!row) return { ok: false, error: '見つかりません: ' + id };
  sh.getRange(row._row, SEMINAR_COLS.indexOf('status') + 1).setValue(status);
  sh.getRange(row._row, SEMINAR_COLS.indexOf('updated_at') + 1).setValue(now_());
  return { ok: true };
}
