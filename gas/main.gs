/**
 * サイネージ本体
 * - doGet: サイネージ画面（index.html）を配信
 * - getData: カレンダー・各シートを読み取り、表示用データを返す（画面から定期的に呼ばれる）
 *
 * setup.gs と同じ Apps Script プロジェクト（サイネージ管理スプレッドシートに紐づく）に置くこと。
 */

const TZ = 'Asia/Tokyo';

function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  let data;
  try {
    data = getData();
  } catch (err) {
    // シート未設定などでも画面自体は表示する（次の自動更新で復帰）
    data = {
      config: { slideSeconds: 20, refreshMinutes: 5 },
      countdown: null, fl: [], dl: [], weekly: [], monthly: [],
      reminders: [], teams: [], recruit: [], people: [], lost: [],
      ticker: [{ tag: 'エラー', text: 'データ取得に失敗しました: ' + err }],
    };
  }
  t.initial = JSON.stringify(data);
  return t.evaluate()
    .setTitle('サイネージ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = readConfig_(ss);
  const today = startOfDay_(new Date());
  const schedule = getSchedule_(ss, today);

  return {
    generatedAt: fmt_(new Date(), 'yyyy/MM/dd HH:mm'),
    config: {
      slideSeconds: Number(cfg['スライド切替秒数']) || 20,
      refreshMinutes: Number(cfg['データ更新間隔(分)']) || 5,
    },
    countdown: getCountdown_(ss, today),
    fl: getLeaders_(cfg['FLカレンダーID'], true),
    dl: getLeaders_(cfg['DLカレンダーID'], false),
    weekly: schedule.weekly,
    monthly: schedule.monthly,
    reminders: getReminders_(ss, today),
    teams: getTeams_(ss, today),
    recruit: getRecruit_(ss, today),
    people: getPeople_(ss, today, Number(cfg['誕生日表示日数']) || 3),
    lost: getLost_(ss, cfg['忘れ物写真フォルダID']),
    ticker: getTicker_(ss, today),
  };
}

/* ================= FL/DL（カレンダー） ================= */

function getLeaders_(calId, requireFlTag) {
  if (!calId) return [];
  try {
    const cal = CalendarApp.getCalendarById(String(calId).trim());
    if (!cal) return [];
    return cal.getEventsForDay(new Date())
      .filter(ev => !ev.isAllDayEvent())
      .map(ev => ({
        s: fmt_(ev.getStartTime(), 'HH:mm'),
        e: fmt_(ev.getEndTime(), 'HH:mm'),
        who: parseName_(ev.getTitle(), requireFlTag),
      }))
      .filter(x => x.who)
      .sort((a, b) => a.s.localeCompare(b.s));
  } catch (err) {
    return [{ s: '', e: '', who: 'カレンダー取得エラー: ' + err }];
  }
}

/**
 * カレンダー取得の切り分け用。Apps Scriptエディタで実行し、
 * 「実行ログ」に出る内容をそのまま確認してください。
 */
function debugCalendars_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = readConfig_(ss);
  ['FLカレンダーID', 'DLカレンダーID'].forEach(key => {
    const id = String(cfg[key] || '').trim();
    Logger.log(key + ' = "' + id + '"');
    if (!id) { Logger.log('  → 設定シートが空です'); return; }
    try {
      const cal = CalendarApp.getCalendarById(id);
      if (!cal) {
        Logger.log('  → getCalendarById は null を返しました（このアカウントに共有されていない/IDが違う可能性）');
        return;
      }
      Logger.log('  → OK: ' + cal.getName() + '（本日のイベント数: ' + cal.getEventsForDay(new Date()).length + '）');
    } catch (err) {
      Logger.log('  → エラー: ' + err);
    }
  });
  Logger.log('実行アカウント: ' + Session.getActiveUser().getEmail());
}

// FL: 「※【FL】井上、岡田」→「井上、岡田」。【FL】から始まらない予定（面談・面接練習等）は除外
// DL: タグなしでそのままの予定名を使う
function parseName_(title, requireFlTag) {
  const raw = String(title || '').replace(/^[※\s]+/, '').trim();
  if (requireFlTag) {
    const m = raw.match(/^【FL】\s*(.*)$/);
    return m ? m[1].trim() : '';
  }
  return raw.replace(/【[^】]*】/g, '').trim();
}

/* ================= 予定（週・月 自動振り分け） ================= */

function getSchedule_(ss, today) {
  // 今週 = 今日を含む月曜〜日曜
  const monday = new Date(today);
  monday.setDate(monday.getDate() - ((today.getDay() + 6) % 7));
  const sunday = new Date(monday);
  sunday.setDate(sunday.getDate() + 6);

  const pool = [];

  // 「予定」シート
  readRows_(ss, '予定').forEach(r => {
    const d = asDate_(r[0]);
    if (!d) return;
    pool.push({ date: d, end: asDate_(r[1]), text: String(r[2] || ''), emph: String(r[3] || '') });
  });

  // 「年間予定」シート → 次に来るその月日に展開
  readRows_(ss, '年間予定').forEach(r => {
    const m = Number(r[0]), day = Number(r[1]);
    if (!m || !day || !r[2]) return;
    let d = new Date(today.getFullYear(), m - 1, day);
    if (d < today) d = new Date(today.getFullYear() + 1, m - 1, day);
    pool.push({ date: d, end: null, text: String(r[2]), emph: '' });
  });

  const weekly = [], monthly = [];
  pool.forEach(ev => {
    const last = ev.end || ev.date;
    if (last < today) return; // 終わった予定
    if (ev.date <= sunday && last >= monday) {
      weekly.push(ev);
    } else if (ev.date.getFullYear() === today.getFullYear() && ev.date.getMonth() === today.getMonth()) {
      monthly.push(ev);
    }
  });
  const toRow = ev => ({
    label: dateLabel_(ev.date) + (ev.end ? '〜' + fmt_(ev.end, 'M/d') : ''),
    text: ev.text,
    emph: ev.emph === '締切',
  });
  const bySort = (a, b) => a.date - b.date;
  return {
    weekly: weekly.sort(bySort).map(toRow),
    monthly: monthly.sort(bySort).map(toRow),
  };
}

/* ================= 各コーナー ================= */

function getReminders_(ss, today) {
  return readRows_(ss, '提出リマインド').map(r => {
    const due = asDate_(r[1]);
    const end = asDate_(r[3]) || due; // 掲載終了日が空欄なら締切日まで
    if (end && end < today) return null;
    return {
      text: String(r[0] || ''),
      target: String(r[2] || ''),
      due: due ? (due.getTime() === today.getTime() ? '本日中' : fmt_(due, 'M/d') + 'まで') : '',
    };
  }).filter(x => x && x.text);
}

function getTeams_(ss, today) {
  return readRows_(ss, '各班の伝達').map(r => {
    const end = asDate_(r[2]);
    if (end && end < today) return null;
    return { team: String(r[0] || ''), text: String(r[1] || '') };
  }).filter(x => x && x.text);
}

function getRecruit_(ss, today) {
  return readRows_(ss, 'シフト・事務募集').map(r => {
    const end = asDate_(r[3]);
    if (end && end < today) return null;
    const when = asDate_(r[0]);
    return {
      when: when ? dateLabel_(when) : String(r[0] || ''),
      text: String(r[1] || ''),
      left: r[2] !== '' && r[2] != null ? 'あと' + r[2] + '名' : '',
    };
  }).filter(x => x && x.text);
}

function getPeople_(ss, today, windowDays) {
  const items = [];
  readRows_(ss, '誕生日').forEach(r => {
    const name = String(r[0] || ''), m = Number(r[1]), d = Number(r[2]);
    if (!name || !m || !d) return;
    // 今年の誕生日との日数差（年またぎ考慮）
    for (const y of [today.getFullYear() - 1, today.getFullYear(), today.getFullYear() + 1]) {
      const bd = new Date(y, m - 1, d);
      const diff = Math.round((bd - today) / 86400000);
      if (Math.abs(diff) <= windowDays) {
        items.push({ tag: m + '/' + d, text: name + ' お誕生日！🎉' });
        break;
      }
    }
  });
  readRows_(ss, '新人紹介').forEach(r => {
    const end = asDate_(r[2]);
    if (end && end < today) return;
    if (r[0]) items.push({ tag: '新人', text: String(r[0]) + '　' + String(r[1] || '') });
  });
  return items;
}

function getLost_(ss, folderId) {
  const items = readRows_(ss, '忘れもの')
    .filter(r => String(r[3] || '保管中') !== '返却済' && r[1])
    .map(r => ({
      item: String(r[1]),
      place: String(r[2] || ''),
      photoName: String(r[4] || '').trim(),
      photo: null,
    }));

  // 写真（最大4枚、サムネイルをbase64で同梱）
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(String(folderId).trim());
      let count = 0;
      items.forEach(it => {
        if (!it.photoName || count >= 4) return;
        const files = folder.getFilesByName(it.photoName);
        if (!files.hasNext()) return;
        const blob = files.next().getThumbnail();
        if (blob) {
          it.photo = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
          count++;
        }
      });
    } catch (err) { /* フォルダ未設定・権限なしは写真なしで続行 */ }
  }
  items.forEach(it => delete it.photoName);
  return items;
}

function getTicker_(ss, today) {
  const items = readRows_(ss, 'テロップ').map(r => {
    const end = asDate_(r[1]);
    if (end && end < today) return null;
    return r[0] ? { tag: 'お知らせ', text: String(r[0]) } : null;
  }).filter(Boolean);

  const ann = getAnniversary_(today);
  if (ann) {
    items.push({
      tag: '今日は何の日',
      text: (today.getMonth() + 1) + '月' + today.getDate() + '日は「' + ann + '」',
    });
  }
  return items;
}

function getCountdown_(ss, today) {
  let best = null;
  readRows_(ss, '年間予定').forEach(r => {
    if (String(r[3] || '') !== '共通テスト') return;
    const m = Number(r[0]), d = Number(r[1]);
    if (!m || !d) return;
    let dt = new Date(today.getFullYear(), m - 1, d);
    if (dt < today) dt = new Date(today.getFullYear() + 1, m - 1, d);
    if (!best || dt < best) best = dt;
  });
  if (!best) return null;
  return { label: '共通テストまで', days: Math.round((best - today) / 86400000) };
}

/* ================= ヘルパー ================= */

function readConfig_(ss) {
  const cfg = {};
  readRows_(ss, '設定').forEach(r => { if (r[0]) cfg[String(r[0]).trim()] = r[1]; });
  return cfg;
}

// ヘッダー行を除き、空行と「#」始まりの行を捨てて返す
function readRows_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(r => r.some(c => c !== '' && c != null))
    .filter(r => String(r[0]).charAt(0) !== '#');
}

function asDate_(v) {
  if (v instanceof Date) return startOfDay_(v);
  if (typeof v === 'string' && v.trim()) {
    const m = v.trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  return null;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, TZ, pattern);
}

const WEEKDAYS_JA = ['日', '月', '火', '水', '木', '金', '土'];
function dateLabel_(d) {
  return fmt_(d, 'M/d') + '(' + WEEKDAYS_JA[d.getDay()] + ')';
}
