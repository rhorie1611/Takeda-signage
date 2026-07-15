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
      countdown: null, fl: [], dl: [], schedule: [],
      reminders: [], notices: { overall: [], teams: [] }, recruit: { fl: [], dl: [], staff: [] },
      birthdays: [],
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
    schedule: schedule,
    reminders: getReminders_(ss, today),
    notices: getNotices_(ss, today),
    recruit: getRecruit_(ss, today),
    birthdays: getBirthdays_(ss, today, Number(cfg['誕生日表示日数']) || 3),
    ticker: getTicker_(ss, today, getNewcomers_(ss, today)),
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

/* ================= 予定（週・月を1本化。直近1週間は強調、それ以降は薄く） ================= */

function getSchedule_(ss, today) {
  // 直近7日間（今日を含む）を強調表示の対象にする
  const recentEnd = new Date(today);
  recentEnd.setDate(recentEnd.getDate() + 6);

  // 過ぎた予定も直近7日分はグレーで残す（それより前は今まで通り表示から外す）
  const pastStart = new Date(today);
  pastStart.setDate(pastStart.getDate() - 6);

  // 表示範囲の終わり = 今月末 or 今日から20日後、遅い方（月末間際でも近い将来が見えるように）
  const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  const minEnd = new Date(today);
  minEnd.setDate(minEnd.getDate() + 20);
  const rangeEnd = endOfMonth > minEnd ? endOfMonth : minEnd;

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

  return pool
    .filter(ev => {
      const last = ev.end || ev.date;
      return last >= pastStart && ev.date <= rangeEnd; // 1週間より前の過去、範囲外の遠い未来は除外
    })
    .sort((a, b) => a.date - b.date)
    .map(ev => {
      const last = ev.end || ev.date;
      const status = last < today ? 'past' : (ev.date <= recentEnd ? 'recent' : 'normal');
      return {
        label: dateLabel_(ev.date) + (ev.end ? '〜' + fmt_(ev.end, 'M/d') : ''),
        text: ev.text,
        emph: ev.emph === '締切',
        status: status,
      };
    });
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
      dueSort: due ? due.getTime() : Infinity, // 締切なしは最後に回す
    };
  })
    .filter(x => x && x.text)
    .sort((a, b) => a.dueSort - b.dueSort);
}

// 各班は常に全班分の枠を表示する（setup.gsのNOTICE_TEAMSが基準）。
// 該当する連絡がない班は text:null を返し、画面側で「ー」とグレー表示する。
function getNotices_(ss, today) {
  const items = readRows_(ss, '校舎からの連絡').map(r => {
    const end = asDate_(r[2]);
    if (end && end < today) return null;
    const scope = String(r[0] || '').trim();
    return { scope: scope, text: String(r[1] || '') };
  }).filter(x => x && x.text);

  const byTeam = {};
  items.filter(x => x.scope && x.scope !== '全体').forEach(x => {
    (byTeam[x.scope] = byTeam[x.scope] || []).push(x.text);
  });
  const teamList = (typeof NOTICE_TEAMS !== 'undefined' && NOTICE_TEAMS.length) ? NOTICE_TEAMS : Object.keys(byTeam);

  return {
    overall: items.filter(x => !x.scope || x.scope === '全体').map(x => x.text),
    teams: teamList.map(team => ({ team: team, text: byTeam[team] ? byTeam[team].join('\n') : null })),
  };
}

// 「FL・DL募集」シート: 日付・区分は空欄なら直前の行の値を引き継ぐ（原稿の縦並びをそのまま転記しやすくするため）
// 表示件数の上限は設けない。過去日を除いて書いてあるものは全部載せる
// 画面は上2/3がFL/DLの表、下1/3が日曜日・事務などその他の募集という構成
function getRecruit_(ss, today) {
  let lastDate = null, lastType = '';
  const parsed = [];
  readRows_(ss, 'FL・DL募集').forEach(r => {
    const d = asDate_(r[0]) || lastDate;
    const type = String(r[1] || '').trim() || lastType;
    const content = String(r[2] || '').trim();
    if (!d || !type || !content) return;
    lastDate = d; lastType = type;
    const count = r[3];
    parsed.push({ date: d, type: type, content: content, count: (count !== '' && count != null) ? count : null });
  });

  const future = parsed.filter(p => p.date >= today).sort((a, b) => a.date - b.date);

  // FL/DLは同じ日付を1行にまとめ、人数は表示しない（時間帯だけを列挙する）
  const groupByDate = list => {
    const byDate = {};
    list.forEach(p => {
      const key = p.date.getTime();
      if (!byDate[key]) byDate[key] = { date: p.date, contents: [] };
      byDate[key].contents.push(p.content);
    });
    return Object.keys(byDate).map(Number).sort((a, b) => a - b)
      .map(key => ({ label: dateLabel_(byDate[key].date), content: byDate[key].contents.join(', ') }));
  };

  const fl = groupByDate(future.filter(p => p.type === 'FL'));
  const dl = groupByDate(future.filter(p => p.type === 'DL'));

  const byDate = {};
  future.filter(p => p.type !== 'FL' && p.type !== 'DL').forEach(p => {
    const key = p.date.getTime();
    if (!byDate[key]) byDate[key] = { date: p.date, types: {} };
    (byDate[key].types[p.type] = byDate[key].types[p.type] || [])
      .push({ content: p.content, count: p.count });
  });
  const staff = Object.keys(byDate).map(Number).sort((a, b) => a - b)
    .map(key => {
      const g = byDate[key];
      return {
        label: dateLabel_(g.date),
        parts: Object.keys(g.types).map(type => ({
          type: type,
          items: g.types[type],
        })),
      };
    });

  return { fl: fl, dl: dl, staff: staff };
}

// 誕生日はテロップに混ぜず、画面右上に一定間隔で出るポップアップ専用にする（他の情報に埋もれて見逃されるのを防ぐため）
function getBirthdays_(ss, today, windowDays) {
  const items = [];
  readRows_(ss, '誕生日').forEach(r => {
    const name = String(r[0] || ''), m = Number(r[1]), d = Number(r[2]);
    if (!name || !m || !d) return;
    // 今年の誕生日との日数差（年またぎ考慮）
    for (const y of [today.getFullYear() - 1, today.getFullYear(), today.getFullYear() + 1]) {
      const bd = new Date(y, m - 1, d);
      const diff = Math.round((bd - today) / 86400000);
      if (Math.abs(diff) <= windowDays) {
        // 今日が誕生日でない場合に「今日が誕生日」と誤解されないよう、日付と時制をはっきり書く
        const dateStr = m + '/' + d;
        const text = diff === 0 ? name + ' 本日お誕生日です！'
          : diff > 0 ? name + ' ' + dateStr + 'にお誕生日を迎えます'
          : name + ' ' + dateStr + 'にお誕生日でした';
        items.push(text);
        break;
      }
    }
  });
  return items;
}

function getNewcomers_(ss, today) {
  const items = [];
  readRows_(ss, '新人紹介').forEach(r => {
    const end = asDate_(r[2]);
    if (end && end < today) return;
    if (r[0]) items.push({ tag: '新人紹介', text: String(r[0]) + '　' + String(r[1] || '') });
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

// テロップには「テロップ」シートに加えて、優先度が低めのコーナー（新人紹介・忘れもの）と
// 「今日は何の日」もまとめて流す（専用パネルを持たせるほどではないが、消したくはない情報のため）。
// 誕生日だけは他の情報に埋もれて見逃されやすいため、別枠のポップアップで表示する（getBirthdays_）。
function getTicker_(ss, today, newcomers) {
  const items = readRows_(ss, 'テロップ').map(r => {
    const end = asDate_(r[1]);
    if (end && end < today) return null;
    return r[0] ? { tag: 'お知らせ', text: String(r[0]) } : null;
  }).filter(Boolean);

  newcomers.forEach(p => items.push(p));

  getLost_(ss).forEach(l => {
    items.push({ tag: '忘れもの', text: l.item + (l.place ? '（' + l.place + '）' : '') + ' → 受付で保管しています' });
  });

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
