/**
 * サイネージ管理スプレッドシート セットアップスクリプト
 *
 * 使い方:
 * 1. 新しいスプレッドシートを作成し「拡張機能 > Apps Script」を開く
 * 2. このファイルの中身を貼り付けて保存
 * 3. 関数「setupAll」を実行（初回は権限の承認が必要）
 * 4. 全シート・入力規則・サンプル行・忘れ物写真フォルダが自動で作られる
 *
 * 再実行しても既存データは消さない（ヘッダーと書式だけ整え直す）。
 */

const HEADER_BG = '#2f6fb2';
const HEADER_FG = '#ffffff';
const NOTE_FG = '#64748b';

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupSchedule_(ss);
  setupAnnual_(ss);
  setupReminder_(ss);
  setupNotices_(ss);
  setupRecruit_(ss);
  setupBirthday_(ss);
  setupNewcomer_(ss);
  setupLostItems_(ss);
  setupTicker_(ss);
  setupConfig_(ss);

  // 案内用の先頭シート（Sheet1）が残っていれば削除
  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);

  // スクリプトエディタから直接実行するとgetUi()が使えない場合があるため、実行ログに出力する
  Logger.log('セットアップ完了！各シートのサンプル行は自由に消してOKです。');
}

/* ================= 各シート ================= */

function setupSchedule_(ss) {
  const sh = getOrCreateSheet_(ss, '予定');
  header_(sh, ['日付', '終了日', '内容', '強調'], [110, 110, 520, 90]);
  noteRow_(sh, '今週分はスライド1「今週の予定」、それ以外の当月分はスライド2「今月のトピック」に自動で振り分けられます。');
  dateRule_(sh, 'A2:B500');
  listRule_(sh, 'D2:D500', ['締切']);
  sample_(sh, [
    [today_(2), '', '全国模試（高3・既卒）集合 8:40', ''],
    [today_(1), '', '夏期講習 面談申込', '締切'],
  ]);
}

function setupAnnual_(ss) {
  const sh = getOrCreateSheet_(ss, '年間予定');
  header_(sh, ['月', '日', '内容', '種別'], [70, 70, 520, 140]);
  noteRow_(sh, '毎年決まっている予定のマスタ。種別に「共通テスト」と書いた行がカウントダウンの基準日になります（毎年日付が変わるので年1回更新してください）。');
  listRule_(sh, 'D2:D200', ['共通テスト', '模試', '入試', 'イベント']);
  sample_(sh, [
    [1, 16, '大学入学共通テスト（1日目）', '共通テスト'],
    [1, 17, '大学入学共通テスト（2日目）', ''],
  ]);
}

function setupReminder_(ss) {
  const sh = getOrCreateSheet_(ss, '提出リマインド');
  header_(sh, ['内容', '締切日', '対象', '掲載終了日'], [420, 110, 90, 120]);
  noteRow_(sh, '締切日当日は「本日中」と赤字表示。掲載終了日が空欄なら締切日の翌日に自動で消えます。');
  dateRule_(sh, 'B2:B500');
  dateRule_(sh, 'D2:D500');
  listRule_(sh, 'C2:C500', ['講師', '生徒']);
  sample_(sh, [
    ['指導報告書（先週分）', today_(0), '講師', ''],
    ['全国模試申込', today_(1), '生徒', ''],
  ]);
}

const NOTICE_TEAMS = ['新人講師研修班', '現役講師研修班', '日程調整班', 'マッチング班', '成績向上班', '環境整備班', 'イベント班', '自習室管理'];

function setupNotices_(ss) {
  const sh = getOrCreateSheet_(ss, '校舎からの連絡');
  header_(sh, ['区分', '内容', '掲載終了日'], [130, 500, 120]);
  noteRow_(sh, '区分が「全体」なら校舎全体の連絡として、それ以外（班名）は班名付きの各班の連絡として表示されます。');
  dateRule_(sh, 'C2:C500');
  listRule_(sh, 'A2:A500', ['全体'].concat(NOTICE_TEAMS));
  sample_(sh, [
    ['全体', '台風接近のため明日は開校時間を変更する可能性があります。詳細は追って連絡します。', today_(2)],
    ['現役講師研修班', 'X月の現役講師研修を必ず受講してください！', today_(3)],
  ]);
}

function setupRecruit_(ss) {
  const sh = getOrCreateSheet_(ss, 'FL・DL募集');
  header_(sh, ['日付', '区分', '内容', '人数'], [110, 100, 320, 70]);
  noteRow_(sh, '「日付」「区分」は直前の行と同じなら空欄でOK（自動で引き継ぎます）。区分は FL / DL / 日曜日 / 事務 など自由入力。人数は空欄なら人数表示なしで時間帯だけ出ます。DLのように時間帯が複数ある日は「内容」欄にカンマ区切りで書けます（例: 13:00-16:00, 20:00-21:00）。');
  dateRule_(sh, 'A2:A500');
  listRule_(sh, 'B2:B500', ['FL', 'DL', '日曜日', '事務']);
  sample_(sh, [
    [today_(1), 'FL', '19:40-20:40', 1],
    ['', '', '20:45-21:45', 2],
    ['', 'DL', '13:00-16:00, 20:00-21:00', ''],
    [today_(2), 'FL', '18:35-19:35', 2],
    [today_(5), '日曜日', '午前 受付応援', 1],
  ]);
}

function setupBirthday_(ss) {
  const sh = getOrCreateSheet_(ss, '誕生日');
  header_(sh, ['名前', '月', '日'], [160, 70, 70]);
  noteRow_(sh, '一度登録すれば毎年、誕生日の前後（設定シートの日数）で自動表示されます。');
  sample_(sh, [['佐藤先生', 7, 11]]);
}

function setupNewcomer_(ss) {
  const sh = getOrCreateSheet_(ss, '新人紹介');
  header_(sh, ['名前', '紹介文', '掲載終了日'], [160, 480, 120]);
  dateRule_(sh, 'C2:C200');
  sample_(sh, [['山本先生', '理系・数学担当。よろしくお願いします！', today_(21)]]);
}

function setupLostItems_(ss) {
  const sh = getOrCreateSheet_(ss, '忘れもの');
  header_(sh, ['見つかった日', '品物', '場所', '状態', '写真ファイル名'], [120, 240, 180, 90, 200]);
  noteRow_(sh, '「返却済」にすると表示から消えます。写真は「忘れ物写真」フォルダに入れて、ここにファイル名を書くだけ（リンク取得は不要）。');
  dateRule_(sh, 'A2:A500');
  listRule_(sh, 'D2:D500', ['保管中', '返却済']);
  sample_(sh, [[today_(-1), '水色の水筒', '2F自習室', '保管中', '']]);
}

function setupTicker_(ss) {
  const sh = getOrCreateSheet_(ss, 'テロップ');
  header_(sh, ['内容', '掲載終了日'], [600, 120]);
  noteRow_(sh, '最下部の帯に流れます。空のときは「今日は何の日」だけが流れます。');
  dateRule_(sh, 'B2:B500');
}

function setupConfig_(ss) {
  const sh = getOrCreateSheet_(ss, '設定');
  header_(sh, ['Key', '値', '説明'], [200, 320, 420]);

  // 忘れ物写真フォルダを自動作成
  const folderId = getOrCreatePhotoFolder_(ss);

  const defaults = [
    ['FLカレンダーID', '', 'FL用GoogleカレンダーのID（カレンダー設定からコピー）'],
    ['DLカレンダーID', '', 'DL用GoogleカレンダーのID'],
    ['スライド切替秒数', 20, '右2/3スライドのローテーション間隔'],
    ['データ更新間隔(分)', 5, 'サイネージがデータを取り直す間隔'],
    ['誕生日表示日数', 2, '誕生日の何日前から表示するか（過去の誕生日は表示しない。今日から○日先まで）'],
    ['誕生日ポップアップ間隔(秒)', 90, '誕生日ポップアップが右上に出る間隔'],
    ['テロップ表示秒数', 8, 'テロップ1件あたりの表示時間（フェード切替の間隔）'],
    ['編集許可メール', '', '編集者ページを使えるGoogleアカウント（カンマ区切り）'],
    ['忘れ物写真フォルダID', folderId, '自動作成済み。写真はこのフォルダへ'],
  ];

  // 既存のKeyは値を保持し、無いKeyだけ追記
  const existing = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat()
    : [];
  const toAdd = defaults.filter(d => existing.indexOf(d[0]) === -1);
  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 3).setValues(toAdd);
  }
}

/* ================= ヘルパー ================= */

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function header_(sh, headers, widths) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontWeight('bold');
  sh.setFrozenRows(1);
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

// A列先頭が「#」の行はサイネージ側で無視される。使い方メモをシート内に残す
function noteRow_(sh, text) {
  if (sh.getLastRow() >= 2) return; // 既にデータがあれば追加しない
  sh.getRange(2, 1).setValue('# ' + text).setFontColor(NOTE_FG).setFontStyle('italic');
}

function sample_(sh, rows) {
  if (sh.getLastRow() >= 3) return; // 再実行時はサンプルを増やさない
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function dateRule_(sh, a1) {
  const rule = SpreadsheetApp.newDataValidation().requireDate()
    .setAllowInvalid(false).setHelpText('日付を 2026/07/12 の形式で入力してください').build();
  sh.getRange(a1).setDataValidation(rule);
}

function listRule_(sh, a1, values) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true)
    .setAllowInvalid(true).build();
  sh.getRange(a1).setDataValidation(rule);
}

function today_(offsetDays) {
  const d = new Date();
  d.setDate(d.getDate() + (offsetDays || 0));
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function getOrCreatePhotoFolder_(ss) {
  const name = '忘れ物写真';
  // スプレッドシートと同じ場所に作る
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const found = parent.getFoldersByName(name);
  if (found.hasNext()) return found.next().getId();
  return parent.createFolder(name).getId();
}
