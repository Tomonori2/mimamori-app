/**
 * みまもり（GAS版） v1.0
 *
 * これは「古いスマホのサーバー」の代わりに、Googleの上で動く見守りの本体です。
 * やることは今までと同じ4つ。
 *   ① 押した時刻を記録する      → スプレッドシートの「記録」シート
 *   ② 毎朝チェックする          → Googleの時間トリガー（15分おき）
 *   ③ 合図がなければ知らせる    → Gmail
 *   ④ ボタン画面に応答する      → doGet / doPost（ウェブアプリ）
 *
 * 大事な考え方：
 *   毎日1通「きょうも元気です」が届くこと自体が、この仕組みが生きている証明です。
 *   届かない日は「本人か、この仕組みのどちらかに何かあった」と分かります。
 *   だから毎日の報告は、面倒でも切らないでください。
 */

var VERSION = 'v1.1';

// 体調の3択。キーと、画面やメールに出す言葉の対応表。
var MOODS = {
  genki:   { label: '元気です',   short: '元気' },
  futsu:   { label: 'ふつうです', short: 'ふつう' },
  shindoi: { label: 'しんどい',   short: 'しんどい' }
};

var SHEET_SETTINGS = '設定';
var SHEET_HISTORY  = '記録';
var SHEET_STATE    = '内部データ';
var SHEET_MAILLOG  = '送信ログ';


/* ============================================================
   メニュー（スプレッドシートを開くと上に「みまもり」が出ます）
   ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('みまもり')
    .addItem('① 初期設定（最初に1回だけ）', 'setup')
    .addItem('② テスト通知を送ってみる', 'testAlert')
    .addItem('③ いまの状態を見る', 'showStatus')
    .addItem('④ きょうの記録をリセット（テスト用）', 'resetToday')
    .addToUi();
}

/** テスト用。「まだ何も押していない」状態に戻す。
 *  「内部データ」シートを手で消すより安全で確実。 */
function resetToday() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert(
    'きょうの記録（押した・知らせた）を消して、まだ何もしていない状態に戻します。\n\nよろしいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (ans !== ui.Button.OK) return;

  ['last_checkin', 'alert_date', 'clear_date', 'report_date', 'today_mood']
    .forEach(function (k) { setState_(k, ''); });

  ui.alert('きょうの記録を消しました。\n\n'
    + '「設定」シートの締め切り時刻を、いまより前の時刻にすると、\n'
    + '15分以内に「合図がありません」が届きます。\n'
    + '確かめたら、締め切り時刻を元に戻してください。');
}


/* ============================================================
   初期設定：シートを用意して、15分おきの見張りを仕掛ける
   ============================================================ */
function setup() {
  var ui = SpreadsheetApp.getUi();

  ensureSheets_();

  var s = getSettings_();
  if (!s.contact_email) {
    ui.alert('「設定」シートの「連絡先メール（家族）」を先に入力してから、'
           + 'もう一度この初期設定を実行してください。');
    return;
  }

  // 古い見張りが残っていると二重に動くので、いったん全部消してから作り直す
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkDeadline') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkDeadline').timeBased().everyMinutes(15).create();

  sendMail_(
    '【みまもり】見守りを開始しました',
    Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm') + ' に、'
      + nameOf_(s) + 'さんの見守りを開始しました。\n\n'
      + '締め切り時刻：' + s.deadline + '\n'
      + '毎日の報告メール：' + (s.daily_report ? 'あり' : 'なし') + '\n\n'
      + '※このお知らせが予定外に届いたときは、見守りが一度止まって\n'
      + '　仕掛け直された、ということです。\n'
      + '（みまもり ' + VERSION + ' からの自動送信）\n'
  );

  ui.alert('初期設定が終わりました。\n\n'
         + '・15分おきに「合図があったか」を見張ります\n'
         + '・連絡先（' + s.contact_email + '）に開始のお知らせを送りました\n\n'
         + '次は「デプロイ」→「新しいデプロイ」でウェブアプリとして公開し、\n'
         + '出てきたURLをボタン画面の設定に貼り付けてください。');
}


/* ============================================================
   見張り：15分おきに動いて「締め切りを過ぎたのに合図がない」を探す
   ============================================================ */
function checkDeadline() {
  var s = getSettings_();
  if (!s.contact_email) return;

  var now = new Date();
  var today = ymd_(now);

  // ★「きょうの日付」も「いまの時刻」も、かならずスプレッドシートの時計で見る。
  //   以前は new Date().setHours() で締め切りを作っていたが、それは
  //   Apps Script側の時計で動くので、2つの設定がずれていると
  //   日付と時刻でちがう時計を混ぜることになる。
  //   いまは両方を同じ HH:mm の文字にそろえて比べているので、その心配がない。
  var nowHM = Utilities.formatDate(now, tz_(), 'HH:mm');

  var checkedToday = (stateDate_('last_checkin') === today);

  if (nowHM >= s.deadline && !checkedToday
      && stateDate_('alert_date') !== today && !alertSentToday_(today)) {
    doAlert_(s);
    setState_('alert_date', today);
  }
}


/** きょうすでに「合図がありません」を送っていないか、送信ログでも確かめる。
 *  内部データの保存がなにかの理由で失敗しても、
 *  同じ通知を何度も送りつけることだけは起こさないための保険。
 *  （一度これで失敗しているので、二重に見張っている） */
function alertSentToday_(today) {
  var sh = sheet_(SHEET_MAILLOG);
  var last = sh.getLastRow();
  if (last < 2) return false;

  var from = Math.max(2, last - 30);       // 直近30件だけ見れば十分
  var rows = sh.getRange(from, 1, last - from + 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    var d = rows[i][0];
    d = (d instanceof Date) ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd')
                            : String(d).slice(0, 10);
    if (d === today && String(rows[i][2]).indexOf('合図がありません') >= 0) return true;
  }
  return false;
}


/* ============================================================
   ボタン画面からの入り口（ウェブアプリ）
   ============================================================ */

// 状態を返す（画面を開いたとき）
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'status';
  if (action === 'checkin') {
    // POSTがうまくいかない環境のための逃げ道
    return json_(checkin_((e.parameter && e.parameter.mood) || 'genki'));
  }
  return json_(status_());
}

// 「元気です」を押したとき
// ※ 本文は text/plain で送ってもらいます。application/json だとブラウザが
//    事前確認（プリフライト）を挟み、GASはそれに答えられないため失敗します。
function doPost(e) {
  var data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    data = {};
  }
  if (data.action === 'status') return json_(status_());
  return json_(checkin_(data.mood || 'genki'));
}

function status_() {
  var s = getSettings_();
  var today = ymd_(new Date());
  return {
    ok: true,
    version: VERSION,
    name: s.name,
    deadline: s.deadline,
    region: s.region,
    heat_alert: s.heat_alert,
    checked_today: (stateDate_('last_checkin') === today),
    today_mood: getState_('today_mood'),
    history: recentRows_(7)
  };
}

function checkin_(mood) {
  if (!MOODS[mood]) mood = 'genki';

  var s = getSettings_();
  var now = new Date();
  var today = ymd_(now);

  setState_('last_checkin', Utilities.formatDate(now, tz_(), "yyyy-MM-dd'T'HH:mm:ss"));
  setState_('today_mood', mood);
  addHistory_(mood, now);

  // どのメールを出すかを決める（今までのPython版と同じ判断です）
  if (s.contact_email) {
    if (stateDate_('alert_date') === today && stateDate_('clear_date') !== today) {
      // すでに「合図がありません」を送ったあとに押された → 無事を伝える
      doAllClear_(s, now);
      setState_('clear_date', today);
      setState_('report_date', today);   // この日の報告は済んだ扱い
    } else if (s.daily_report && stateDate_('report_date') !== today) {
      // その日はじめての合図 → 毎日の報告
      doDailyReport_(s, now, mood);
      setState_('report_date', today);
    }
  }

  return { ok: true, time: Utilities.formatDate(now, tz_(), 'HH:mm'), mood: mood };
}


/* ============================================================
   メールの中身（Python版の文面をそのまま引き継いでいます）
   ============================================================ */
function doAlert_(s) {
  var name = nameOf_(s);
  sendMail_(
    '【みまもり】' + name + 'さんから、きょうの元気の合図がありません',
    name + 'さんから、本日の「元気です」の合図が、\n'
      + '締め切り時刻（' + s.deadline + '）までにありませんでした。\n\n'
      + '念のため、電話や訪問で様子を確認してください。\n\n'
      + '── 最近1週間の記録 ──\n'
      + recentLines_(7) + '\n\n'
      + '（この通知は みまもり ' + VERSION + ' から自動送信されています）\n'
  );
}

function doAllClear_(s, now) {
  var name = nameOf_(s);
  sendMail_(
    '【みまもり】' + name + 'さんから、遅れて元気の合図がありました',
    '先ほど「合図がありません」とお知らせしましたが、その後\n'
      + Utilities.formatDate(now, tz_(), 'HH:mm') + ' に '
      + name + 'さんが「元気です」を押されました。\n\n'
      + '押し忘れだったようです。ひとまずご安心ください。\n'
      + '（みまもり ' + VERSION + ' からの自動送信）\n'
  );
}

function doDailyReport_(s, now, mood) {
  var name = nameOf_(s);
  var hhmm = Utilities.formatDate(now, tz_(), 'HH:mm');
  var subject, head;

  if (mood === 'shindoi') {
    subject = '【みまもり】' + name + 'さん、きょうは体調がすぐれないようです';
    head = name + 'さんが本日 ' + hhmm + ' に合図をされましたが、\n'
         + '体調は「しんどい」を選ばれています。\n\n'
         + 'お時間があれば、声をかけてあげてください。\n';
  } else {
    subject = '【みまもり】' + name + 'さん、きょうも元気です';
    head = name + 'さんが本日 ' + hhmm + ' に合図をされました。\n'
         + '体調は「' + MOODS[mood].short + '」でした。\n';
  }

  sendMail_(subject,
    head + '\n'
      + '── 最近1週間の記録 ──\n'
      + recentLines_(7) + '\n\n'
      + '※ 押した時刻が少しずつ遅くなっていくときは、\n'
      + '　 体調が落ちてきているサインのことがあります。\n\n'
      + '※このお知らせは毎日1回届きます。\n'
      + '　届かない日は、ご本人か見守りの仕組みに何かあった可能性があります。\n'
      + '（みまもり ' + VERSION + ' からの自動送信）\n'
  );
}

/**
 * メールを送る。
 * ここで使っている MailApp は「メールを送信する」だけの許可で動きます。
 * 受信箱を読んだり消したりする権限までは求めません。
 * 人に配るものなので、要求する権限は狭いほどよい、という考え方です。
 */
function sendMail_(subject, body) {
  var s = getSettings_();
  if (!s.contact_email) return;
  MailApp.sendEmail(s.contact_email, subject, body);
  sheet_(SHEET_MAILLOG).appendRow([
    Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'),
    s.contact_email, subject
  ]);
}


/* ============================================================
   メニューの② ③
   ============================================================ */
function testAlert() {
  var s = getSettings_();
  if (!s.contact_email) {
    SpreadsheetApp.getUi().alert('先に「設定」シートの連絡先メールを入力してください。');
    return;
  }
  doAlert_(s);
  SpreadsheetApp.getUi().alert('テスト通知を ' + s.contact_email + ' に送りました。\n'
    + '受信箱を確認してください。');
}

function showStatus() {
  var st = status_();
  var triggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'checkDeadline';
  });
  SpreadsheetApp.getUi().alert(
    'みまもり ' + VERSION + '\n\n'
    + '見張り（15分おき）：' + (triggers.length ? '動いています ✓' : '止まっています ✗ → 初期設定をやり直してください') + '\n'
    + 'きょうの合図：' + (st.checked_today ? 'あり（' + (MOODS[st.today_mood] || {}).short + '）' : 'まだありません') + '\n'
    + '締め切り時刻：' + st.deadline
  );
}


/* ============================================================
   シートの読み書き（ここから下は裏方です）
   ============================================================ */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Tokyo'; }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function nameOf_(s) { return s.name || '見守り対象の方'; }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) sh = ss_().insertSheet(name);
  return sh;
}

function ensureSheets_() {
  var st = sheet_(SHEET_SETTINGS);
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 7, 2).setValues([
      ['項目', '値'],
      ['見守る方のお名前', ''],
      ['連絡先メール（家族）', ''],
      ['締め切り時刻', '10:00'],
      ['毎日の報告メール', 'する'],
      ['熱中症の警戒を出す', 'する'],
      ['お住まいの都道府県', '愛知県']
    ]);
    st.setColumnWidth(1, 200);
    st.setColumnWidth(2, 260);
    st.getRange('A1:B1').setFontWeight('bold');
  }

  var hi = sheet_(SHEET_HISTORY);
  if (hi.getLastRow() === 0) hi.appendRow(['日付', '時刻', '体調']);

  var ml = sheet_(SHEET_MAILLOG);
  if (ml.getLastRow() === 0) ml.appendRow(['日時', '宛先', '件名']);

  sheet_(SHEET_STATE);
}

function getSettings_() {
  var sh = ss_().getSheetByName(SHEET_SETTINGS);
  if (!sh || sh.getLastRow() < 2) {
    return { name: '', contact_email: '', deadline: '10:00',
             daily_report: true, heat_alert: true, region: '愛知県' };
  }
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var v = {};
  rows.forEach(function (r) { v[String(r[0]).trim()] = r[1]; });

  // 締め切り時刻は、手で「10:00」と書いても、時刻として入力されても読めるようにする。
  // さらに「9:00」のような1桁の書き方を「09:00」にそろえる。
  // そろえないと文字として比べたときに 9:00 > 10:00 と判定されてしまう。
  var dl = v['締め切り時刻'];
  if (dl instanceof Date) dl = Utilities.formatDate(dl, tz_(), 'HH:mm');
  dl = String(dl || '10:00').trim();
  var hm = dl.match(/^(\d{1,2}):(\d{2})/);
  dl = hm ? ('0' + hm[1]).slice(-2) + ':' + hm[2] : '10:00';

  return {
    name: String(v['見守る方のお名前'] || '').trim(),
    contact_email: String(v['連絡先メール（家族）'] || '').trim(),
    deadline: dl,
    daily_report: String(v['毎日の報告メール'] || 'する').trim() !== 'しない',
    heat_alert: String(v['熱中症の警戒を出す'] || 'する').trim() !== 'しない',
    region: String(v['お住まいの都道府県'] || '愛知県').trim()
  };
}

function getState_(key) {
  var sh = sheet_(SHEET_STATE);
  if (sh.getLastRow() === 0) return '';
  var rows = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      var v = rows[i][1];
      // 古い版が「日付データ」として保存してしまった分を、文字にそろえ直す。
      if (v instanceof Date) return Utilities.formatDate(v, tz_(), "yyyy-MM-dd'T'HH:mm:ss");
      return String(v || '');
    }
  }
  return '';
}

/** 内部データの日付を、かならず「yyyy-MM-dd」の10文字にそろえて取り出す。
 *  保存の形がぶれても比較が壊れないようにするための入口。 */
function stateDate_(key) {
  return String(getState_(key) || '').slice(0, 10);
}

function setState_(key, value) {
  var sh = sheet_(SHEET_STATE);
  var last = sh.getLastRow();
  var row = 0;

  if (last > 0) {
    var rows = sh.getRange(1, 1, last, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === key) { row = i + 1; break; }
    }
  }
  if (!row) {
    row = last + 1;
    sh.getRange(row, 1).setValue(key);
  }

  // ★「@」＝文字として扱う書式。これを付けないと、"2026-08-25" のような文字列を
  //   スプレッドシートが日付データに変換してしまい、読み返したとき別の形になる。
  //   その結果「きょうはもう知らせた」の記録が一致せず、15分ごとに通知が出続けた。
  var cell = sh.getRange(row, 2);
  cell.setNumberFormat('@');
  cell.setValue(value);
}

/** その日はじめて押した時刻を残す（あとから上書きしない。起きて動き出した時刻の目安になるので）。 */
function addHistory_(mood, now) {
  var sh = sheet_(SHEET_HISTORY);
  var today = ymd_(now);
  var last = sh.getLastRow();

  if (last >= 2) {
    var lastDate = sh.getRange(last, 1).getValue();
    if (lastDate instanceof Date) lastDate = ymd_(lastDate);
    if (String(lastDate) === today) {
      sh.getRange(last, 3).setValue(MOODS[mood].short);  // 体調だけ最新に更新
      return;
    }
  }
  sh.appendRow([today, Utilities.formatDate(now, tz_(), 'HH:mm'), MOODS[mood].short]);
}

function recentRows_(days) {
  var sh = sheet_(SHEET_HISTORY);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - days + 1);
  return sh.getRange(from, 1, last - from + 1, 3).getValues().map(function (r) {
    var d = r[0] instanceof Date ? ymd_(r[0]) : String(r[0]);
    return { d: d, t: String(r[1]), m: String(r[2]) };
  });
}

/** メール本文用に「08/23（土） 07:12  元気」の形にする。 */
function recentLines_(days) {
  var wd = ['日', '月', '火', '水', '木', '金', '土'];
  var rows = recentRows_(days);
  if (!rows.length) return '  （まだ記録がありません）';
  return rows.map(function (r) {
    var w = '  ';
    var parts = r.d.split('-');
    if (parts.length === 3) {
      w = wd[new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])).getDay()];
    }
    return '  ' + r.d.slice(5).replace('-', '/') + '（' + w + '） ' + r.t + '  ' + r.m;
  }).join('\n');
}
