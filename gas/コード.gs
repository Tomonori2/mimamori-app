/**
 * みまもり（GAS版） v1.4
 *
 * これは「古いスマホのサーバー」の代わりに、Googleの上で動く見守りの本体です。
 * やることは今までと同じ4つ。
 *   ① 押した時刻を記録する      → スプレッドシートの「記録」シート
 *   ② 毎朝チェックする          → Googleの時間トリガー（15分おき）
 *   ③ 合図がなければ知らせる    → メール送信（MailApp）
 *   ④ ボタン画面に応答する      → doGet / doPost（ウェブアプリ）
 *
 * 大事な考え方：
 *   毎日1通「きょうも元気です」が届くこと自体が、この仕組みが生きている証明です。
 *   届かない日は「本人か、この仕組みのどちらかに何かあった」と分かります。
 *   だから毎日の報告は、面倒でも切らないでください。
 *
 * v1.2で足したもの：合図があった「場所」
 *   ボタンを押したその瞬間の位置だけを記録します。
 *   ★ できないこと：アプリを閉じている間、裏でずっと居場所を追いかけること。
 *     ブラウザのアプリにはその力がありません。だから「合図がありません」の
 *     お知らせに出る場所は、あくまで【最後に押したときの場所】です。
 *     いま現在どこにいるか、ではありません。ここを取り違えないでください。
 *
 * v1.3で直したもの：テスト通知が本物の見張りを止めていた不具合
 *   メニュー②のテスト通知が、本物のお知らせと件名まで同じでした。
 *   見張りは「きょうもう "合図がありません" を送ったか」を送信ログの件名で
 *   確かめているので、テストを1回押しただけで、その日の本物のお知らせが
 *   出なくなっていました。テストの件名に【テスト】を入れ、
 *   数に入れないようにして直しています。
 *
 * v1.4で足したもの：ご家族用の画面のための入り口（action=family）
 *   ご家族が、メールを待たずに「いまどうなっているか」を見られるようにします。
 *   ★ これは【読むだけ】の入り口です。ここから合図を記録することはできません。
 *     ご家族の端末から押せてしまうと、ご本人が押していない日でも「元気」と
 *     記録され、お知らせが止まってしまうためです。
 *   ★ 見張り（15分おきのトリガー）が生きているかも返します。見張りが止まると
 *     「合図がなくてもメールが来ない」という、いちばん危ない静かな故障になるので、
 *     ご家族の画面でいつでも確かめられるようにしました。
 *   ※ シートの形は変わりません。「① 初期設定」を押し直す必要はありません。
 */

var VERSION = 'v1.4';

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

  ['last_checkin', 'alert_date', 'clear_date', 'report_date', 'today_mood',
   'last_loc', 'last_loc_at']
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
    var subject = String(rows[i][2]);
    // 【テスト】と付いたものは数に入れない。テストしただけで
    // その日の本物のアラートが止まってしまうのを防ぐため。
    if (subject.indexOf('【テスト】') >= 0) continue;
    if (d === today && subject.indexOf('合図がありません') >= 0) return true;
  }
  return false;
}


/* ============================================================
   ボタン画面からの入り口（ウェブアプリ）
   ============================================================ */

// 状態を返す（画面を開いたとき）
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'status';
  if (action === 'checkin') {
    // POSTがうまくいかない環境のための逃げ道
    return json_(checkin_(p.mood || 'genki', { lat: p.lat, lon: p.lon, acc: p.acc }));
  }
  if (action === 'family') return json_(familyStatus_());   // ご家族用（読むだけ）
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
  if (data.action === 'family') return json_(familyStatus_());   // ご家族用（読むだけ）
  return json_(checkin_(data.mood || 'genki', data.loc));
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
    track_location: s.track_location,
    checked_today: (stateDate_('last_checkin') === today),
    today_mood: getState_('today_mood'),
    history: recentRows_(7)
  };
}

/**
 * ご家族用の状態（v1.4）。
 *
 * ★ ここは【読むだけ】です。記録もメール送信も一切しません。
 *   ご家族の端末から合図を送れてしまうと、ご本人が押していない日でも
 *   「元気」と記録され、お知らせが止まってしまいます。
 *
 * ボタン画面用の status_() と分けてあるのは、
 *   ・ご家族には、もっと厚い情報（見張りの生死・場所・14日分の記録）が要る
 *   ・ご本人の画面には、増やしたくない
 * という理由です。status_() は今までのまま変えていないので、
 * 古いボタン画面のままの方がいても、今までどおり動きます。
 */
function familyStatus_() {
  var s = getSettings_();
  var now = new Date();
  var today = ymd_(now);
  var nowHM = Utilities.formatDate(now, tz_(), 'HH:mm');

  return {
    ok: true,
    version: VERSION,
    name: s.name,
    deadline: s.deadline,
    daily_report: s.daily_report,
    track_location: s.track_location,
    now: Utilities.formatDate(now, tz_(), 'yyyy-MM-dd HH:mm'),
    // 締め切りを過ぎているか。日付と時刻でちがう時計を混ぜないよう、
    // checkDeadline と同じ「HH:mm の文字くらべ」で判断する。
    past_deadline: (nowHM >= s.deadline),
    checked_today: (stateDate_('last_checkin') === today),
    today_mood: getState_('today_mood'),
    last_checkin: getState_('last_checkin'),
    alerted_today: (stateDate_('alert_date') === today),
    cleared_today: (stateDate_('clear_date') === today),
    watching: watching_(),
    last_loc: lastLocInfo_(s),
    history: recentRows_(14)
  };
}

/**
 * 15分おきの見張りが仕掛けられているかを確かめる。
 * 見張りが止まると「合図がなくてもメールが来ない」という、
 * 外からは元気に見えるのに中身が死んでいる状態になります。
 * これを黙って起こさないための確認です。
 *
 * ※ ここで使う ScriptApp は setup() でも使っているので、
 *   すでにお使いの方に新しい許可を求めることはありません。
 * ※ 何かの理由で調べられなかったときは、false（止まっている）ではなく
 *   null（分からない）を返します。うその警告を出さないためです。
 */
function watching_() {
  try {
    var ts = ScriptApp.getProjectTriggers();
    for (var i = 0; i < ts.length; i++) {
      if (ts[i].getHandlerFunction() === 'checkDeadline') return true;
    }
    return false;
  } catch (err) {
    return null;
  }
}

/** 最後に合図があったときの場所。地図のリンクまで作って返す。 */
function lastLocInfo_(s) {
  if (!s.track_location) return null;
  var raw = String(getState_('last_loc') || '').split(',');
  var loc = cleanLoc_({ lat: raw[0], lon: raw[1], acc: raw[2] });
  if (!loc) return null;
  return {
    lat: loc.lat, lon: loc.lon, acc: loc.acc,
    at: String(getState_('last_loc_at') || ''),
    map: mapUrl_(loc)
  };
}

function checkin_(mood, rawLoc) {
  if (!MOODS[mood]) mood = 'genki';

  var s = getSettings_();
  var now = new Date();
  var today = ymd_(now);

  // 場所は「設定シートでするになっている」ときだけ受け取る。
  // 家族がシートで「しない」にすれば、押す人のスマホを触らなくても記録は止まる。
  var loc = s.track_location ? cleanLoc_(rawLoc) : null;

  setState_('last_checkin', Utilities.formatDate(now, tz_(), "yyyy-MM-dd'T'HH:mm:ss"));
  setState_('today_mood', mood);
  if (loc) {
    setState_('last_loc', loc.lat + ',' + loc.lon + ',' + loc.acc);
    setState_('last_loc_at', Utilities.formatDate(now, tz_(), 'yyyy-MM-dd HH:mm'));
  }
  addHistory_(mood, now, loc);

  // どのメールを出すかを決める（今までのPython版と同じ判断です）
  if (s.contact_email) {
    if (stateDate_('alert_date') === today && stateDate_('clear_date') !== today) {
      // すでに「合図がありません」を送ったあとに押された → 無事を伝える
      doAllClear_(s, now, loc);
      setState_('clear_date', today);
      setState_('report_date', today);   // この日の報告は済んだ扱い
    } else if (s.daily_report && stateDate_('report_date') !== today) {
      // その日はじめての合図 → 毎日の報告
      doDailyReport_(s, now, mood, loc);
      setState_('report_date', today);
    }
  }

  return { ok: true, time: Utilities.formatDate(now, tz_(), 'HH:mm'), mood: mood };
}


/* ============================================================
   メールの中身（Python版の文面をそのまま引き継いでいます）
   ============================================================ */
/**
 * 「合図がありません」のお知らせ。
 * isTest が true のときは、件名に【テスト】を入れて本物と見分けられるようにする。
 * これは見た目の問題ではなく、見守りが止まらないための作りです。
 * alertSentToday_ は送信ログの件名で「きょうもう送ったか」を判断するので、
 * テストと本物が同じ件名だと、テストした日は本物のアラートが出なくなってしまう。
 */
function doAlert_(s, isTest) {
  var name = nameOf_(s);
  sendMail_(
    '【みまもり】' + (isTest ? '【テスト】' : '') + name + 'さんから、きょうの元気の合図がありません',
    (isTest ? '※これはテスト送信です。実際に合図がなかったわけではありません。\n\n' : '')
      + name + 'さんから、本日の「元気です」の合図が、\n'
      + '締め切り時刻（' + s.deadline + '）までにありませんでした。\n\n'
      + '念のため、電話や訪問で様子を確認してください。\n'
      + lastLocBlock_(s)
      + '\n── 最近1週間の記録 ──\n'
      + recentLines_(7) + '\n\n'
      + '（この通知は みまもり ' + VERSION + ' から自動送信されています）\n'
  );
}

function doAllClear_(s, now, loc) {
  var name = nameOf_(s);
  sendMail_(
    '【みまもり】' + name + 'さんから、遅れて元気の合図がありました',
    '先ほど「合図がありません」とお知らせしましたが、その後\n'
      + Utilities.formatDate(now, tz_(), 'HH:mm') + ' に '
      + name + 'さんが「元気です」を押されました。\n\n'
      + '押し忘れだったようです。ひとまずご安心ください。\n'
      + locBlock_(s, loc, '押されたときの場所')
      + '（みまもり ' + VERSION + ' からの自動送信）\n'
  );
}

function doDailyReport_(s, now, mood, loc) {
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
    head
      + locBlock_(s, loc, '合図があった場所')
      + '── 最近1週間の記録 ──\n'
      + recentLines_(7) + '\n\n'
      + '※ 押した時刻が少しずつ遅くなっていくときは、\n'
      + '　 体調が落ちてきているサインのことがあります。\n\n'
      + '※このお知らせは毎日1回届きます。\n'
      + '　届かない日は、ご本人か見守りの仕組みに何かあった可能性があります。\n'
      + '（みまもり ' + VERSION + ' からの自動送信）\n'
  );
}

/* ============================================================
   合図があった場所（v1.2）
   ============================================================ */

/**
 * ボタン画面から届いた位置を、安全な形に整える。
 * ・数字として読めないもの、地球の外の値（緯度90より大きいなど）ははじく
 * ・小数点以下5桁（＝およそ1m）に丸める。それ以上細かくしても意味がない
 * 変な値をそのままシートやメールに入れないための関所です。
 */
function cleanLoc_(raw) {
  if (!raw) return null;
  var lat = Number(raw.lat), lon = Number(raw.lon), acc = Number(raw.acc);
  if (!isFinite(lat) || !isFinite(lon)) return null;
  if (lat < -90 || lat > 90 || lon < -180 || lon > 180) return null;
  if (lat === 0 && lon === 0) return null;   // 取れなかったときの典型的なゴミ
  return {
    lat: Math.round(lat * 100000) / 100000,
    lon: Math.round(lon * 100000) / 100000,
    acc: (isFinite(acc) && acc > 0) ? Math.round(acc) : ''
  };
}

/** 地図で開けるリンク。スマホならGoogleマップのアプリが開きます。 */
function mapUrl_(loc) {
  return 'https://www.google.com/maps?q=' + loc.lat + ',' + loc.lon;
}

/** メール本文に入れる「場所」のかたまり。 */
function locBlock_(s, loc, title) {
  if (!s.track_location) return '\n';
  if (!loc) {
    return '\n※ 場所は分かりませんでした。\n'
         + '　 スマホで位置情報が許可されていないか、電波が届かなかったようです。\n\n';
  }
  return '\n── ' + title + ' ──\n'
       + '  ' + loc.lat + ', ' + loc.lon
       + (loc.acc ? '（誤差 およそ ' + loc.acc + 'm）' : '') + '\n'
       + '  地図：' + mapUrl_(loc) + '\n\n';
}

/** 「合図がありません」のメール用。最後に押されたときの場所を出す。 */
function lastLocBlock_(s) {
  if (!s.track_location) return '\n';
  var raw = String(getState_('last_loc') || '').split(',');
  var loc = cleanLoc_({ lat: raw[0], lon: raw[1], acc: raw[2] });
  if (!loc) return '\n';

  var at = String(getState_('last_loc_at') || '');
  return '\n── 最後に合図があったときの場所 ──\n'
       + (at ? '  ' + at + '\n' : '')
       + '  ' + loc.lat + ', ' + loc.lon
       + (loc.acc ? '（誤差 およそ ' + loc.acc + 'm）' : '') + '\n'
       + '  地図：' + mapUrl_(loc) + '\n'
       + '  ※ これは【最後にボタンを押したときの場所】です。\n'
       + '　　 いま現在いる場所ではありません。\n\n';
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
  doAlert_(s, true);
  SpreadsheetApp.getUi().alert('テスト通知を ' + s.contact_email + ' に送りました。\n'
    + '受信箱を確認してください。\n\n'
    + '件名に【テスト】と入っています。\n'
    + 'この送信は、きょうの本物の見張りには影響しません。');
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
    + '締め切り時刻：' + st.deadline + '\n'
    + '場所の記録：' + (st.track_location ? 'する' : 'しない')
    + (st.track_location && getState_('last_loc')
        ? '（最後：' + getState_('last_loc_at') + '）' : '')
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

var HISTORY_HEADER = ['日付', '時刻', '体調', '緯度', '経度', '誤差(m)', '地図'];

function ensureSheets_() {
  var st = sheet_(SHEET_SETTINGS);
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 8, 2).setValues([
      ['項目', '値'],
      ['見守る方のお名前', ''],
      ['連絡先メール（家族）', ''],
      ['締め切り時刻', '10:00'],
      ['毎日の報告メール', 'する'],
      ['熱中症の警戒を出す', 'する'],
      ['お住まいの都道府県', '愛知県'],
      ['合図があった場所を記録する', 'する']
    ]);
    st.setColumnWidth(1, 200);
    st.setColumnWidth(2, 260);
    st.getRange('A1:B1').setFontWeight('bold');
  } else {
    // 前の版から使っている方のために、足りない項目だけ下に追加する。
    // すでにある行には触らないので、設定が消えることはありません。
    var have = {};
    st.getRange(1, 1, st.getLastRow(), 1).getValues()
      .forEach(function (r) { have[String(r[0]).trim()] = true; });
    if (!have['合図があった場所を記録する']) {
      st.appendRow(['合図があった場所を記録する', 'する']);
    }
  }

  var hi = sheet_(SHEET_HISTORY);
  if (hi.getLastRow() === 0) {
    hi.appendRow(HISTORY_HEADER);
  } else if (hi.getLastColumn() < HISTORY_HEADER.length) {
    // 3列だった「記録」シートを7列に広げる（すでにある中身はそのまま）
    if (hi.getMaxColumns() < HISTORY_HEADER.length) {
      hi.insertColumnsAfter(hi.getMaxColumns(), HISTORY_HEADER.length - hi.getMaxColumns());
    }
    hi.getRange(1, 1, 1, HISTORY_HEADER.length).setValues([HISTORY_HEADER]);
  }

  var ml = sheet_(SHEET_MAILLOG);
  if (ml.getLastRow() === 0) ml.appendRow(['日時', '宛先', '件名']);

  sheet_(SHEET_STATE);
}

function getSettings_() {
  var sh = ss_().getSheetByName(SHEET_SETTINGS);
  if (!sh || sh.getLastRow() < 2) {
    return { name: '', contact_email: '', deadline: '10:00',
             daily_report: true, heat_alert: true, region: '愛知県',
             track_location: true };
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
    region: String(v['お住まいの都道府県'] || '愛知県').trim(),
    track_location: String(v['合図があった場所を記録する'] || 'する').trim() !== 'しない'
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

/** その日はじめて押した時刻を残す（あとから上書きしない。起きて動き出した時刻の目安になるので）。
 *  場所は押しなおすたびに最新にする（そのほうが「いちばん新しく分かっている場所」になるため）。 */
function addHistory_(mood, now, loc) {
  var sh = sheet_(SHEET_HISTORY);
  var today = ymd_(now);
  var last = sh.getLastRow();

  if (last >= 2) {
    var lastDate = sh.getRange(last, 1).getValue();
    if (lastDate instanceof Date) lastDate = ymd_(lastDate);
    if (String(lastDate) === today) {
      sh.getRange(last, 3).setValue(MOODS[mood].short);  // 体調だけ最新に更新
      if (loc) sh.getRange(last, 4, 1, 4).setValues([locCells_(loc)]);
      return;
    }
  }
  sh.appendRow([today, Utilities.formatDate(now, tz_(), 'HH:mm'), MOODS[mood].short]
    .concat(loc ? locCells_(loc) : ['', '', '', '']));
}

/** 「記録」シートの右4列（緯度・経度・誤差・地図）ぶんの中身。 */
function locCells_(loc) {
  return [loc.lat, loc.lon, loc.acc, mapUrl_(loc)];
}

function recentRows_(days) {
  var sh = sheet_(SHEET_HISTORY);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var from = Math.max(2, last - days + 1);
  return sh.getRange(from, 1, last - from + 1, 3).getValues().map(function (r) {
    var d = r[0] instanceof Date ? ymd_(r[0]) : String(r[0]);
    // 時刻も日付と同じ用心をする。"07:12" と書き込んでも、スプレッドシートが
    // 時刻データに変換していることがあり、そのまま文字にすると
    // 「Sat Dec 30 1899 07:12:00 ...」という表示になってしまう。
    var t = r[1] instanceof Date ? Utilities.formatDate(r[1], tz_(), 'HH:mm') : String(r[1]);
    return { d: d, t: t, m: String(r[2]) };
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
