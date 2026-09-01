/*
  無線クエスト2
  Version 2.6
*/

// ===== シート定義 =====

const SHEET = {
  bank:  '問題バンク',
  inbox: '投稿箱',
  op:    'オペレーター',
  log:   '交信ログ',
  raid:  'レイド',
  conf:  '設定'
};

const HEAD = {
  bank:  ['問題ID', 'ジャンル', '難易度', '問題文', '選択肢A', '選択肢B', '選択肢C', '選択肢D', '正解', '解説', '作成者', '公開'],
  inbox: ['投稿日時', '投稿者', 'ジャンル', '難易度', '問題文', '選択肢A', '選択肢B', '選択肢C', '選択肢D', '正解', '解説', '審査'],
  op:    ['ハンドル', 'レベル', '経験値', '進行', '取得スキル', '撃破数', '正答数', '回答数', '投稿数', '最終交信',
          '電池', '週キー', '週経験値', '投稿報酬受取'],
  log:   ['日時', 'ハンドル', 'ジャンル', '問題ID', '誤答数', '所要秒', '結果'],
  raid:  ['週キー', 'ハンドル', 'あたえたダメージ', '最終更新'],
  conf:  ['項目', '値', '説明']
};

/** 週替わりレイドのボス。週ごとに順番に入れ替わる。 */
const RAID_BOSSES = [
  { name: '大電離圏クジラ', shape: 'wisp' },
  { name: 'ジャミングの王', shape: 'golem' },
  { name: '無許可電波の竜', shape: 'dragon' },
  { name: 'ノイズの海坊主', shape: 'blob' }
];

/**
 * ジャンルがそのまま冒険のエリアになる。
 * 並び順が到達順で、ひとつ前のエリアのボスを倒すと次のエリアが開く。
 */
const AREAS = [
  { key: '電波の基礎',   boss: '電離層ドラゴン', skill: '伝搬解析',   shape: 'dragon',  hue: 205 },
  { key: '運用と通信術', boss: '混信の亡霊',     skill: '交信術',     shape: 'ghost',   hue: 275 },
  { key: 'アンテナ',     boss: '定在波のぬし',   skill: '空中線調整', shape: 'tower',   hue: 150 },
  { key: '無線工学',     boss: '発振の魔人',     skill: '回路設計',   shape: 'crystal', hue: 35  },
  { key: 'IoT無線とLTE-M', boss: 'テール電流の魔', skill: '低消費電力設計', shape: 'wisp', hue: 320 },
  { key: '電波法規',     boss: '条文の守護者',   skill: '法令遵守',   shape: 'golem',   hue: 0   }
];

/**
 * エリアを持たないジャンル。息抜きとして、どのエリアの戦闘にも混ざって出てきます。
 * ここに名前を足せば、問題シート側でも同じ名前が使えるようになります。
 */
const BONUS_GENRES = ['いっぱんクイズ', 'なぞなぞ'];

/** 1エリアで集めるアイテムの数。3つそろうとボスに挑める。 */
const ITEMS_PER_AREA = 3;

/**
 * 進行はビットで持つ。1つめを取ったら 1、2つめで 2、3つめで 4。
 * 3つそろうと 7 になる。どのアイテムを取ったかまで1つの数で覚えられる。
 */
const PROGRESS_FULL = (1 << ITEMS_PER_AREA) - 1;

/** 道中のモンスター。難易度に応じて選ばれる。 */
const MOBS = [
  { name: 'ノイズスライム', shape: 'blob',    tier: 1 },
  { name: 'ハムのおばけ',   shape: 'ghost',   tier: 1 },
  { name: 'フェージング',   shape: 'wisp',    tier: 1 },
  { name: '高調波コウモリ', shape: 'bat',     tier: 2 },
  { name: 'SWRゴースト',    shape: 'wisp',    tier: 2 },
  { name: '寄生発振むし',   shape: 'blob',    tier: 2 },
  { name: 'スプリアス',     shape: 'crystal', tier: 3 },
  { name: 'パイルアップ',   shape: 'golem',   tier: 3 },
  { name: '不法電波の影',   shape: 'bat',     tier: 3 }
];

const DEFAULT_CONF = [
  ['投稿経験値', '120', '問題を1問投稿したときにもらえる経験値'],
  ['投稿の自動承認', 'OFF', 'ON にすると投稿がすぐ出題に混ざります。OFF なら「投稿箱」で承認するまで出ません'],
  ['1日の投稿上限', '5', '1人が1日に投稿できる問題数。経験値めあての連投を防ぎます'],
  ['レベルアップ経験値', '50', 'レベルが1つ上がるのに必要な経験値'],
  ['制限時間', '20', '1問に答えられる秒数。早く答えるほど相手に大きなダメージを与えます'],
  ['週間レイドHP', '4000', 'みんなで削る週替わりボスのHP。人数が多いほど大きくします'],
  ['投稿ボーナス', '2', '自分の投稿問題が1回出題されるごとにもらえる経験値'],
  ['電池の上限', '3', '持ち歩ける予備バッテリーの数。戦闘中にHPを回復できます'],
  ['ランキング表示人数', '20', '「みんなの状況」に出す人数']
];

// ===== ウェブアプリ入口 =====

function doGet() {
  return HtmlService.createTemplateFromFile('musen').evaluate()
    .setTitle('無線クエスト2')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** スプレッドシートを開いたときに出る運営用メニュー。 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('無線クエスト2')
    .addItem('シートを準備する', 'setupSheets')
    .addItem('選択した投稿を承認する', 'approveSelected')
    .addItem('選択した投稿を却下する', 'rejectSelected')
    .addToUi();
}

// ===== セットアップ =====

/**
 * 5枚のシートを用意して、初期の問題を入れる。
 * すでに中身がある場合は消さず、足りないものだけ足す。
 */
function setupSheets() {
  const book = SpreadsheetApp.getActiveSpreadsheet();

  const bank = prepareSheet_(book, SHEET.bank, HEAD.bank);
  if (bank.getLastRow() <= 1) {
    const seed = seedQuestions_();
    bank.getRange(2, 1, seed.length, HEAD.bank.length).setValues(seed);
    bank.setColumnWidth(4, 360);
    bank.setColumnWidth(10, 360);
  }

  prepareSheet_(book, SHEET.inbox, HEAD.inbox).setColumnWidth(5, 360);
  prepareSheet_(book, SHEET.op, HEAD.op);
  prepareSheet_(book, SHEET.log, HEAD.log);
  prepareSheet_(book, SHEET.raid, HEAD.raid);

  const conf = prepareSheet_(book, SHEET.conf, HEAD.conf);
  if (conf.getLastRow() <= 1) {
    conf.getRange(2, 1, DEFAULT_CONF.length, HEAD.conf.length).setValues(DEFAULT_CONF);
    conf.setColumnWidth(3, 420);
  }

  const blank = book.getSheetByName('シート1') || book.getSheetByName('Sheet1');
  if (blank && book.getSheets().length > 1 && blank.getLastRow() === 0) book.deleteSheet(blank);

  return '準備できました';
}

function prepareSheet_(book, name, head) {
  let sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  const top = sheet.getRange(1, 1, 1, head.length).getValues()[0];
  // 後ろに列が増えたときは見出しを書き足す。既存の記録はそのまま残る
  const short = top[0] === head[0] && top[head.length - 1] === '';
  if (top.join('') === '' || top[0] !== head[0] || short) {
    sheet.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#dbeafe');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 誰かが最初にURLを開いた時点でシートを用意する。
 * 運営がメニューを触らなくても遊び始められるようにするための保険。
 * 同時アクセスで初期問題が二重に入らないよう、準備が要るときだけ鍵を取る。
 */
function warmUp_(book) {
  const bank = book.getSheetByName(SHEET.bank);
  if (bank && bank.getLastRow() > 1) return;

  const gate = LockService.getScriptLock();
  if (!gate.tryLock(20000)) return;
  try {
    const retry = book.getSheetByName(SHEET.bank);
    if (retry && retry.getLastRow() > 1) return;
    setupSheets();
  } catch (err) {
    Logger.log('準備に失敗: ' + err.message);
  } finally {
    gate.releaseLock();
  }
}

// ===== クライアントから呼ばれる関数 =====

/**
 * 起動時に1回だけ呼ぶ。問題を全部と、その人の進行と、みんなの状況をまとめて返す。
 * 以後は1問ごとに通信しないので、大人数が同時に遊んでも詰まりにくい。
 */
function bootstrap(handle) {
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    warmUp_(book);
    const conf = readConf_(book);
    const questions = collectQuestions_(book);

    if (!questions.length) {
      return { ok: false, reason: '出題できる問題がありません。「問題バンク」を確認してください。' };
    }

    return {
      ok: true,
      areas: AREAS,
      bonusGenres: BONUS_GENRES,
      maps: AREA_MAPS,
      itemsPerArea: ITEMS_PER_AREA,
      mobs: MOBS,
      questions: questions,
      me: handle ? loadOperator_(book, handle) : null,
      roster: readRoster_(book, conf),
      raid: getRaid(handle),
      rules: {
        expPerLevel: toRange_(conf['レベルアップ経験値'], 1, 9999, 50),
        timeLimit: toRange_(conf['制限時間'], 5, 300, 20),
        cellCap: toRange_(conf['電池の上限'], 0, 9, 3),
        postExp: toRange_(conf['投稿経験値'], 0, 9999, 120),
        postLimit: toRange_(conf['1日の投稿上限'], 0, 99, 5)
      }
    };
  } catch (err) {
    return { ok: false, reason: '読み込みに失敗しました: ' + err.message };
  }
}

/** ハンドルを決めたあと、その人の続きだけを取りに行く軽い呼び出し。 */
function signIn(handle) {
  try {
    const name = cleanText_(handle, 16);
    if (!name) return { ok: false, reason: 'ハンドルを入れてください' };
    const book = SpreadsheetApp.getActiveSpreadsheet();
    return { ok: true, me: loadOperator_(book, name) };
  } catch (err) {
    return { ok: false, reason: err.message };
  }
}

/**
 * 冒険の区切り（モンスター撃破・ボス撃破・基地局にもどったとき）でまとめて保存する。
 * 1問ごとには送らない。
 */
function syncProgress(payload) {
  const gate = LockService.getScriptLock();
  if (!gate.tryLock(10000)) return { ok: false, retry: true, reason: '混み合っています' };
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const name = cleanText_(payload && payload.handle, 16);
    if (!name) return { ok: false, reason: 'ハンドルが未設定です' };

    const lines = (payload && payload.log) || [];
    if (lines.length > 150) return { ok: false, retry: true, reason: '一度に送れるのは150件までです' };

    if (lines.length) {
      const log = book.getSheetByName(SHEET.log) || prepareSheet_(book, SHEET.log, HEAD.log);
      const rows = lines.map(function (line) {
        return [
          new Date(),
          name,
          cleanText_(line.area, 30),
          cleanText_(line.questionId, 20),
          Math.max(0, Number(line.misses) || 0),
          Math.round((Number(line.seconds) || 0) * 10) / 10,
          line.cleared ? '撃破' : '交戦'
        ];
      });
      log.getRange(log.getLastRow() + 1, 1, rows.length, HEAD.log.length).setValues(rows);
    }

    return { ok: true, me: writeOperator_(book, name, payload) };
  } catch (err) {
    return { ok: false, retry: true, reason: '保存に失敗しました: ' + err.message };
  } finally {
    gate.releaseLock();
  }
}

/** みんなのレベルを取り直す。ランキング画面を開いたときだけ呼ぶ。 */
function fetchRoster() {
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    return { ok: true, roster: readRoster_(book, readConf_(book)) };
  } catch (err) {
    return { ok: false, roster: [] };
  }
}

/**
 * 問題を投稿する。投稿そのものに大きな経験値を出すのが狙いなので、
 * 承認を待たずにその場で経験値を渡す。出題に混ざるかどうかは承認しだい。
 */
function submitQuestion(draft) {
  const gate = LockService.getScriptLock();
  if (!gate.tryLock(10000)) return { ok: false, retry: true, reason: '混み合っています' };
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const conf = readConf_(book);
    const name = cleanText_(draft && draft.handle, 16);
    if (!name) return { ok: false, reason: 'ハンドルが未設定です' };

    const area = cleanText_(draft && draft.area, 30);
    const knownGenre = AREAS.some(function (a) { return a.key === area; }) ||
                       BONUS_GENRES.indexOf(area) !== -1;
    if (!knownGenre) return { ok: false, reason: 'ジャンルを選んでください' };

    const level = toRange_(draft && draft.level, 1, 3, 1);
    const text = cleanText_(draft && draft.text, 300);
    if (text.length < 8) return { ok: false, reason: '問題文が短すぎます（8文字以上）' };

    const picks = ['a', 'b', 'c', 'd'].map(function (key) {
      return cleanText_(draft && draft.choices && draft.choices[key], 120);
    });
    if (picks.some(function (choice) { return !choice; })) {
      return { ok: false, reason: '選択肢を4つとも埋めてください' };
    }
    if (new Set(picks).size < 4) return { ok: false, reason: '同じ選択肢が混ざっています' };

    const answer = String((draft && draft.answer) || '').toUpperCase();
    if (answer.length !== 1 || 'ABCD'.indexOf(answer) === -1) {
      return { ok: false, reason: '正解を選んでください' };
    }

    const inbox = book.getSheetByName(SHEET.inbox) || prepareSheet_(book, SHEET.inbox, HEAD.inbox);
    const stop = checkPostGuard_(book, inbox, name, text, conf);
    if (stop) return { ok: false, reason: stop };

    const autoPass = String(conf['投稿の自動承認'] || 'OFF').trim().toUpperCase() === 'ON';
    inbox.appendRow([
      new Date(), name, area, level, text,
      picks[0], picks[1], picks[2], picks[3], answer,
      cleanText_(draft && draft.note, 300),
      autoPass ? '承認' : '審査中'
    ]);

    const reward = toRange_(conf['投稿経験値'], 0, 9999, 120);
    return {
      ok: true,
      gained: reward,
      published: autoPass,
      me: addPostReward_(book, name, reward, conf)
    };
  } catch (err) {
    return { ok: false, retry: true, reason: '投稿に失敗しました: ' + err.message };
  } finally {
    gate.releaseLock();
  }
}

/** 連投と重複を止める。問題が増えないまま経験値だけ増えるのを防ぐため。 */
function checkPostGuard_(book, inbox, name, text, conf) {
  const limit = toRange_(conf['1日の投稿上限'], 0, 99, 5);
  const last = inbox.getLastRow();

  if (last > 1) {
    const rows = inbox.getRange(2, 1, last - 1, 5).getValues();
    const midnight = new Date();
    midnight.setHours(0, 0, 0, 0);

    let today = 0;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][4]).trim() === text) return '同じ問題文がすでに投稿されています';
      if (String(rows[i][1]).trim() !== name) continue;
      if (rows[i][0] instanceof Date && rows[i][0] >= midnight) today++;
    }
    if (limit && today >= limit) return '今日の投稿はここまでです（1日 ' + limit + ' 問まで）';
  }

  const bank = book.getSheetByName(SHEET.bank);
  if (bank && bank.getLastRow() > 1) {
    const known = bank.getRange(2, 4, bank.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < known.length; i++) {
      if (String(known[i][0]).trim() === text) return '同じ問題文がすでに出題されています';
    }
  }
  return '';
}

// ===== 運営メニュー =====

function approveSelected() { stampInbox_('承認'); }
function rejectSelected()  { stampInbox_('却下'); }

function stampInbox_(mark) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET.inbox) {
    ui.alert('「' + SHEET.inbox + '」シートで、対象の行を選んでから実行してください。');
    return;
  }

  const picked = sheet.getActiveRange();
  const from = Math.max(2, picked.getRow());
  const count = picked.getRow() + picked.getNumRows() - from;
  if (count < 1) { ui.alert('見出し行より下のデータ行を選んでください。'); return; }

  sheet.getRange(from, HEAD.inbox.indexOf('審査') + 1, count, 1).setValue(mark);
  ui.alert(count + ' 件を「' + mark + '」にしました。');
}

// ===== 問題の読み出し =====

/**
 * 「問題バンク」の公開ぶんと、「投稿箱」の承認ぶんを合わせて出題対象にする。
 * 承認した投稿はコピーせずそのまま出すので、運営の手間は承認だけで済む。
 */
function collectQuestions_(book) {
  const out = [];
  const seen = {};

  const bank = book.getSheetByName(SHEET.bank);
  if (bank && bank.getLastRow() > 1) {
    const rows = bank.getRange(2, 1, bank.getLastRow() - 1, HEAD.bank.length).getValues();
    rows.forEach(function (row, i) {
      if (!isOn_(row[11])) return;
      const q = shapeQuestion_(String(row[0] || ('B' + (i + 1))), row[1], row[2], row[3],
                               [row[4], row[5], row[6], row[7]], row[8], row[9], row[10]);
      if (q && !seen[q.text]) { seen[q.text] = true; out.push(q); }
    });
  }

  const inbox = book.getSheetByName(SHEET.inbox);
  if (inbox && inbox.getLastRow() > 1) {
    const rows = inbox.getRange(2, 1, inbox.getLastRow() - 1, HEAD.inbox.length).getValues();
    rows.forEach(function (row, i) {
      if (String(row[11] || '').trim() !== '承認') return;
      const q = shapeQuestion_('P' + (i + 1), row[2], row[3], row[4],
                               [row[5], row[6], row[7], row[8]], row[9], row[10], row[1]);
      if (q && !seen[q.text]) { seen[q.text] = true; q.fromPlayer = true; out.push(q); }
    });
  }

  return out;
}

/** 1行を出題できる形に整える。出題できない行は null を返して捨てる。 */
function shapeQuestion_(id, area, level, text, rawChoices, answer, note, author) {
  const body = cleanText_(text, 300);
  if (!body) return null;

  const genre = cleanText_(area, 30);
  const known = AREAS.some(function (a) { return a.key === genre; }) ||
                BONUS_GENRES.indexOf(genre) !== -1;
  if (!known) return null;

  const choices = rawChoices.map(function (choice) { return cleanText_(choice, 120); });
  if (choices.some(function (choice) { return !choice; })) return null;

  const at = 'ABCD'.indexOf(String(answer == null ? '' : answer).trim().toUpperCase());
  if (at === -1) return null;

  return {
    id: cleanText_(id, 20),
    area: genre,
    level: toRange_(level, 1, 3, 1),
    text: body,
    choices: choices,
    answer: at,
    note: cleanText_(note, 300),
    author: cleanText_(author, 16)
  };
}

// ===== オペレーターの読み書き =====

function loadOperator_(book, handle) {
  const name = cleanText_(handle, 16);
  const blank = {
    handle: name, level: 1, exp: 0, progress: {}, skills: [],
    defeats: 0, hits: 0, tries: 0, posts: 0,
    cells: 0, weekKey: weekKey_(), weekExp: 0, authorPaid: 0
  };
  const sheet = book.getSheetByName(SHEET.op);
  if (!sheet || sheet.getLastRow() <= 1) return blank;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEAD.op.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== name) continue;
    return {
      handle: name,
      level: Math.max(1, Number(rows[i][1]) || 1),
      exp: Math.max(0, Number(rows[i][2]) || 0),
      progress: parseProgress_(rows[i][3]),
      skills: splitList_(rows[i][4]),
      defeats: Number(rows[i][5]) || 0,
      hits: Number(rows[i][6]) || 0,
      tries: Number(rows[i][7]) || 0,
      posts: Number(rows[i][8]) || 0,
      cells: Number(rows[i][10]) || 0,
      weekKey: String(rows[i][11] || '').trim(),
      weekExp: Number(rows[i][12]) || 0,
      authorPaid: Number(rows[i][13]) || 0
    };
  }
  return blank;
}

function writeOperator_(book, name, payload) {
  const sheet = book.getSheetByName(SHEET.op) || prepareSheet_(book, SHEET.op, HEAD.op);
  const current = loadOperator_(book, name);

  // 経験値や進行は減らさない。通信の行き違いで巻き戻るのを防ぐ
  const nextExp = Math.max(current.exp, Math.max(0, Number(payload.exp) || 0));
  const thisWeek = weekKey_();
  // 週が変わったら週間の記録だけ0に戻す。通算の経験値はそのまま
  const carried = current.weekKey === thisWeek ? current.weekExp : 0;

  const merged = {
    handle: name,
    level: Math.max(current.level, Math.max(1, Number(payload.level) || 1)),
    exp: nextExp,
    progress: mergeProgress_(current.progress, payload.progress),
    skills: mergeList_(current.skills, payload.skills),
    defeats: Math.max(current.defeats, Number(payload.defeats) || 0),
    hits: Math.max(current.hits, Number(payload.hits) || 0),
    tries: Math.max(current.tries, Number(payload.tries) || 0),
    posts: current.posts,
    cells: clampCells_(book, payload.cells, current.cells),
    weekKey: thisWeek,
    weekExp: carried + Math.max(0, nextExp - current.exp),
    authorPaid: current.authorPaid
  };

  saveOperatorRow_(sheet, merged);
  return merged;
}

function addPostReward_(book, name, reward, conf) {
  const sheet = book.getSheetByName(SHEET.op) || prepareSheet_(book, SHEET.op, HEAD.op);
  const me = loadOperator_(book, name);
  const thisWeek = weekKey_();
  me.weekExp = (me.weekKey === thisWeek ? me.weekExp : 0) + reward;
  me.weekKey = thisWeek;
  me.exp += reward;
  me.posts += 1;
  me.level = levelFor_(me.exp, toRange_(conf['レベルアップ経験値'], 1, 9999, 50));
  saveOperatorRow_(sheet, me);
  return me;
}

function saveOperatorRow_(sheet, me) {
  const line = [
    me.handle, me.level, me.exp, formatProgress_(me.progress), me.skills.join(','),
    me.defeats, me.hits, me.tries, me.posts, new Date(),
    me.cells || 0, me.weekKey || weekKey_(), me.weekExp || 0, me.authorPaid || 0
  ];
  const last = sheet.getLastRow();
  if (last > 1) {
    const names = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === me.handle) {
        sheet.getRange(i + 2, 1, 1, HEAD.op.length).setValues([line]);
        return;
      }
    }
  }
  sheet.appendRow(line);
}

/** 誰がどこまで進んでいるかの一覧。経験値の多い順。 */
function readRoster_(book, conf) {
  const sheet = book.getSheetByName(SHEET.op);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEAD.op.length).getValues();
  const list = rows.filter(function (row) {
    return String(row[0]).trim();
  }).map(function (row) {
    const skills = splitList_(row[4]);
    const week = weekKey_();
    return {
      handle: String(row[0]).trim(),
      level: Math.max(1, Number(row[1]) || 1),
      exp: Math.max(0, Number(row[2]) || 0),
      skills: skills,
      posts: Number(row[8]) || 0,
      seenAt: row[9] instanceof Date ? row[9].getTime() : 0,
      weekExp: String(row[11] || '').trim() === week ? (Number(row[12]) || 0) : 0
    };
  });

  list.sort(function (x, y) { return y.exp - x.exp || y.level - x.level; });
  return list.slice(0, toRange_(conf['ランキング表示人数'], 1, 200, 20));
}

// ===== 交信ログの集計 =====

/**
 * 交信ログを読む。件数が増えても重くならないよう、新しいほうから上限ぶんだけ見る。
 */
function recentLog_(book, limit) {
  const sheet = book.getSheetByName(SHEET.log);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const total = sheet.getLastRow() - 1;
  const take = Math.min(total, limit || 4000);
  const from = 2 + (total - take);
  return sheet.getRange(from, 1, take, HEAD.log.length).getValues();
}

/**
 * その人がどのジャンルでつまずいているかを出す。
 * 誤答数と所要秒の平均を見て、苦手な順に並べる。
 */
function getInsights(handle) {
  try {
    const name = cleanText_(handle, 16);
    if (!name) return { ok: false, rows: [] };

    const book = SpreadsheetApp.getActiveSpreadsheet();
    const box = {};
    recentLog_(book, 4000).forEach(function (row) {
      if (String(row[1]).trim() !== name) return;
      const area = String(row[2] || '').trim();
      if (!area) return;
      if (!box[area]) box[area] = { area: area, asked: 0, clean: 0, misses: 0, seconds: 0 };
      const cell = box[area];
      const miss = Number(row[4]) || 0;
      cell.asked++;
      cell.misses += miss;
      cell.seconds += Number(row[5]) || 0;
      if (!miss) cell.clean++;
    });

    const rows = Object.keys(box).map(function (key) {
      const cell = box[key];
      return {
        area: cell.area,
        asked: cell.asked,
        cleanRate: Math.round((cell.clean / cell.asked) * 100),
        avgMiss: Math.round((cell.misses / cell.asked) * 100) / 100,
        avgSeconds: Math.round((cell.seconds / cell.asked) * 10) / 10
      };
    });
    rows.sort(function (x, y) { return x.cleanRate - y.cleanRate; });
    return { ok: true, rows: rows };
  } catch (err) {
    return { ok: false, rows: [], reason: err.message };
  }
}

/**
 * 自分が投稿した問題が何回出題され、どれくらい正解されたかを返す。
 * まだ受け取っていない出題ぶんの経験値も、ここで渡す。
 */
function getMyQuestions(handle) {
  const gate = LockService.getScriptLock();
  if (!gate.tryLock(8000)) return { ok: false, reason: '混み合っています', rows: [] };
  try {
    const name = cleanText_(handle, 16);
    if (!name) return { ok: false, reason: 'ハンドルが未設定です', rows: [] };

    const book = SpreadsheetApp.getActiveSpreadsheet();
    const conf = readConf_(book);
    const mine = {};

    const bank = book.getSheetByName(SHEET.bank);
    if (bank && bank.getLastRow() > 1) {
      bank.getRange(2, 1, bank.getLastRow() - 1, HEAD.bank.length).getValues().forEach(function (row, i) {
        if (String(row[10] || '').trim() !== name) return;
        mine[String(row[0] || ('B' + (i + 1)))] = { id: String(row[0]), text: String(row[3]), state: '出題中' };
      });
    }

    const inbox = book.getSheetByName(SHEET.inbox);
    if (inbox && inbox.getLastRow() > 1) {
      inbox.getRange(2, 1, inbox.getLastRow() - 1, HEAD.inbox.length).getValues().forEach(function (row, i) {
        if (String(row[1] || '').trim() !== name) return;
        const state = String(row[11] || '').trim();
        mine['P' + (i + 1)] = { id: 'P' + (i + 1), text: String(row[4]), state: state };
      });
    }

    const stat = {};
    recentLog_(book, 4000).forEach(function (row) {
      const id = String(row[3] || '').trim();
      if (!mine[id]) return;
      if (!stat[id]) stat[id] = { asked: 0, clean: 0 };
      stat[id].asked++;
      if (!(Number(row[4]) || 0)) stat[id].clean++;
    });

    let served = 0;
    const rows = Object.keys(mine).map(function (id) {
      const hit = stat[id] || { asked: 0, clean: 0 };
      served += hit.asked;
      return {
        id: id,
        text: mine[id].text,
        state: mine[id].state,
        asked: hit.asked,
        cleanRate: hit.asked ? Math.round((hit.clean / hit.asked) * 100) : null
      };
    });
    rows.sort(function (x, y) { return y.asked - x.asked; });

    // まだ払っていない出題ぶんだけ、あとから経験値を渡す
    const perAsk = toRange_(conf['投稿ボーナス'], 0, 99, 2);
    const me = loadOperator_(book, name);
    const owed = Math.max(0, served - me.authorPaid) * perAsk;
    let paid = null;
    if (owed > 0) {
      const sheet = book.getSheetByName(SHEET.op) || prepareSheet_(book, SHEET.op, HEAD.op);
      const thisWeek = weekKey_();
      me.weekExp = (me.weekKey === thisWeek ? me.weekExp : 0) + owed;
      me.weekKey = thisWeek;
      me.exp += owed;
      me.authorPaid = served;
      me.level = levelFor_(me.exp, toRange_(conf['レベルアップ経験値'], 1, 9999, 50));
      saveOperatorRow_(sheet, me);
      paid = me;
    }

    return { ok: true, rows: rows, served: served, gained: owed, me: paid };
  } catch (err) {
    return { ok: false, reason: err.message, rows: [] };
  } finally {
    gate.releaseLock();
  }
}

// ===== 週替わりの協力レイド =====

/** 週ごとに変わる合言葉。これが変わると週間の記録とレイドが仕切り直しになる。 */
function weekKey_() {
  const now = new Date();
  const thursday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  // ISOと同じ数え方。木曜日を含む週をその年の第n週とする
  thursday.setDate(thursday.getDate() + 3 - ((thursday.getDay() + 6) % 7));
  const first = new Date(thursday.getFullYear(), 0, 4);
  const week = 1 + Math.round(
    ((thursday.getTime() - first.getTime()) / 86400000 - 3 + ((first.getDay() + 6) % 7)) / 7);
  return thursday.getFullYear() + '-W' + (week < 10 ? '0' + week : week);
}

function raidBossFor_(week) {
  let sum = 0;
  for (let i = 0; i < week.length; i++) sum += week.charCodeAt(i);
  return RAID_BOSSES[sum % RAID_BOSSES.length];
}

/** 今週のレイドの様子。残りHPと、貢献した人の一覧を返す。 */
function getRaid(handle) {
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const conf = readConf_(book);
    const week = weekKey_();
    const maxHp = toRange_(conf['週間レイドHP'], 100, 999999, 4000);
    const name = cleanText_(handle, 16);

    const sheet = book.getSheetByName(SHEET.raid);
    const board = {};
    let dealt = 0;
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, HEAD.raid.length).getValues().forEach(function (row) {
        if (String(row[0] || '').trim() !== week) return;
        const who = String(row[1] || '').trim();
        const hit = Math.max(0, Number(row[2]) || 0);
        if (!who) return;
        board[who] = (board[who] || 0) + hit;
        dealt += hit;
      });
    }

    const list = Object.keys(board).map(function (who) {
      return { handle: who, damage: board[who] };
    }).sort(function (x, y) { return y.damage - x.damage; });

    const boss = raidBossFor_(week);
    return {
      ok: true,
      week: week,
      name: boss.name,
      shape: boss.shape,
      maxHp: maxHp,
      hp: Math.max(0, maxHp - dealt),
      down: dealt >= maxHp,
      board: list.slice(0, 20),
      mine: name ? (board[name] || 0) : 0
    };
  } catch (err) {
    return { ok: false, reason: err.message };
  }
}

/** レイドに与えたダメージを持ち寄る。1回の申告に上限を置く。 */
function hitRaid(payload) {
  const gate = LockService.getScriptLock();
  if (!gate.tryLock(10000)) return { ok: false, retry: true, reason: '混み合っています' };
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const name = cleanText_(payload && payload.handle, 16);
    if (!name) return { ok: false, reason: 'ハンドルが未設定です' };

    const damage = Math.max(0, Math.min(2000, Math.round(Number(payload && payload.damage) || 0)));
    if (damage > 0) {
      const sheet = book.getSheetByName(SHEET.raid) || prepareSheet_(book, SHEET.raid, HEAD.raid);
      sheet.appendRow([weekKey_(), name, damage, new Date()]);
    }
    return { ok: true, raid: getRaid(name) };
  } catch (err) {
    return { ok: false, retry: true, reason: err.message };
  } finally {
    gate.releaseLock();
  }
}

/** 電池の数を、設定した上限のなかに収める。 */
function clampCells_(book, raw, fallback) {
  const conf = readConf_(book);
  const cap = toRange_(conf['電池の上限'], 0, 9, 3);
  const n = Number(raw);
  if (!isFinite(n)) return Math.min(fallback || 0, cap);
  return Math.max(0, Math.min(cap, Math.floor(n)));
}

// ===== 進行状況の入れ物 =====
// 「電波の基礎:4,アンテナ:6」のような1行の文字列にして1セルに収める。

function parseProgress_(raw) {
  const out = {};
  String(raw == null ? '' : raw).split(',').forEach(function (part) {
    const at = part.lastIndexOf(':');
    if (at < 1) return;
    const key = part.slice(0, at).trim();
    const step = Number(part.slice(at + 1));
    if (key && isFinite(step)) out[key] = Math.max(0, Math.min(PROGRESS_FULL, Math.floor(step)));
  });
  return out;
}

function formatProgress_(map) {
  return AREAS.filter(function (area) { return map[area.key]; })
    .map(function (area) { return area.key + ':' + map[area.key]; })
    .join(',');
}

function mergeProgress_(current, incoming) {
  const safe = (incoming && typeof incoming === 'object') ? incoming : {};
  const fresh = parseProgress_(formatProgress_(safe));
  const next = {};
  AREAS.forEach(function (area) {
    next[area.key] = Math.max(current[area.key] || 0, fresh[area.key] || 0);
  });
  return next;
}

// ===== 小道具 =====

function readConf_(book) {
  const conf = {};
  DEFAULT_CONF.forEach(function (row) { conf[row[0]] = row[1]; });

  const sheet = book.getSheetByName(SHEET.conf);
  if (!sheet || sheet.getLastRow() <= 1) return conf;

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (row) {
    const key = String(row[0] || '').trim();
    if (key) conf[key] = row[1];
  });
  return conf;
}

function levelFor_(exp, perLevel) {
  return Math.max(1, Math.floor(Math.max(0, exp) / perLevel) + 1);
}

function toRange_(raw, low, high, fallback) {
  const n = Number(raw);
  if (!isFinite(n) || n < low || n > high) return fallback;
  return Math.floor(n);
}

function isOn_(raw) {
  if (raw === true) return true;
  const s = String(raw == null ? '' : raw).trim().toUpperCase();
  return s === 'TRUE' || s === 'ON' || s === '○' || s === '公開' || s === '1';
}

function splitList_(raw) {
  return String(raw == null ? '' : raw).split(',')
    .map(function (s) { return s.trim(); })
    .filter(String);
}

function mergeList_(current, incoming) {
  const box = {};
  const extra = Array.isArray(incoming) ? incoming : [];
  current.concat(extra).forEach(function (s) {
    const name = String(s == null ? '' : s).trim();
    if (name) box[name] = true;
  });
  return Object.keys(box);
}

/**
 * シートに書く前に文字列を安全にする。
 * 先頭が = + - @ だと数式として解釈されてしまうので、頭に ' を足して打ち消す。
 */
function cleanText_(value, limit) {
  let s = String(value == null ? '' : value).replace(/[\x00-\x1f\x7f]/g, '').trim();
  if (limit) s = s.slice(0, limit);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// ===== 初期の問題 =====
// 初期投入する問題は questions.gs にあります。
// GAS は .gs ファイルをまとめて読み込むので、別ファイルでもそのまま呼べます。
