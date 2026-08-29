/*
  無線クエスト
  Version 1.0
*/

// ===== シート定義 =====

const SHEET = {
  bank:  '問題バンク',
  inbox: '投稿箱',
  op:    'オペレーター',
  log:   '交信ログ',
  conf:  '設定'
};

const HEAD = {
  bank:  ['問題ID', 'ジャンル', '難易度', '問題文', '選択肢A', '選択肢B', '選択肢C', '選択肢D', '正解', '解説', '作成者', '公開'],
  inbox: ['投稿日時', '投稿者', 'ジャンル', '難易度', '問題文', '選択肢A', '選択肢B', '選択肢C', '選択肢D', '正解', '解説', '審査'],
  op:    ['ハンドル', 'レベル', '経験値', '進行', '取得スキル', '撃破数', '正答数', '回答数', '投稿数', '最終交信'],
  log:   ['日時', 'ハンドル', 'ジャンル', '問題ID', '誤答数', '所要秒', '結果'],
  conf:  ['項目', '値', '説明']
};

/**
 * ジャンルがそのまま冒険のエリアになる。
 * 並び順が到達順で、ひとつ前のエリアのボスを倒すと次のエリアが開く。
 */
const AREAS = [
  { key: '電波の基礎',   boss: '電離層ドラゴン', skill: '伝搬解析',   shape: 'dragon',  hue: 205 },
  { key: '運用と通信術', boss: '混信の亡霊',     skill: '交信術',     shape: 'ghost',   hue: 275 },
  { key: 'アンテナ',     boss: '定在波のぬし',   skill: '空中線調整', shape: 'tower',   hue: 150 },
  { key: '無線工学',     boss: '発振の魔人',     skill: '回路設計',   shape: 'crystal', hue: 35  },
  { key: '電波法規',     boss: '条文の守護者',   skill: '法令遵守',   shape: 'golem',   hue: 0   }
];

/** 1エリアあたりの道のり。最後のノードがボス。 */
const NODES_PER_AREA = 6;

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
  ['ランキング表示人数', '20', '「みんなの状況」に出す人数']
];

// ===== ウェブアプリ入口 =====

function doGet() {
  return HtmlService.createTemplateFromFile('musen').evaluate()
    .setTitle('無線クエスト')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** スプレッドシートを開いたときに出る運営用メニュー。 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('無線クエスト')
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
  if (top.join('') === '' || top[0] !== head[0]) {
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
      nodesPerArea: NODES_PER_AREA,
      mobs: MOBS,
      questions: questions,
      me: handle ? loadOperator_(book, handle) : null,
      roster: readRoster_(book, conf),
      rules: {
        expPerLevel: toRange_(conf['レベルアップ経験値'], 1, 9999, 50),
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
    if (!AREAS.some(function (a) { return a.key === area; })) {
      return { ok: false, reason: 'ジャンルを選んでください' };
    }

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
  if (!AREAS.some(function (a) { return a.key === genre; })) return null;

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
    defeats: 0, hits: 0, tries: 0, posts: 0
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
      posts: Number(rows[i][8]) || 0
    };
  }
  return blank;
}

function writeOperator_(book, name, payload) {
  const sheet = book.getSheetByName(SHEET.op) || prepareSheet_(book, SHEET.op, HEAD.op);
  const current = loadOperator_(book, name);

  // 経験値や進行は減らさない。通信の行き違いで巻き戻るのを防ぐ
  const merged = {
    handle: name,
    level: Math.max(current.level, Math.max(1, Number(payload.level) || 1)),
    exp: Math.max(current.exp, Math.max(0, Number(payload.exp) || 0)),
    progress: mergeProgress_(current.progress, payload.progress),
    skills: mergeList_(current.skills, payload.skills),
    defeats: Math.max(current.defeats, Number(payload.defeats) || 0),
    hits: Math.max(current.hits, Number(payload.hits) || 0),
    tries: Math.max(current.tries, Number(payload.tries) || 0),
    posts: current.posts
  };

  saveOperatorRow_(sheet, merged);
  return merged;
}

function addPostReward_(book, name, reward, conf) {
  const sheet = book.getSheetByName(SHEET.op) || prepareSheet_(book, SHEET.op, HEAD.op);
  const me = loadOperator_(book, name);
  me.exp += reward;
  me.posts += 1;
  me.level = levelFor_(me.exp, toRange_(conf['レベルアップ経験値'], 1, 9999, 50));
  saveOperatorRow_(sheet, me);
  return me;
}

function saveOperatorRow_(sheet, me) {
  const line = [
    me.handle, me.level, me.exp, formatProgress_(me.progress), me.skills.join(','),
    me.defeats, me.hits, me.tries, me.posts, new Date()
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
    return {
      handle: String(row[0]).trim(),
      level: Math.max(1, Number(row[1]) || 1),
      exp: Math.max(0, Number(row[2]) || 0),
      skills: skills,
      posts: Number(row[8]) || 0,
      seenAt: row[9] instanceof Date ? row[9].getTime() : 0
    };
  });

  list.sort(function (x, y) { return y.exp - x.exp || y.level - x.level; });
  return list.slice(0, toRange_(conf['ランキング表示人数'], 1, 200, 20));
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
    if (key && isFinite(step)) out[key] = Math.max(0, Math.min(NODES_PER_AREA, Math.floor(step)));
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

/**
 * setupSheets() が最初の1回だけ入れる問題。
 * 運営はこの行を書き換えても、シートに直接足していってもよい。
 */
function seedQuestions_() {
  const rows = [
    ['W101', '電波の基礎', 1, '電波が真空中を伝わる速さは、およそどれくらいですか。',
     '秒速約3万km', '秒速約30万km', '秒速約300万km', '秒速約3000万km', 'B',
     '光と同じ速さで、毎秒およそ30万km（3×10^8 m/s）です。'],
    ['W102', '電波の基礎', 1, '1MHz は何Hzですか。',
     '1000Hz', '10万Hz', '100万Hz', '1億Hz', 'C',
     'M（メガ）は100万倍を表します。'],
    ['W103', '電波の基礎', 1, '周波数が高くなると、電波の波長はどうなりますか。',
     '長くなる', '短くなる', '変わらない', '周波数とは関係がない', 'B',
     '波長×周波数＝電波の速さ。速さは一定なので、周波数が上がれば波長は短くなります。'],
    ['W201', '電波の基礎', 2, '145MHz の電波の波長は、およそ何メートルですか。',
     '約0.7m', '約2m', '約6m', '約20m', 'B',
     '波長[m] ≒ 300 ÷ 周波数[MHz]。300÷145 ≒ 2.07m なので「2mバンド」と呼ばれます。'],
    ['W202', '電波の基礎', 2, '短波が遠くまで届くのは、おもに何で反射するからですか。',
     '対流圏', '電離層', 'オゾン層', '海面', 'B',
     '上空の電離層（おもにF層）と地表の間で反射をくり返して遠くまで届きます。'],
    ['W203', '電波の基礎', 2, '昼間にあらわれて中波を吸収してしまう電離層はどれですか。',
     'D層', 'E層', 'F1層', 'F2層', 'A',
     'D層は昼間だけ現れます。夜に消えるので、遠くのAM放送は夜のほうが聞こえやすくなります。'],
    ['W301', '電波の基礎', 3, '夏によく発生し、VHF帯で異常伝搬を起こす電離層は何と呼ばれますか。',
     'スポラディックE層', 'オーロラ層', 'ダクト層', 'グレイライン層', 'A',
     'Eスポと略します。ふだんは届かない距離まで50MHz帯などが飛びます。'],
    ['W302', '電波の基礎', 3, '電力が2倍になったとき、デシベルではおよそ何dBの増加ですか。',
     '約2dB', '約3dB', '約6dB', '約10dB', 'B',
     '10×log10(2) ≒ 3.01dB。10倍なら10dBです。'],
    ['W303', '電波の基礎', 3, '電波の偏波面が大地に対して水平になっているものを何といいますか。',
     '垂直偏波', '水平偏波', '円偏波', '楕円偏波', 'B',
     '送信側と受信側で偏波が合っていないと、受信できる信号が大きく弱まります。'],

    ['O101', '運用と通信術', 1, '相手を特定せず、誰にでも呼びかけるときに使う略号はどれですか。',
     'CQ', 'SOS', 'QRT', 'DX', 'A',
     '「CQ CQ こちらは…」と送ります。一般呼出の合図です。'],
    ['O102', '運用と通信術', 1, '交信した証としてやりとりするカードを何といいますか。',
     'QSLカード', 'QRZカード', 'RSTカード', 'DXカード', 'A',
     'QSLは「受信証明」を表すQ符号です。'],
    ['O103', '運用と通信術', 1, 'フォネティックコードで「A」はどう読みますか。',
     'Alfa', 'America', 'Apple', 'Adam', 'A',
     'ITUで決められた読み方です。Bは Bravo、Cは Charlie と続きます。'],
    ['O201', '運用と通信術', 2, 'Q符号の「QTH」は何を表しますか。',
     'こちらの位置', '周波数を変える', '送信をやめる', '了解した', 'A',
     '「QTHは東京です」のように使います。'],
    ['O202', '運用と通信術', 2, 'Q符号の「QRM」は何を表しますか。',
     '他局からの混信', '自然雑音による混信', '送信出力を下げる', '交信を終える', 'A',
     '人が出した電波による混信がQRM、雷などの自然雑音による混信はQRNです。'],
    ['O203', '運用と通信術', 2, 'RSTレポートの「S」は何を表しますか。',
     '了解度', '信号強度', '音調', '安定度', 'B',
     'R＝了解度、S＝信号強度、T＝音調。電信のときにTまで送ります。'],
    ['O301', '運用と通信術', 3, 'Q符号の「QSY」は何を表しますか。',
     '周波数を変更する', '受信できない', 'もう一度送ってほしい', '交信を終える', 'A',
     '混信を避けて別の周波数に移るときに使います。'],
    ['O302', '運用と通信術', 3, '交信の最後に送る「73」はどんな意味ですか。',
     'さようなら・ごきげんよう', '了解しました', 'もう一度お願いします', '出力を下げます', 'A',
     '電信時代からの数字略号です。88は「愛と口づけ」を意味します。'],
    ['O303', '運用と通信術', 3, 'モールス符号の「・・・ ーーー ・・・」はどれを表しますか。',
     'SOS', 'CQ', 'OSO', 'EEE', 'A',
     '国際的な遭難信号です。文字の区切りをつけず、ひと続きに送ります。'],

    ['A101', 'アンテナ', 1, '無線機とアンテナをつなぐ、外側が網状の導体になっているケーブルを何といいますか。',
     '同軸ケーブル', 'LANケーブル', '光ファイバ', '平行フィーダ', 'A',
     '網状の外部導体が内側の芯線を包み、電波の漏れと飛び込みを抑えます。'],
    ['A102', 'アンテナ', 1, 'アマチュア無線で最もよく使う同軸ケーブルの特性インピーダンスはどれですか。',
     '25Ω', '50Ω', '75Ω', '300Ω', 'B',
     '無線機側は50Ωが標準です。75Ωはおもにテレビ受信に使われます。'],
    ['A103', 'アンテナ', 1, '電波を出さずに送信機を調整するため、アンテナの代わりにつなぐものは何ですか。',
     'ダミーロード', 'アッテネータ', 'バラン', 'スタブ', 'A',
     '擬似負荷ともいいます。電力を熱に変えるので、まわりに迷惑をかけません。'],
    ['A201', 'アンテナ', 2, '半波長ダイポールアンテナの全長は、波長のどれだけですか。',
     '4分の1', '2分の1', '4分の3', '1波長', 'B',
     '左右に4分の1波長ずつのエレメントを伸ばし、中央で給電します。'],
    ['A202', 'アンテナ', 2, '半波長ダイポールアンテナの給電点インピーダンスは、およそいくらですか。',
     '約12Ω', '約36Ω', '約73Ω', '約300Ω', 'C',
     '50Ωの同軸とわりあい近いので、そのままつないでも大きくは崩れません。'],
    ['A203', 'アンテナ', 2, 'SWR（定在波比）が1.0のとき、どんな状態ですか。',
     '整合が完全にとれている', 'まったく整合していない', 'アンテナが断線している', '送信出力がゼロ', 'A',
     '進行波だけで反射波がない状態です。数値が大きいほど反射が増えています。'],
    ['A301', 'アンテナ', 3, '八木・宇田アンテナで、放射器より長く、後方に置く素子を何といいますか。',
     '反射器', '導波器', '整合器', '給電器', 'A',
     '長い素子が電波を前へ押し返します。英語ではリフレクタです。'],
    ['A302', 'アンテナ', 3, '八木・宇田アンテナで、放射器より短く、前方に並べる素子を何といいますか。',
     '反射器', '導波器', '接地板', 'トラップ', 'B',
     '導波器（ディレクタ）を増やすほど、鋭い指向性と高い利得が得られます。'],
    ['A303', 'アンテナ', 3, '4分の1波長垂直接地アンテナで、地面の代わりに放射状に張る導体を何といいますか。',
     'ラジアル', 'スタブ', 'バラン', 'トラップ', 'A',
     'ラジアルが大地の役目をします。これがないと性能が大きく落ちます。'],

    ['E101', '無線工学', 1, '10Ωの抵抗に2Aの電流が流れているとき、両端の電圧は何Vですか。',
     '5V', '12V', '20V', '100V', 'C',
     'オームの法則 E＝I×R より 2×10＝20V です。'],
    ['E102', '無線工学', 1, '同じ電池を2個直列につなぐと、電圧はどうなりますか。',
     '半分になる', '変わらない', '2倍になる', '4倍になる', 'C',
     '直列では電圧が加算されます。並列なら電圧はそのままで長もちします。'],
    ['E103', '無線工学', 1, '電気抵抗の単位はどれですか。',
     'アンペア', 'ボルト', 'ワット', 'オーム', 'D',
     '記号はΩ。アンペアは電流、ボルトは電圧、ワットは電力の単位です。'],
    ['E201', '無線工学', 2, 'コイルのリアクタンスは、周波数が高くなるとどうなりますか。',
     '大きくなる', '小さくなる', '変わらない', 'ゼロになる', 'A',
     'XL＝2πfL。周波数に比例して増えるので、高い周波数ほど通しにくくなります。'],
    ['E202', '無線工学', 2, 'コンデンサのリアクタンスは、周波数が高くなるとどうなりますか。',
     '大きくなる', '小さくなる', '変わらない', '無限大になる', 'B',
     'Xc＝1/(2πfC)。周波数に反比例するので、高い周波数ほどよく通します。'],
    ['E203', '無線工学', 2, '直列共振回路が共振したとき、インピーダンスはどうなりますか。',
     '最大になる', '最小になる', '変化しない', '負になる', 'B',
     'コイルとコンデンサのリアクタンスが打ち消し合い、抵抗ぶんだけが残ります。'],
    ['E301', '無線工学', 3, 'LC回路の共振周波数を表す式はどれですか。',
     'f ＝ 2π√(LC)', 'f ＝ 1/(2π√(LC))', 'f ＝ LC/(2π)', 'f ＝ 1/(2πLC)', 'B',
     'LやCが大きいほど、共振周波数は低くなります。'],
    ['E302', '無線工学', 3, '送信機から出てしまう、必要な電波以外の不要な成分をまとめて何といいますか。',
     'スプリアス', 'フェージング', 'サイドトーン', 'デビエーション', 'A',
     '高調波や寄生発振などが含まれます。ほかの無線や放送への妨害の原因になります。'],
    ['E303', '無線工学', 3, '送信機の出力に含まれる高調波を減らすために入れるフィルタはどれですか。',
     '高域通過フィルタ', '低域通過フィルタ', '帯域阻止フィルタ', 'ノッチフィルタ', 'B',
     '目的の周波数より高い成分を通さない低域通過フィルタ（LPF）を入れます。'],

    ['L101', '電波法規', 1, '日本で無線局を開設するとき、必要になるのはどれですか。',
     '総務大臣の免許', '市町村長の許可', '警察署への届出', '特に手続きはいらない', 'A',
     '電波法にもとづき、無線局には免許が必要です。'],
    ['L102', '電波法規', 1, 'アマチュア局の免許の有効期間は何年ですか。',
     '1年', '3年', '5年', '10年', 'C',
     '期間が切れる前に再免許の申請をします。'],
    ['L103', '電波法規', 1, '電波法の目的は、電波の公平かつ能率的な利用を確保して何を増進することですか。',
     '産業の発展', '公共の福祉', '国際競争力', '技術の輸出', 'B',
     '電波法第1条に書かれている目的です。'],
    ['L201', '電波法規', 2, '第二級アマチュア無線技士が操作できる空中線電力の上限はどれですか。',
     '10W', '50W', '200W', '制限なし', 'C',
     '第一級は制限なし、第二級は200Wまでと決められています。'],
    ['L202', '電波法規', 2, '電波法でいう「電波」とは、どの範囲の電磁波ですか。',
     '300万メガヘルツ以下', '3000メガヘルツ以下', '30メガヘルツ以下', 'すべての電磁波', 'A',
     '300万MHz（＝3THz）以下の周波数の電磁波と定義されています。'],
    ['L203', '電波法規', 2, '交信のときに必ず送信しなければならないものはどれですか。',
     '本名', '住所', '呼出符号（コールサイン）', '無線機の型番', 'C',
     '自局が誰であるかを示すため、呼出符号の送出が求められます。'],
    ['L301', '電波法規', 3, '非常通信が認められるのは、おもにどんなときですか。',
     '有線通信が使えず、人命救助などのために急を要するとき', '交信相手が見つからないとき',
     '出力を上げて遠くと交信したいとき', 'コンテストで順位を上げたいとき', 'A',
     '地震や台風などで、ふだんの通信手段が使えない場合が想定されています。'],
    ['L302', '電波法規', 3, 'アマチュア業務にあてはまらないのはどれですか。',
     '個人的な興味による自己訓練', '技術的研究のための通信', '金銭上の利益のための通信',
     '無線技術に興味を持つ人どうしの交信', 'C',
     'アマチュア業務は、金銭上の利益のためのものでないことが前提です。'],
    ['L303', '電波法規', 3, '免許状に記載された事項を変更したいとき、正しいのはどれですか。',
     '変更してから事後に報告すればよい', 'あらかじめ手続きが必要',
     '変更しても手続きはいらない', '無線機を買い替えたときだけ必要', 'B',
     '設置場所や無線設備を変えるときは、あらかじめ許可や届出が必要です。']
  ];

  return rows.map(function (row) { return row.concat(['初期問題', true]); });
}
