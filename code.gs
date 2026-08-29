/*
  地理クエスト
  Copyright (c) 2026 kimura yoshiki
  Note: https://note.com/cozy_auklet6005
  X: https://x.com/kimura_0314
  Version 1.8

  [利用規約・ライセンス]
  1. 利用許諾区分
      【無料版】
      - 未改変の状態(オリジナル)であれば、自由に再配布・共有が可能です。
      - 個人的な利用の範囲内であれば、自由に改変して使用できます。
      - ただし、改変したものを第三者へ再配布することは固く禁止します。

      【個人版(有料購入者)】
      - 購入者本人のみに利用権が付与されます。
      - 購入者本人が使用する限り、個人のPCおよび会社のPCで使用可能です。

      【法人版】
      - 購入した組織(法人)内部において、人数制限なく自由に複製・利用・改変が可能です。
      - 組織内部での配布は許可されます。

  2. 禁止事項
      - 改変されたソフトウェアの再配布(無料版・有料版問わず)。
      - 有料版を権利のない第三者へ配布、公開、販売する行為。

  3. 免責事項
      - 本ソフトウェアは「現状のまま」提供されます。本ソフトウェアの使用によって生じた
        いかなる損害(データ消失、業務の中断、予期せぬ動作等)についても、
        開発者は一切の責任を負いません。利用者は自己の責任において本ソフトウェアを使用するものとします。
*/

// ===== 定数 =====

const SHEET_QUESTION = '問題';
const SHEET_LOG = 'ログ';
const SHEET_PROGRESS = '進捗';
const SHEET_CONFIG = '設定';

const HEADER_QUESTION = ['問題ID', '分野', '難易度', '問題文', '選択肢1', '選択肢2', '選択肢3', '選択肢4', '正解番号', '解説'];
const HEADER_LOG = ['日時', 'クラス', '名前', '問題ID', '選んだ番号', '正誤', '解答秒数', 'レベル'];
const HEADER_PROGRESS = ['クラス', '名前', 'レベル', '経験値', '最終プレイ日時'];
const HEADER_CONFIG = ['項目', '値', '説明'];

// 設定シートの既定値。先生がシート上で書き換えられる
const DEFAULT_CONFIG = [
  ['クラス名', '1年A組,1年B組,1年C組,1年D組', 'カンマ区切り。タイトル画面の選択肢になります'],
  ['エンカウント率', '8', '1歩あたり敵に出会う確率(%)。数字を小さくすると出会いにくくなります'],
  ['レベルアップ必要経験値', '20', 'レベルが1上がるのに必要な経験値。小さいほど早く成長します'],
  ['出題分野', '', '空欄なら全分野。特定分野だけ出したい時は分野名をカンマ区切りで指定します'],
  ['効果音', 'ON', 'ON か OFF。教室で音を出したくない場合は OFF にします（生徒が自分で切り替えることもできます）']
];

// ===== Webアプリ入口 =====

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('無題');
  return template.evaluate()
    .setTitle('地理クエスト')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== 初回セットアップ =====

/**
 * 先生が最初に1回だけ実行する関数。
 * 4枚のシートを作り、サンプル問題を入れる。
 * 既にシートがある場合は中身を消さず、足りないものだけ足す。
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const qSheet = ensureSheet_(ss, SHEET_QUESTION, HEADER_QUESTION);
  if (qSheet.getLastRow() <= 1) {
    const rows = buildSampleQuestions_();
    qSheet.getRange(2, 1, rows.length, HEADER_QUESTION.length).setValues(rows);
    qSheet.setColumnWidth(4, 320);
    qSheet.setColumnWidth(10, 320);
  }

  ensureSheet_(ss, SHEET_LOG, HEADER_LOG);
  ensureSheet_(ss, SHEET_PROGRESS, HEADER_PROGRESS);

  const cSheet = ensureSheet_(ss, SHEET_CONFIG, HEADER_CONFIG);
  if (cSheet.getLastRow() <= 1) {
    cSheet.getRange(2, 1, DEFAULT_CONFIG.length, HEADER_CONFIG.length).setValues(DEFAULT_CONFIG);
    cSheet.setColumnWidth(2, 260);
    cSheet.setColumnWidth(3, 400);
  }

  // 既定のシート「シート1」が空のまま残っていたら消す
  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) {
    ss.deleteSheet(first);
  }

  Logger.log('セットアップ完了。問題シートに ' + Math.max(0, qSheet.getLastRow() - 1) + ' 問入っています。');
  return 'セットアップ完了';
}

/**
 * 初回アクセス時にシートが無ければ自動で作る。
 * 先生がGASエディタで setup を実行しなくても、URLを開いた時点で使い始められる。
 * 授業の開始直後に全員が同時にアクセスしても二重にサンプル問題が入らないよう、
 * 準備が要るときだけロックを取る。
 */
function ensureInitialized_(ss) {
  const qSheet = ss.getSheetByName(SHEET_QUESTION);
  if (qSheet && qSheet.getLastRow() > 1) return;   // 準備済み。ここが大半のケース

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;                // 他の人が初期化中なら任せる
  try {
    const again = ss.getSheetByName(SHEET_QUESTION);
    if (again && again.getLastRow() > 1) return;   // ロック待ちの間に誰かが終わらせた
    setup();
  } catch (err) {
    Logger.log('初期化に失敗: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function ensureSheet_(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const current = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const needHeader = current.join('') === '' || current[0] !== header[0];
  if (needHeader) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#e8eaf6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ===== クライアントから呼ばれる関数 =====

/**
 * ゲーム起動時に1回だけ呼ばれる。
 * 問題を全件と、その生徒の進捗をまとめて返す。
 * 以後は通信せずブラウザ側だけでゲームが進む(同時実行数を節約するための設計)。
 */
function getGameData(playerName, className) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureInitialized_(ss);
    const config = readConfig_(ss);

    const qSheet = ss.getSheetByName(SHEET_QUESTION);
    if (!qSheet || qSheet.getLastRow() <= 1) {
      return { ok: false, message: '問題シートが空です。GASエディタで setup を実行してください。' };
    }

    const values = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, HEADER_QUESTION.length).getValues();
    const fieldFilter = String(config['出題分野'] || '').split(',').map(function (s) { return s.trim(); }).filter(String);

    const questions = [];
    values.forEach(function (row, i) {
      const text = String(row[3] || '').trim();
      const answer = parseInt(row[8], 10);
      if (!text || !answer) return;
      const field = String(row[1] || 'その他').trim();
      if (fieldFilter.length && fieldFilter.indexOf(field) === -1) return;

      const choices = [row[4], row[5], row[6], row[7]].map(function (c) { return String(c == null ? '' : c).trim(); });
      if (choices.filter(String).length < 2) return;
      if (answer < 1 || answer > choices.length) return;
      // 正解に指定された選択肢が空欄だと、その問題は永久に正解できなくなる。出題しない
      if (!choices[answer - 1]) return;

      questions.push({
        id: String(row[0] || ('Q' + (i + 1))),
        field: field,
        level: Number(row[2]) || 1,
        text: text,
        choices: choices,
        answer: answer,
        note: String(row[9] || '').trim()
      });
    });

    if (!questions.length) {
      return { ok: false, message: '出題できる問題がありません。問題シートを確認してください。' };
    }

    const progress = (playerName && className) ? readProgress_(ss, className, playerName) : { level: 1, exp: 0 };

    return {
      ok: true,
      questions: questions,
      progress: progress,
      config: {
        classes: String(config['クラス名'] || '').split(',').map(function (s) { return s.trim(); }).filter(String),
        encounterRate: clampNumber_(config['エンカウント率'], 1, 60, 8),
        expPerLevel: clampNumber_(config['レベルアップ必要経験値'], 1, 9999, 20),
        soundEnabled: String(config['効果音'] == null ? 'ON' : config['効果音']).trim().toUpperCase() !== 'OFF'
      }
    };
  } catch (err) {
    return { ok: false, message: '読み込みに失敗しました: ' + err.message };
  }
}

/**
 * 名前が決まったあとに、その生徒の続きだけを取りに行く軽い関数。
 * 問題は起動時に取得済みなので、ここでは進捗しか読まない。
 */
function getProgress(playerName, className) {
  try {
    if (!playerName || !className) return { ok: true, progress: { level: 1, exp: 0 } };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return { ok: true, progress: readProgress_(ss, String(className).trim(), String(playerName).trim()) };
  } catch (err) {
    return { ok: false, progress: { level: 1, exp: 0 } };
  }
}

/**
 * 解答ログと進捗をまとめて保存する。
 * 1問ごとではなく、10問たまった時とゲーム終了時にだけ呼ばれる。
 */
function saveSession(payload) {
  const lock = LockService.getScriptLock();
  // 待ちすぎない。取れなければ失敗を返し、クライアント側が次回にまとめて送り直す
  if (!lock.tryLock(8000)) {
    return { ok: false, retry: true, message: '混み合っています' };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const className = sanitizeCell_(payload && payload.className);
    const playerName = sanitizeCell_(payload && payload.playerName);
    if (!className || !playerName) {
      return { ok: false, message: '名前とクラスが未設定です' };
    }

    const logs = (payload && payload.logs) || [];
    // 1回で受け取る上限。超えた分は保存されないので、クライアント側で分けて送る
    if (logs.length > 200) {
      return { ok: false, retry: true, message: '一度に送れるのは200件までです' };
    }
    if (logs.length) {
      const logSheet = ss.getSheetByName(SHEET_LOG) || ensureSheet_(ss, SHEET_LOG, HEADER_LOG);
      const rows = logs.map(function (l) {
        return [
          new Date(),
          className,
          playerName,
          sanitizeCell_(l.questionId),
          Number(l.selected) || 0,
          l.correct ? '○' : '×',
          Math.round((Number(l.seconds) || 0) * 10) / 10,
          Number(l.level) || 1
        ];
      });
      logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, HEADER_LOG.length).setValues(rows);
    }

    writeProgress_(ss, className, playerName, Number(payload.level) || 1, Number(payload.exp) || 0);
    return { ok: true };
  } catch (err) {
    return { ok: false, retry: true, message: '保存に失敗しました: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// ===== 進捗の読み書き =====

function readProgress_(ss, className, playerName) {
  const sheet = ss.getSheetByName(SHEET_PROGRESS);
  if (!sheet || sheet.getLastRow() <= 1) return { level: 1, exp: 0 };

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_PROGRESS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === className && String(values[i][1]).trim() === playerName) {
      return { level: Number(values[i][2]) || 1, exp: Number(values[i][3]) || 0 };
    }
  }
  return { level: 1, exp: 0 };
}

function writeProgress_(ss, className, playerName, level, exp) {
  const sheet = ss.getSheetByName(SHEET_PROGRESS) || ensureSheet_(ss, SHEET_PROGRESS, HEADER_PROGRESS);
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === className && String(values[i][1]).trim() === playerName) {
        sheet.getRange(i + 2, 3, 1, 3).setValues([[level, exp, new Date()]]);
        return;
      }
    }
  }
  sheet.appendRow([className, playerName, level, exp, new Date()]);
}

// ===== 設定 =====

function readConfig_(ss) {
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const config = {};
  DEFAULT_CONFIG.forEach(function (row) { config[row[0]] = row[1]; });
  if (!sheet || sheet.getLastRow() <= 1) return config;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  values.forEach(function (row) {
    const key = String(row[0] || '').trim();
    if (key) config[key] = row[1];
  });
  return config;
}

function clampNumber_(raw, min, max, fallback) {
  const n = Number(raw);
  if (!isFinite(n) || n < min || n > max) return fallback;
  return n;
}

/**
 * スプレッドシートに書き込む前に、先頭の = + - @ を無害化する。
 * 生徒が名前欄に数式を入れて他人のシートを壊すのを防ぐ(数式インジェクション対策)。
 */
function sanitizeCell_(value) {
  let s = String(value == null ? '' : value).trim().slice(0, 60);
  s = s.replace(/[\x00-\x1f\x7f]/g, '');
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

// ===== サンプル問題 =====

/**
 * setup() が最初の1回だけ入れるサンプル問題。
 * 先生はこの行を書き換えたり、下に追記したりして自由に差し替えられる。
 */
function buildSampleQuestions_() {
  return [
    ['Q01', '世界の地形', 1, '世界でいちばん面積が大きい国はどこですか。', 'ロシア', 'カナダ', '中国', 'アメリカ合衆国', 1, '日本の約45倍の面積があります。'],
    ['Q02', '世界の地形', 1, '六大陸のうち、いちばん面積が大きいのはどれですか。', 'アフリカ大陸', 'ユーラシア大陸', '北アメリカ大陸', '南アメリカ大陸', 2, 'ヨーロッパとアジアを合わせた大陸です。'],
    ['Q03', '世界の地形', 2, '三大洋のうち、いちばん面積が大きいのはどれですか。', '大西洋', 'インド洋', '太平洋', '北極海', 3, '日本の東側に広がる海です。'],
    ['Q04', '世界の地形', 2, '世界でいちばん高い山はどれですか。', 'エベレスト', 'キリマンジャロ', 'モンブラン', 'アコンカグア', 1, 'ヒマラヤ山脈にあり、標高は8,000mを超えます。'],
    ['Q05', '世界の地形', 2, 'アフリカ大陸を北へ流れ、地中海にそそぐ川はどれですか。', 'アマゾン川', 'ナイル川', 'ミシシッピ川', 'ライン川', 2, '古代エジプト文明が栄えた川です。'],
    ['Q06', '世界の地形', 3, '南アメリカ大陸を東西に流れ、流域面積が世界最大の川はどれですか。', 'アマゾン川', 'ナイル川', 'ドナウ川', 'インダス川', 1, '流域面積とは、その川に雨水が集まってくる範囲のことです。'],
    ['Q07', '世界の気候', 1, '赤道付近に広がる、1年中暑くて雨が多い気候帯はどれですか。', '寒帯', '乾燥帯', '熱帯', '冷帯', 3, '熱帯雨林が広がっています。'],
    ['Q08', '世界の気候', 2, '砂漠が広がり、1年を通して雨がとても少ない気候帯はどれですか。', '乾燥帯', '温帯', '熱帯', '寒帯', 1, 'サハラ砂漠などが代表です。'],
    ['Q09', '世界の気候', 2, '日本の大部分がふくまれる気候帯はどれですか。', '熱帯', '乾燥帯', '温帯', '寒帯', 3, '四季の変化がはっきりしています。'],
    ['Q10', '世界の気候', 3, '夏に乾燥し、冬に雨が多い気候を何といいますか。', '西岸海洋性気候', '地中海性気候', '温暖湿潤気候', 'ステップ気候', 2, 'イタリアやスペインなど地中海のまわりに広がります。'],
    ['Q11', '世界の気候', 3, '一年中こおったままの土地が広がる、極地に近い気候帯はどれですか。', '寒帯', '冷帯', '乾燥帯', '温帯', 1, '北極や南極のまわりに広がります。'],
    ['Q12', '日本の地形', 1, '日本でいちばん高い山はどれですか。', '北岳', '富士山', '穂高岳', '御嶽山', 2, '標高3,776m、静岡県と山梨県にまたがります。'],
    ['Q13', '日本の地形', 1, '日本でいちばん長い川はどれですか。', '利根川', '石狩川', '信濃川', '北上川', 3, '長野県から新潟県へ流れます。'],
    ['Q14', '日本の地形', 1, '日本でいちばん面積が大きい湖はどれですか。', '霞ヶ浦', '猪苗代湖', 'サロマ湖', '琵琶湖', 4, '滋賀県にあります。'],
    ['Q15', '日本の地形', 2, '日本の都道府県はいくつありますか。', '43', '45', '47', '49', 3, '1都1道2府43県です。'],
    ['Q16', '日本の地形', 2, '本州の中央部を南北に走る、けわしい山々のことを何といいますか。', '日本アルプス', 'カルデラ', 'リアス海岸', '扇状地', 1, '飛騨・木曽・赤石の3つの山脈をまとめた呼び名です。'],
    ['Q17', '日本の地形', 3, '海岸線が複雑に入りくんだ、のこぎりの歯のような地形を何といいますか。', '三角州', 'リアス海岸', '砂丘', '台地', 2, '三陸海岸や志摩半島が代表です。'],
    ['Q18', '日本の地形', 3, '川が山地から平地に出るところにできる、扇形の地形を何といいますか。', '三角州', '扇状地', 'カルデラ', '干潟', 2, '水はけがよく、果樹園に使われることが多い地形です。'],
    ['Q19', '日本の地形', 3, '日本の標準時のもとになる経線が通る都市はどこですか。', '兵庫県明石市', '東京都新宿区', '京都府京都市', '福岡県福岡市', 1, '東経135度の経線が通っています。'],
    ['Q20', '日本の産業', 1, '日本で自動車の生産がとくにさかんな、愛知県を中心とした工業地帯はどれですか。', '京浜工業地帯', '中京工業地帯', '阪神工業地帯', '北九州工業地帯', 2, '豊田市などが中心です。'],
    ['Q21', '日本の産業', 2, 'りんごの生産量が日本でいちばん多い県はどこですか。', '長野県', '山形県', '青森県', '岩手県', 3, '涼しい気候がりんごづくりに向いています。'],
    ['Q22', '日本の産業', 2, '冬でも暖かい気候をいかして、野菜の出荷時期を早める栽培方法を何といいますか。', '促成栽培', '抑制栽培', '輪作', '二毛作', 1, '高知平野や宮崎平野でさかんです。'],
    ['Q23', '日本の産業', 3, '涼しい気候をいかして、夏でも高原で野菜を作る栽培方法を何といいますか。', '促成栽培', '抑制栽培', '稲作', '酪農', 2, '長野県の高原レタスなどが代表です。'],
    ['Q24', '日本の産業', 3, '北海道でさかんな、乳牛を飼って牛乳やバターをつくる農業を何といいますか。', '酪農', '畑作', '果樹栽培', '林業', 1, '根釧台地などでさかんです。'],
    ['Q25', '日本の産業', 3, '原料を輸入し、製品にして輸出する日本の貿易のしかたを何といいますか。', '中継貿易', '加工貿易', '自由貿易', '保護貿易', 2, '資源が少ない日本が発展させてきた形です。'],
    ['Q26', '地図の読み方', 1, '地図で、同じ高さの地点を結んだ線を何といいますか。', '等高線', '経線', '緯線', '等圧線', 1, '線の間隔がせまいほど、傾きが急です。'],
    ['Q27', '地図の読み方', 2, '2万5千分の1の地図で1cmは、実際には何mですか。', '25m', '250m', '2,500m', '25,000m', 2, '25,000cm = 250m になります。'],
    ['Q28', '地図の読み方', 2, '地図で方位が示されていないとき、上はふつうどの方位を表しますか。', '東', '西', '南', '北', 4, '方位記号がない地図では、上が北と決まっています。'],
    ['Q29', '地図の読み方', 3, '赤道を0度として、南北のいちを表す線を何といいますか。', '緯線', '経線', '等高線', '日付変更線', 1, '南北それぞれ90度まであります。'],
    ['Q30', '地図の読み方', 3, 'イギリスのロンドンを通る、経度0度の線を何といいますか。', '日付変更線', '本初子午線', '赤道', '回帰線', 2, 'ここを基準に東経・西経を測ります。']
  ];
}
