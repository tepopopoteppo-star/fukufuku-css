/**
 * ペット供養祭 申込フォーム → スライドショー パイプライン
 * =====================================================
 * フォーム送信 → onFormSubmit 発火 → 回答シート最新行を読む → 写真を取得
 *   → ペット1頭につきスライド1枚（写真・お名前・ふりがな・参加回ラベル・一言メッセージ）を生成。
 *
 * 設計方針（引き継ぎ資料セクション3・6に準拠）:
 *   - スプレッドシート参照は openById で統一（コンテナバインド/スタンドアロン両対応）。
 *   - 列参照はすべて「ヘッダー名（部分一致）」ベース。列の増減・並べ替えに強い。
 *   - 写真は 1送信=1枚（単一URL）前提。複数連結の正規化は無害なので残置。
 *   - 一言メッセージは本文枠として配置。空欄ならメッセージ枠を出さない。
 *   - 音楽(BGM)は SlidesApp/Slides API では挿入不可。README / セクション8の手動手順で対応。
 */

/* ============================================================
 * 設定（CONFIG）
 * ============================================================ */
var CONFIG = {
  SPREADSHEET_ID: '1EWZeJxLKgoXCZzpBsQ2ny4HxTEbm6P6ucm6awkzYt4A',
  PRESENTATION_ID: '1EVy_prP1oQ4E662ZqHN4uVOGHp4cNFnr5sxLaHIyTZ4',
  FORM_ID: '1aby0FmzfKXSIbPiaNtc-n-bY1ae3s6yZEHF5h0yc4ps',
  NOTIFY_EMAIL: 'tepopopoteppo@gmail.com'
};

/**
 * ヘッダー名の部分一致キーワード。
 * フォームの設問文言が多少変わっても拾えるよう、候補を複数持たせる。
 * いずれも「最初に一致した列」を採用する。
 */
var COLUMN_KEYS = {
  photo:    ['写真'],
  name:     ['お名前', 'ペットのお名前', 'ペット名', '名前'], // ふりがな/読み を含む列は除外（下の resolveColumns で対処）
  furigana: ['ふりがな', 'よみ', '読み', 'フリガナ'],
  session:  ['参加回', '回', 'セッション', '時間'],
  message:  ['一言メッセージ', 'メッセージ', '一言', 'コメント']
};

/**
 * 参加回ラベル。フォームの選択肢テキストに含まれる時刻でマッチさせ、表示ラベルへ変換。
 * 旧値（12:00 / 14:00 / 納骨式15:00）は意図的に含めない（残存検出のため）。
 */
var SESSION_RULES = [
  { match: '10:00', label: '供養祭　10:00の回' },
  { match: '12:30', label: '供養祭　12:30の回' },
  { match: '15:00', label: '供養祭　15:00の回' },
  { match: '16:00', label: '納骨式　16:00' }
];

/* ============================================================
 * トリガー本体
 * ============================================================ */

/**
 * フォーム送信トリガー。
 * 注意: ファイルアップロード列は e.namedValues に入らないため、
 *       写真URLを含む全列を「回答シートの最終行」から直接読む。
 */
function onFormSubmit(e) {
  try {
    var sheet = getResponseSheet();
    var lastRow = sheet.getLastRow();
    var headers = getHeaders(sheet);
    var values = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];

    var slide = generateSlideForRow(headers, values);
    Logger.log('✓ スライド生成完了（row ' + lastRow + '）slideId=' + (slide ? slide.getObjectId() : 'なし'));
  } catch (err) {
    Logger.log('✗ onFormSubmit エラー: ' + err.message + '\n' + (err.stack || ''));
    notifyError('onFormSubmit', err);
  }
}

/* ============================================================
 * 共通ヘルパー（シート / ヘッダー / 列解決）
 * ============================================================ */

/** 回答シート（getActiveSheet は使わない）。openById＝バインド/スタンドアロン両対応。 */
function getResponseSheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheets()[0];
}

/** 1行目のヘッダー配列。 */
function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * ヘッダー名（部分一致）から列インデックス（0始まり）を返す。見つからなければ -1。
 * @param {Array} headers ヘッダー配列
 * @param {Array<string>} keywords いずれかを含めば一致
 * @param {Array<string>=} excludes これらを含む列は除外
 */
function findColumnIndex(headers, keywords, excludes) {
  excludes = excludes || [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]);
    var hit = keywords.some(function (k) { return h.indexOf(k) !== -1; });
    var bad = excludes.some(function (k) { return h.indexOf(k) !== -1; });
    if (hit && !bad) return i;
  }
  return -1;
}

/**
 * 全フィールドの列インデックスをまとめて解決。
 * 「お名前」列が「ふりがな/読み」を含む列に誤マッチしないよう excludes で保護。
 */
function resolveColumns(headers) {
  return {
    photo:    findColumnIndex(headers, COLUMN_KEYS.photo),
    name:     findColumnIndex(headers, COLUMN_KEYS.name, COLUMN_KEYS.furigana),
    furigana: findColumnIndex(headers, COLUMN_KEYS.furigana),
    session:  findColumnIndex(headers, COLUMN_KEYS.session),
    message:  findColumnIndex(headers, COLUMN_KEYS.message)
  };
}

/** 参加回の生値 → 表示ラベル。未知の値は警告ログを出し生値をそのまま返す。 */
function toSessionLabel(rawSession) {
  var s = String(rawSession || '');
  for (var i = 0; i < SESSION_RULES.length; i++) {
    if (s.indexOf(SESSION_RULES[i].match) !== -1) return SESSION_RULES[i].label;
  }
  Logger.log('⚠ 既知の参加回（10:00/12:30/15:00/16:00）に一致しません: "' + s + '"');
  return s;
}

/**
 * 写真セル（単一URL想定）から Drive ファイルIDを抽出。
 * 複数連結（__ や ", "）が来ても先頭1件を採用するよう正規化してから抽出。
 */
function extractDriveFileId(cell) {
  var raw = String(cell || '').trim();
  if (!raw) return null;
  var normalized = raw.split('__').join(' ').split(',').join(' '); // 連結正規化（1枚運用では無害）
  var m = normalized.match(/[-\w]{25,}/); // Drive ファイルID（25文字以上の英数・ハイフン・アンダースコア）
  return m ? m[0] : null;
}

/* ============================================================
 * スライド生成
 * ============================================================ */

/**
 * 1行ぶんのデータから 1枚のスライドを生成して返す。
 * @param {Array} headers ヘッダー配列
 * @param {Array} values  対象行の値配列
 * @return {Slide}
 */
function generateSlideForRow(headers, values) {
  var cols = resolveColumns(headers);

  var petName = cols.name >= 0 ? String(values[cols.name] || '').trim() : '';
  var furigana = cols.furigana >= 0 ? String(values[cols.furigana] || '').trim() : '';
  var sessionLabel = cols.session >= 0 ? toSessionLabel(values[cols.session]) : '';
  var message = cols.message >= 0 ? String(values[cols.message] || '').trim() : '';
  var photoCell = cols.photo >= 0 ? values[cols.photo] : '';

  var presentation = SlidesApp.openById(CONFIG.PRESENTATION_ID);
  var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // スライド寸法（既定 10in x 5.625in = 720pt x 405pt 相当）。
  var pageW = presentation.getPageWidth();
  var pageH = presentation.getPageHeight();

  // --- レイアウト領域 ---
  var margin = 24;
  var photoAreaW = pageW * 0.52;          // 左：写真エリア
  var textAreaX = photoAreaW + margin;    // 右：テキストエリア開始X
  var textAreaW = pageW - textAreaX - margin;

  // --- 写真 ---
  var fileId = extractDriveFileId(photoCell);
  if (fileId) {
    try {
      var blob = DriveApp.getFileById(fileId).getBlob();
      var img = slide.insertImage(blob);
      fitImageInto(img, margin, margin, photoAreaW - margin, pageH - margin * 2);
    } catch (e) {
      Logger.log('⚠ 写真取得失敗（認可切れ or ID不正）: ' + e.message + ' / id=' + fileId);
      insertTextBox(slide, '（写真を読み込めませんでした）', margin, pageH / 2 - 20, photoAreaW - margin, 40, 14, false);
    }
  } else {
    Logger.log('⚠ 写真URLからファイルIDを抽出できません: ' + photoCell);
    insertTextBox(slide, '（写真なし）', margin, pageH / 2 - 20, photoAreaW - margin, 40, 14, false);
  }

  // --- テキスト（上から: ふりがな → お名前 → 参加回ラベル → 一言メッセージ） ---
  var y = margin + 20;

  if (furigana) {
    insertTextBox(slide, furigana, textAreaX, y, textAreaW, 28, 16, false);
    y += 30;
  }

  // お名前（主役。大きめ）
  insertTextBox(slide, petName || '（お名前未入力）', textAreaX, y, textAreaW, 60, 36, true);
  y += 70;

  if (sessionLabel) {
    insertTextBox(slide, sessionLabel, textAreaX, y, textAreaW, 30, 16, false);
    y += 40;
  }

  // 一言メッセージ（空欄なら枠を出さない）
  if (message) {
    var msgH = pageH - y - margin;
    if (msgH < 60) msgH = 60;
    var box = insertTextBox(slide, message, textAreaX, y, textAreaW, msgH, 18, false);
    // 長文対策: 枠内に収まるよう自動縮小。
    box.getText().getTextStyle(); // no-op（参照確保）
    try {
      box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    } catch (e) {}
    autoshrinkText(box, 18, msgH, textAreaW);
  }

  return slide;
}

/** 画像をアスペクト比維持のまま指定領域に収める（中央寄せ）。 */
function fitImageInto(img, areaX, areaY, areaW, areaH) {
  var w = img.getWidth();
  var h = img.getHeight();
  if (!w || !h) return;
  var scale = Math.min(areaW / w, areaH / h);
  var newW = w * scale;
  var newH = h * scale;
  img.setWidth(newW);
  img.setHeight(newH);
  img.setLeft(areaX + (areaW - newW) / 2);
  img.setTop(areaY + (areaH - newH) / 2);
}

/** テキストボックスを挿入して返す。 */
function insertTextBox(slide, text, x, y, w, h, fontSize, bold) {
  var box = slide.insertTextBox(text, x, y, w, h);
  var style = box.getText().getTextStyle();
  style.setFontSize(fontSize);
  style.setBold(!!bold);
  return box;
}

/**
 * 長文を枠に収めるための簡易自動縮小。
 * 文字数と行数の概算から、はみ出しそうならフォントを段階的に下げる。
 */
function autoshrinkText(box, baseSize, boxH, boxW) {
  var text = box.getText().asString();
  var len = text.replace(/\s/g, '').length;
  var size = baseSize;
  // ざっくり: 全角1文字 ≒ size pt 幅、1行に boxW/size 文字、収容行数 ≒ boxH/(size*1.4)。
  while (size > 10) {
    var charsPerLine = Math.max(1, Math.floor(boxW / size));
    var lines = Math.ceil(len / charsPerLine);
    var capacity = Math.floor(boxH / (size * 1.4));
    if (lines <= capacity) break;
    size -= 1;
  }
  box.getText().getTextStyle().setFontSize(size);
}

/* ============================================================
 * 診断・テスト用
 * ============================================================ */

/**
 * 診断: 最新行の列構成・写真URL・ファイルID抽出・DriveApp認可を一括確認。
 * 事前にフォームからテスト送信を1件しておき、本関数を実行 → 「実行ログ」を確認。
 */
function diagnoseLatestResponse() {
  var sheet = getResponseSheet();
  var lastRow = sheet.getLastRow();
  var headers = getHeaders(sheet);
  var values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  Logger.log('=== 最新行 row ' + lastRow + ' / 列数 ' + headers.length + ' ===');
  headers.forEach(function (h, i) { Logger.log('[' + i + '] ' + h + '  =>  ' + values[i]); });

  var cols = resolveColumns(headers);
  Logger.log('--- 列解決結果 ---');
  Logger.log('写真     idx=' + cols.photo);
  Logger.log('お名前   idx=' + cols.name);
  Logger.log('ふりがな idx=' + cols.furigana);
  Logger.log('参加回   idx=' + cols.session);
  Logger.log('メッセージidx=' + cols.message);

  if (cols.session >= 0) {
    Logger.log('参加回ラベル: ' + toSessionLabel(values[cols.session]));
  }

  if (cols.photo === -1) { Logger.log('⚠ 見出しに「写真」を含む列がありません'); return; }
  var raw = String(values[cols.photo]).trim();
  Logger.log('写真セル(生): ' + raw);
  var fileId = extractDriveFileId(raw);
  if (!fileId) { Logger.log('⚠ URLからファイルIDを抽出できません'); return; }
  Logger.log('抽出ID: ' + fileId);
  try {
    var f = DriveApp.getFileById(fileId);
    Logger.log('✓ DriveApp OK: ' + f.getName() + ' / ' + f.getMimeType());
  } catch (e) {
    Logger.log('⚠ DriveApp取得失敗（認可切れ or ID不正）: ' + e.message);
  }
}

/**
 * 単体テスト: フォーム送信せずに最新行からスライドを1枚生成する。
 * 反復検証用。生成後の slideId をログに出す。
 */
function testGenerateSlideFromLatestRow() {
  var sheet = getResponseSheet();
  var lastRow = sheet.getLastRow();
  var headers = getHeaders(sheet);
  var values = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
  var slide = generateSlideForRow(headers, values);
  Logger.log('✓ テストスライド生成: slideId=' + slide.getObjectId());
}

/**
 * 単体テスト（写真なし）: ダミーデータでレイアウトのみ検証。
 * Drive 認可やフォーム送信に依存せず、テキスト配置・空欄分岐を確認できる。
 */
function testGenerateSlideWithDummy() {
  var headers = ['タイムスタンプ', 'ペットのお名前', 'ふりがな（読み）', '参加回', '写真', '一言メッセージ'];
  var values = [
    new Date(),
    'ふく',
    'ふく',
    '供養祭 12:30',
    '', // 写真なし（→「（写真なし）」表示の分岐確認）
    'いつもそばにいてくれてありがとう。たくさんの思い出をありがとう。これからもずっと忘れないよ。'
  ];
  var slide = generateSlideForRow(headers, values);
  Logger.log('✓ ダミースライド生成: slideId=' + slide.getObjectId());
}

/**
 * 単体テスト（空メッセージ）: 一言メッセージ空欄時にメッセージ枠が出ないことを確認。
 */
function testGenerateSlideEmptyMessage() {
  var headers = ['タイムスタンプ', 'ペットのお名前', 'ふりがな（読み）', '参加回', '写真', '一言メッセージ'];
  var values = [new Date(), 'もも', 'もも', '供養祭 10:00', '', ''];
  var slide = generateSlideForRow(headers, values);
  Logger.log('✓ 空メッセージスライド生成: slideId=' + slide.getObjectId());
}

/* ============================================================
 * 一括再生成 / トリガー設定 / 通知
 * ============================================================ */

/**
 * 全回答行からスライドを一括再生成（追記）。
 * 既存スライドを消したい場合は事前に Slides UI で手動削除するか、別途クリア関数を用意のこと。
 */
function rebuildAllSlides() {
  var sheet = getResponseSheet();
  var headers = getHeaders(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('データ行がありません'); return; }
  var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var ok = 0;
  data.forEach(function (row, i) {
    try {
      generateSlideForRow(headers, row);
      ok++;
    } catch (e) {
      Logger.log('⚠ row ' + (i + 2) + ' 生成失敗: ' + e.message);
    }
  });
  Logger.log('✓ 一括再生成 ' + ok + '/' + data.length + ' 件');
}

/**
 * onFormSubmit トリガーを（重複なく）設置する。エディタで一度実行。
 * フォーム送信トリガーはスプレッドシートに対して設定する。
 */
function installFormSubmitTrigger() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function (t) {
    return t.getHandlerFunction() === 'onFormSubmit' &&
           t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT;
  });
  if (exists) { Logger.log('既に onFormSubmit トリガーが存在します'); return; }
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log('✓ onFormSubmit トリガーを設置しました');
}

/** エラー時のメール通知。 */
function notifyError(where, err) {
  try {
    MailApp.sendEmail(
      CONFIG.NOTIFY_EMAIL,
      '[供養祭スライド] エラー: ' + where,
      'エラー: ' + err.message + '\n\nスタック:\n' + (err.stack || '(なし)')
    );
  } catch (e) {
    Logger.log('通知メール送信失敗: ' + e.message);
  }
}
