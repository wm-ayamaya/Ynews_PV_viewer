/**
 * Yahoo!ニュース Insights 記事一覧TSVのアップロード用 Web アプリ
 * 共有ドライブ上のフォルダへ、元ファイルと整理済みTSVの両方を保存する。
 *
 * 整理内容: 列は「配信日時」「記事タイトル」「合計PV数」「URL」。配信日時の昇順。
 * 指定した開始日（年月日）以降の配信のみ残す（当日 0:00 を含む）。
 *
 * 事前設定（スクリプトプロパティ）:
 * - UPLOAD_FOLDER_ID … アップロード時の保存先フォルダの ID
 * - DATA_FOLDER_ID … 整理済みTSVの置き場（閲覧画面はここを参照）。未設定時は UPLOAD_FOLDER_ID と同じ扱い
 *
 * デプロイ: Web アプリとして公開。「次のユーザーとして実行」= 自分
 */
var PROP_UPLOAD_FOLDER_ID = 'UPLOAD_FOLDER_ID';
var PROP_DATA_FOLDER_ID = 'DATA_FOLDER_ID';

var HDR_TITLE = '記事タイトル';
var HDR_DATE = '配信日時';
var HDR_URL = 'URL';
/** 半角括弧／全角括弧の表記揺れに対応 */
var HDR_PV_CANDIDATES = ['合計(PV数)', '合計（PV数）'];
var OUT_HEADER = '配信日時\t記事タイトル\t合計PV数\tURL';

/** アップロード処理が付与する整理ファイル名の末尾パターン */
var PROCESSED_FILENAME_PATTERN = /_整理_\d{8}\.tsv$/i;

/** 閲覧画面で返す行の上限（google.script.run の転送量対策） */
var MAX_VIEW_ROWS = 5000;

function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? String(e.parameter.page) : '';
  var xfo = HtmlService.XFrameOptionsMode.ALLOWALL;
  if (page === 'view') {
    return HtmlService.createHtmlOutputFromFile('view')
      .setTitle('Yahoo!ニュース記事・PV数一覧')
      .setXFrameOptionsMode(xfo);
  }
  return HtmlService.createHtmlOutputFromFile('Upload')
    .setTitle('記事一覧TSVアップロード')
    .setXFrameOptionsMode(xfo);
}

/**
 * 同一デプロイの Web アプリ URL（クエリなし）。ナビ用。
 * @return {string}
 */
function getWebAppBaseUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (err) {
    return '';
  }
}

function getUploadStatus() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_UPLOAD_FOLDER_ID);
  return { ready: !!id };
}

/**
 * @param {string} base64Content
 * @param {string} fileName
 * @return {{
 *   ok: boolean,
 *   message?: string,
 *   originalFileUrl?: string,
 *   processedFileUrl?: string,
 *   rowCount?: number,
 *   filterFromDate?: string
 * }}
 */
function saveUploadedTsv(base64Content, fileName, fromDateIso) {
  if (!base64Content || !fileName) {
    return { ok: false, message: 'ファイルが空か、名前がありません。' };
  }

  var fromIso = String(fromDateIso || '').trim();
  if (!fromIso) {
    return { ok: false, message: '整理に含める配信日の開始日（年月日）を指定してください。' };
  }
  if (!parseIsoDateStart(fromIso)) {
    return { ok: false, message: '開始日は YYYY-MM-DD 形式で指定してください。' };
  }

  var normalizedName = String(fileName).trim();
  if (!/\.tsv$/i.test(normalizedName)) {
    return { ok: false, message: '拡張子は .tsv のファイルのみアップロードできます。' };
  }

  var folderId = PropertiesService.getScriptProperties().getProperty(PROP_UPLOAD_FOLDER_ID);
  if (!folderId) {
    return {
      ok: false,
      message:
        '管理者向け: スクリプトプロパティ「' +
        PROP_UPLOAD_FOLDER_ID +
        '」に保存先フォルダ ID を設定してください。',
    };
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    return {
      ok: false,
      message: '保存先フォルダにアクセスできません。フォルダ ID と共有ドライブの権限を確認してください。',
    };
  }

  var mime = 'text/tab-separated-values; charset=utf-8';
  var bytes;
  try {
    bytes = Utilities.base64Decode(base64Content);
  } catch (e) {
    return { ok: false, message: 'ファイルのデコードに失敗しました。' };
  }

  var text = decodeTsvTextForInsights(bytes);
  var built = buildSortedSummaryTsv(text, fromIso);
  if (built.error) {
    return { ok: false, message: built.error };
  }

  var originalBlob = Utilities.newBlob(bytes, mime, normalizedName);
  var dateTag = fromIso.replace(/-/g, '');
  var processedName = normalizedName.replace(/\.tsv$/i, '_整理_' + dateTag + '.tsv');
  var processedBlob = Utilities.newBlob(
    '\uFEFF' + built.tsvBody,
    mime,
    processedName
  );

  try {
    var originalFile = folder.createFile(originalBlob);
    var processedFile = folder.createFile(processedBlob);
    return {
      ok: true,
      originalFileUrl: originalFile.getUrl(),
      processedFileUrl: processedFile.getUrl(),
      rowCount: built.rowCount,
      filterFromDate: fromIso,
    };
  } catch (e) {
    return { ok: false, message: '保存に失敗しました: ' + e.message };
  }
}

/**
 * BOM と文字コード候補を試し、Insights ヘッダーが解読できる文字列を返す。
 * メモ帳の「Unicode」＝ UTF-16 LE（BOM FF FE）などに対応。
 * @param {byte[]} bytes
 * @return {string}
 */
function decodeTsvTextForInsights(bytes) {
  var len = bytes.length;
  var b0 = len > 0 ? bytes[0] & 0xff : 0;
  var b1 = len > 1 ? bytes[1] & 0xff : 0;
  var b2 = len > 2 ? bytes[2] & 0xff : 0;

  var order = [];
  if (len >= 2 && b0 === 0xff && b1 === 0xfe) {
    order.push('UTF-16LE');
  }
  if (len >= 2 && b0 === 0xfe && b1 === 0xff) {
    order.push('UTF-16BE');
  }
  if (len >= 3 && b0 === 0xef && b1 === 0xbb && b2 === 0xbf) {
    order.push('UTF-8');
  }

  var defaults = ['UTF-8', 'Shift_JIS', 'UTF-16LE', 'UTF-16BE'];
  for (var d = 0; d < defaults.length; d++) {
    if (order.indexOf(defaults[d]) < 0) {
      order.push(defaults[d]);
    }
  }

  for (var i = 0; i < order.length; i++) {
    try {
      var text = Utilities.newBlob(bytes).getDataAsString(order[i]);
      if (findInsightsHeaderIndices(text)) {
        return text;
      }
    } catch (e) {
      continue;
    }
  }
  try {
    return Utilities.newBlob(bytes).getDataAsString('UTF-8');
  } catch (e2) {
    try {
      return Utilities.newBlob(bytes).getDataAsString('UTF-16LE');
    } catch (e3) {
      return Utilities.newBlob(bytes).getDataAsString();
    }
  }
}

function trimHeaderCell(s) {
  return String(s || '')
    .replace(/^\uFEFF/, '')
    .replace(/\r$/, '')
    .trim();
}

/**
 * @param {string} text
 * @return {{ headerIdx: number, colTitle: number, colDate: number, colPv: number, colUrl: number } | null}
 */
function findInsightsHeaderIndices(text) {
  var stripped = text.replace(/^\uFEFF/, '');
  var lines = stripped.split(/\r?\n/);
  var maxScan = Math.min(lines.length, 40);
  for (var h = 0; h < maxScan; h++) {
    var rawCells = lines[h].split('\t');
    var cells = [];
    for (var j = 0; j < rawCells.length; j++) {
      cells.push(trimHeaderCell(rawCells[j]));
    }
    var ti = cells.indexOf(HDR_TITLE);
    var di = cells.indexOf(HDR_DATE);
    var ui = cells.indexOf(HDR_URL);
    var pi = -1;
    for (var k = 0; k < HDR_PV_CANDIDATES.length; k++) {
      pi = cells.indexOf(HDR_PV_CANDIDATES[k]);
      if (pi >= 0) break;
    }
    if (ti >= 0 && di >= 0 && pi >= 0 && ui >= 0) {
      return { headerIdx: h, colTitle: ti, colDate: di, colPv: pi, colUrl: ui };
    }
  }
  return null;
}

/**
 * @param {string} iso YYYY-MM-DD
 * @return {Date | null} その日 0:00（スクリプトタイムゾーン）
 */
function parseIsoDateStart(iso) {
  var m = String(iso)
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  var y = Number(m[1]);
  var mo = Number(m[2]) - 1;
  var d = Number(m[3]);
  if (mo < 0 || mo > 11 || d < 1 || d > 31) return null;
  var dt = new Date(y, mo, d, 0, 0, 0, 0);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) return null;
  return dt;
}

/**
 * Insights 記事一覧TSVから必要列のみ抽出し、配信日時昇順のTSV本文（BOMなし）を返す
 * @param {string} text
 * @param {string} fromDateIso この日 0:00 以降の配信のみ（YYYY-MM-DD）。必須。
 * @return {{ tsvBody: string, rowCount: number } | { error: string }}
 */
function buildSortedSummaryTsv(text, fromDateIso) {
  var meta = findInsightsHeaderIndices(text);
  if (!meta) {
    return {
      error:
        'Insights 記事一覧形式のヘッダー（' +
        HDR_TITLE +
        '・' +
        HDR_DATE +
        '・' +
        HDR_PV_CANDIDATES.join(' / ') +
        '・' +
        HDR_URL +
        '）が見つかりません。UTF-8 / Shift_JIS / UTF-16 LE の TSV をご利用ください。',
    };
  }

  var headerIdx = meta.headerIdx;
  var colTitle = meta.colTitle;
  var colDate = meta.colDate;
  var colPv = meta.colPv;
  var colUrl = meta.colUrl;

  var filterStart = parseIsoDateStart(fromDateIso);
  if (!filterStart) {
    return { error: '開始日の形式が正しくありません（YYYY-MM-DD）。' };
  }
  var filterMs = filterStart.getTime();

  var stripped = text.replace(/^\uFEFF/, '');
  var lines = stripped.split(/\r?\n/);

  /** @type {{ sortKey: number, dateRaw: string, title: string, pv: string, url: string }[]} */
  var rows = [];
  for (var r = headerIdx + 1; r < lines.length; r++) {
    var line = lines[r];
    if (!line) continue;
    var rowCells = line.split('\t');
    if (rowCells.length <= Math.max(colTitle, colDate, colPv, colUrl)) continue;

    var dateRaw = String(rowCells[colDate] || '').trim();
    var title = String(rowCells[colTitle] || '');
    var pvRaw = rowCells[colPv];
    var urlRaw = String(rowCells[colUrl] || '').trim();
    if (!dateRaw && !title.trim()) continue;

    var d = parsePublishDate(dateRaw);
    if (!d || d.getTime() < filterMs) {
      continue;
    }
    var sortKey = d.getTime();
    var pvDigits = normalizePvCell(pvRaw);

    rows.push({ sortKey: sortKey, dateRaw: dateRaw, title: title, pv: pvDigits, url: urlRaw });
  }

  rows.sort(function (a, b) {
    if (a.sortKey !== b.sortKey) return a.sortKey - b.sortKey;
    return a.title.localeCompare(b.title, 'ja');
  });

  var outLines = [OUT_HEADER];
  for (var i = 0; i < rows.length; i++) {
    outLines.push(
      escapeTsvField(rows[i].dateRaw) +
        '\t' +
        escapeTsvField(rows[i].title) +
        '\t' +
        escapeTsvField(rows[i].pv) +
        '\t' +
        escapeTsvField(rows[i].url)
    );
  }

  return { tsvBody: outLines.join('\n'), rowCount: rows.length };
}

/**
 * 例: 2026/04/05(日) 15:27
 * @param {string} str
 * @return {Date | null}
 */
function parsePublishDate(str) {
  var m = String(str)
    .trim()
    .match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\([^)]*\)\s+(\d{1,2}):(\d{1,2})$/);
  if (!m) return null;
  var y = Number(m[1]);
  var mo = Number(m[2]) - 1;
  var d = Number(m[3]);
  var hh = Number(m[4]);
  var mm = Number(m[5]);
  return new Date(y, mo, d, hh, mm, 0, 0);
}

function normalizePvCell(cell) {
  if (cell == null || cell === '') return '0';
  return String(cell).replace(/,/g, '').trim() || '0';
}

function escapeTsvField(s) {
  var str = String(s);
  if (/[\t\n\r"]/.test(str)) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * TSV 1行をタブ区切りで分割（ダブルクォート・エスケープ対応）
 * @param {string} line
 * @return {string[]}
 */
function parseTsvLine(line) {
  var result = [];
  var cur = '';
  var inQ = false;
  for (var i = 0; i < line.length; i++) {
    var c = line.charAt(i);
    if (inQ) {
      if (c === '"') {
        if (i + 1 < line.length && line.charAt(i + 1) === '"') {
          cur += '"';
          i++;
        } else {
          inQ = false;
        }
      } else {
        cur += c;
      }
    } else {
      if (c === '\t') {
        result.push(cur);
        cur = '';
      } else if (c === '"') {
        inQ = true;
      } else {
        cur += c;
      }
    }
  }
  result.push(cur);
  return result;
}

/**
 * 整理済みTSVバイト列を文字列化（UTF-8 / Shift_JIS / UTF-16 LE を簡易判定）
 * @param {byte[]} bytes
 * @return {string}
 */
function decodeTextForProcessedView(bytes) {
  var encs = ['UTF-8', 'Shift_JIS', 'UTF-16LE'];
  for (var e = 0; e < encs.length; e++) {
    try {
      var t = Utilities.newBlob(bytes).getDataAsString(encs[e]).replace(/^\uFEFF/, '');
      if (t.indexOf('配信日時') !== -1 && t.indexOf('記事タイトル') !== -1) {
        return t;
      }
    } catch (err) {
      continue;
    }
  }
  return Utilities.newBlob(bytes).getDataAsString('UTF-8').replace(/^\uFEFF/, '');
}

/**
 * データ置き場（DATA_FOLDER_ID、未設定時は UPLOAD_FOLDER_ID）内の
 * 「*_整理_YYYYMMDD.tsv」のうち最終更新が新しい1件を読み、表用データを返す。
 * @return {{
 *   ok: boolean,
 *   message?: string,
 *   fileName?: string,
 *   fileUrl?: string,
 *   lastUpdated?: string,
 *   headers?: string[],
 *   rows?: string[][],
 *   rowCount?: number,
 *   truncated?: boolean,
 *   folderNote?: string
 * }}
 */
function resolveProcessedDataFolder_() {
  var props = PropertiesService.getScriptProperties();
  var dataId = props.getProperty(PROP_DATA_FOLDER_ID);
  if (dataId && String(dataId).trim()) {
    return {
      folderId: String(dataId).trim(),
      folderNote: '参照フォルダ: データ置き場（DATA_FOLDER_ID）',
    };
  }
  var uploadId = props.getProperty(PROP_UPLOAD_FOLDER_ID);
  return {
    folderId: uploadId ? String(uploadId).trim() : '',
    folderNote: '参照フォルダ: アップロード先（DATA_FOLDER_ID 未設定のため UPLOAD_FOLDER_ID）',
  };
}

function loadLatestProcessedArticles() {
  var resolved = resolveProcessedDataFolder_();
  var folderId = resolved.folderId;
  if (!folderId) {
    return {
      ok: false,
      message:
        'データ置き場が未設定です。スクリプトプロパティ「' +
        PROP_DATA_FOLDER_ID +
        '」に整理TSV用フォルダ ID を設定するか、「' +
        PROP_UPLOAD_FOLDER_ID +
        '」を設定してください。',
    };
  }
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e1) {
    return { ok: false, message: 'データ置き場フォルダにアクセスできません。ID と共有ドライブの権限を確認してください。' };
  }

  var it = folder.getFiles();
  /** @type {{ f: Object, t: number }[]} */
  var found = [];
  while (it.hasNext()) {
    var f = it.next();
    var n = f.getName();
    if (PROCESSED_FILENAME_PATTERN.test(n)) {
      found.push({ f: f, t: f.getLastUpdated().getTime() });
    }
  }
  if (found.length === 0) {
    return {
      ok: false,
      message:
        'データ置き場に整理済みTSV（ファイル名が「*_整理_YYYYMMDD.tsv」）が見つかりません。' +
        (resolved.folderNote ? '（' + resolved.folderNote + '）' : ''),
    };
  }
  found.sort(function (a, b) {
    return b.t - a.t;
  });

  var file = found[0].f;
  var bytes = file.getBlob().getBytes();
  var text = decodeTextForProcessedView(bytes);
  var lines = text.split(/\r?\n/).filter(function (ln) {
    return ln.length > 0;
  });
  if (lines.length === 0) {
    return { ok: false, message: 'ファイルが空です。' };
  }

  var headers = parseTsvLine(lines[0]);
  var rows = [];
  var truncated = false;
  var maxData = MAX_VIEW_ROWS;
  for (var r = 1; r < lines.length; r++) {
    if (rows.length >= maxData) {
      truncated = true;
      break;
    }
    rows.push(parseTsvLine(lines[r]));
  }

  var tz = Session.getScriptTimeZone();
  var lastStr = Utilities.formatDate(file.getLastUpdated(), tz, 'yyyy-MM-dd HH:mm:ss');

  return {
    ok: true,
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    lastUpdated: lastStr,
    headers: headers,
    rows: rows,
    rowCount: rows.length,
    truncated: truncated,
    folderNote: resolved.folderNote,
  };
}
