/**
 * Google Drive上の算数PDFをスキャンして、スプレッドシート「QSリスト」タブへ同期する。
 *
 * 仕様:
 * - A列: PDF相対パス（キー、例: "算数/6年/xxx.pdf"）
 * - B列: タイトル（ファイル名から拡張子除去）
 * - C列: 修正内容（fix系タグの文言を "/" 区切りで連結）
 * - D列: その他タグ（通常タグを "/" 区切りで連結）
 * - 既存行のE列以降は一切触らない
 * - 新規PDFは行追加、消えたPDFは行削除
 *
 * 必要設定:
 * - DRIVE_ROOT_FOLDER_ID: 「算数」フォルダ（またはその親）を指すフォルダID
 * - DRIVE_SCAN_SUBPATH: ルート配下のどこを同期対象にするか（例: "算数"）
 * - TEXT_MAPPING_FILE_ID: （任意）text_mapping.csv をDriveに置いた場合のファイルID
 */

const CONFIG = {
  SHEET_NAME: 'QSリスト',
  START_ROW: 2,
  // A,B,C,D を自動管理
  COL_KEY: 1,
  COL_TITLE: 2,
  COL_FIX: 3,
  COL_TAGS: 4,

  // Drive scan
  DRIVE_ROOT_FOLDER_ID: 'PUT_FOLDER_ID_HERE',
  DRIVE_SCAN_SUBPATH: '算数', // ルート直下が算数ならそのまま

  // Optional: text_mapping.csv（printviewerの形式）をDriveにアップロードしてIDを入れる
  TEXT_MAPPING_FILE_ID: '',
};

/**
 * メイン（時間主導トリガーで定期実行推奨）
 */
function syncQsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet not found: ${CONFIG.SHEET_NAME}`);

  const mappings = CONFIG.TEXT_MAPPING_FILE_ID ? loadTextMappingsFromDrive_(CONFIG.TEXT_MAPPING_FILE_ID) : {};

  // DriveからPDF一覧（キー: 相対パス）を取得
  const pdfs = listPdfsUnderSubpath_(CONFIG.DRIVE_ROOT_FOLDER_ID, CONFIG.DRIVE_SCAN_SUBPATH);

  // シート現状のキー→行番号
  const lastRow = Math.max(sheet.getLastRow(), CONFIG.START_ROW);
  const numRows = Math.max(0, lastRow - CONFIG.START_ROW + 1);
  const existing = numRows > 0
    ? sheet.getRange(CONFIG.START_ROW, CONFIG.COL_KEY, numRows, 1).getValues().map(r => (r[0] || '').toString().trim())
    : [];

  const keyToRow = new Map();
  existing.forEach((k, i) => {
    if (k) keyToRow.set(k, CONFIG.START_ROW + i);
  });

  const currentKeys = new Set(pdfs.map(p => p.key));

  // 削除（下から）
  const rowsToDelete = [];
  for (const [k, row] of keyToRow.entries()) {
    if (!currentKeys.has(k)) rowsToDelete.push(row);
  }
  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

  // 追加 or 更新（A-Dのみ）
  // 削除後に行番号が変わるので再取得
  const lastRow2 = Math.max(sheet.getLastRow(), CONFIG.START_ROW);
  const numRows2 = Math.max(0, lastRow2 - CONFIG.START_ROW + 1);
  const existing2 = numRows2 > 0
    ? sheet.getRange(CONFIG.START_ROW, CONFIG.COL_KEY, numRows2, 1).getValues().map(r => (r[0] || '').toString().trim())
    : [];
  const keyToRow2 = new Map();
  existing2.forEach((k, i) => {
    if (k) keyToRow2.set(k, CONFIG.START_ROW + i);
  });

  // 追加は末尾
  const updates = [];
  const appends = [];
  pdfs.forEach(p => {
    const tags = mappings[p.key] || [];
    const fixTexts = tags.filter(t => t.is_fix_tag).map(t => t.juku_name).filter(Boolean);
    const otherTexts = tags.filter(t => !t.is_fix_tag).map(t => t.juku_name).filter(Boolean);
    const fixJoined = uniq_(fixTexts).join('/');
    const otherJoined = uniq_(otherTexts).join('/');
    const title = stripPdfExt_(p.name);

    const rowData = [p.key, title, fixJoined, otherJoined];
    const row = keyToRow2.get(p.key);
    if (row) {
      updates.push({ row, values: rowData });
    } else {
      appends.push(rowData);
    }
  });

  // 更新（A-Dだけ書く）
  updates.forEach(u => {
    sheet.getRange(u.row, CONFIG.COL_KEY, 1, 4).setValues([u.values]);
  });
  if (appends.length) {
    sheet.getRange(sheet.getLastRow() + 1, CONFIG.COL_KEY, appends.length, 4).setValues(appends);
  }
}

/**
 * Drive: root配下の subpath（例: "算数"）を辿り、そこ以下のPDFを列挙して相対パスキーを作る。
 */
function listPdfsUnderSubpath_(rootFolderId, subpath) {
  const root = DriveApp.getFolderById(rootFolderId);
  const parts = (subpath || '').split('/').filter(Boolean);
  let base = root;
  parts.forEach(p => {
    const it = base.getFoldersByName(p);
    if (!it.hasNext()) throw new Error(`Folder not found under root: ${p}`);
    base = it.next();
  });

  const out = [];
  walk_(base, '', out);
  // key は "算数/..." 形式に揃える
  out.forEach(o => { o.key = (subpath ? (subpath.replace(/\\/g, '/').replace(/^\/|\/$/g, '') + '/') : '') + o.key; });
  return out;

  function walk_(folder, rel, acc) {
    const files = folder.getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      const f = files.next();
      acc.push({ key: (rel ? rel + '/' : '') + f.getName(), name: f.getName(), id: f.getId() });
    }
    const folders = folder.getFolders();
    while (folders.hasNext()) {
      const sub = folders.next();
      const name = sub.getName();
      walk_(sub, rel ? rel + '/' + name : name, acc);
    }
  }
}

/**
 * text_mapping.csv (printviewer形式) をDriveから読み、file_pathキーでタグ配列を返す。
 * 期待列: file_path,juku_name,text_name,commenter,needs_fix,is_fix_tag
 */
function loadTextMappingsFromDrive_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const csvText = file.getBlob().getDataAsString('UTF-8');
  const rows = Utilities.parseCsv(csvText);
  if (!rows || rows.length < 2) return {};
  const header = rows[0];
  const idx = (name) => header.indexOf(name);
  const iPath = idx('file_path');
  const iJuku = idx('juku_name');
  const iNeeds = idx('needs_fix');
  const iIsFix = idx('is_fix_tag');
  if (iPath < 0 || iJuku < 0) return {};

  const map = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const path = (row[iPath] || '').toString().trim();
    const juku = (row[iJuku] || '').toString().trim();
    if (!path || !juku) continue;
    const needsFix = parseBool_(row[iNeeds]);
    const isFix = iIsFix >= 0 ? parseBool_(row[iIsFix]) : needsFix;
    if (!map[path]) map[path] = [];
    map[path].push({ juku_name: juku, needs_fix: needsFix, is_fix_tag: isFix });
  }
  return map;
}

function parseBool_(v) {
  const s = (v == null ? '' : String(v)).trim().toLowerCase();
  return ['1', 'true', 'yes', 'y', 'on'].includes(s);
}

function stripPdfExt_(name) {
  return name && name.toLowerCase().endsWith('.pdf') ? name.slice(0, -4) : name;
}

function uniq_(arr) {
  const out = [];
  const set = new Set();
  arr.forEach(x => {
    const s = String(x);
    if (!set.has(s)) { set.add(s); out.push(s); }
  });
  return out;
}

