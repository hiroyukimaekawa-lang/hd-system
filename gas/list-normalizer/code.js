/**
 * HD事業部 CSV統合・正規化・重複排除・チェーン店除外 GAS テンプレート
 * 
 * ※ 本スクリプトは基本骨格テンプレートです。
 * ※ APIキー等の機密情報は、スクリプトプロパティ「PropertiesService」で管理します。
 */

// =====================================================================
// 設定・プロパティ取得
// =====================================================================
function getApiKey() {
  // 安全設計: APIキーはコードに直書きせず、スクリプトプロパティから取得
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty('EXTERNAL_SEARCH_API_KEY');
  if (!apiKey) {
    console.warn('警告: EXTERNAL_SEARCH_API_KEY がスクリプトプロパティに設定されていません。');
  }
  return apiKey;
}

// =====================================================================
// CSV正規化・パース処理
// =====================================================================
/**
 * 住所から都道府県と市区町村をパースする
 * @param {string} address 
 * @return {object} {prefecture, city}
 */
function parseAddress(address) {
  if (!address) return { prefecture: '', city: '' };
  
  // 郵便番号や「日本、」のプレフィックスを除去
  let cleanAddress = address.replace(/(?:〒\d{3}-\d{4}\s*|日本、\s*)/g, '').trim();
  
  const regex = /^((?:北海道|東京都|大阪府|京都府|.{2,3}県))?((?:.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村]))?(.+)?$/;
  const match = cleanAddress.match(regex);
  
  if (!match) return { prefecture: '', city: '' };
  return {
    prefecture: match[1] || '',
    city: match[2] || ''
  };
}

/**
 * 電話番号を正規化する（全角の半角化、ハイフン有無の統一）
 * @param {string} phone 
 * @return {string}
 */
function normalizePhoneNumber(phone) {
  if (!phone) return '';
  // 全角数字・記号を半角に変換
  let clean = phone.replace(/[０-９ー－]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 65248);
  });
  // 数字以外の文字を除外して標準的なハイフン形式にする
  clean = clean.replace(/[^0-9]/g, '');
  
  if (clean.startsWith('0') && clean.length >= 9) {
    // 携帯電話または一般固定電話の簡易整形
    if (clean.startsWith('090') || clean.startsWith('080') || clean.startsWith('070')) {
      return clean.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
    } else if (clean.startsWith('03') || clean.startsWith('06')) {
      return clean.replace(/(\d{2})(\d{4})(\d{4})/, '$1-$2-$3');
    } else {
      return clean.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }
  }
  return phone;
}

// =====================================================================
// 重複排除ロジック
// =====================================================================
/**
 * リスト上の重複店舗をマージまたは排除する
 */
function deduplicateList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const header = data[0];
  const rows = data.slice(1);
  
  const phoneIdx = header.indexOf('電話番号');
  const nameIdx = header.indexOf('店名');
  const addrIdx = header.indexOf('住所');
  
  const seenPhones = new Map(); // phone -> rowIndex
  const seenNameAddrs = new Map(); // "name_address" -> rowIndex
  const rowsToKeep = [];
  const duplicateCandidates = [];

  rows.forEach((row, idx) => {
    const phone = normalizePhoneNumber(row[phoneIdx]);
    const name = String(row[nameIdx]).trim().replace(/\s+/g, '');
    const address = String(row[addrIdx]).trim().replace(/\s+/g, '');
    const nameAddrKey = `${name}_${address}`;
    
    let isDuplicate = false;
    
    // 1. 電話番号による完全一致
    if (phone && seenPhones.has(phone)) {
      isDuplicate = true;
      duplicateCandidates.push({ index: idx + 2, reason: '電話番号重複', name: row[nameIdx] });
    } 
    // 2. 店名＋住所による完全一致
    else if (name && address && seenNameAddrs.has(nameAddrKey)) {
      isDuplicate = true;
      duplicateCandidates.push({ index: idx + 2, reason: '店名・住所重複', name: row[nameIdx] });
    }
    
    if (!isDuplicate) {
      if (phone) seenPhones.set(phone, idx);
      if (name && address) seenNameAddrs.set(nameAddrKey, idx);
      rowsToKeep.push(row);
    }
  });
  
  console.log(`重複排除完了: 保持=${rowsToKeep.length}件, 排除=${duplicateCandidates.length}件`);
}

// =====================================================================
// コムデスク投入用CSV生成機能
// =====================================================================
/**
 * コムデスク投入フォーマットに合わせてCSVファイルを生成しダウンロードURLを表示
 */
function exportForComdesk() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const header = data[0];
  const rows = data.slice(1);
  
  const prefIdx = header.indexOf('都道府県');
  const cityIdx = header.indexOf('市区町村');
  const genreIdx = header.indexOf('ジャンル');
  const nameIdx = header.indexOf('店名');
  const phoneIdx = header.indexOf('電話番号');
  
  // 都道府県_市区町村_業種 ごとにグループ化
  const groups = {};
  
  rows.forEach(row => {
    const pref = String(row[prefIdx] || '').trim();
    const city = String(row[cityIdx] || '').trim();
    const genre = String(row[genreIdx] || '').trim();
    const name = String(row[nameIdx] || '').trim();
    const phone = normalizePhoneNumber(row[phoneIdx]);
    
    // 電話番号が無い店舗は架電できないためスキップ
    if (!phone) return;
    
    const key = `${pref}_${city}_${genre}`;
    if (!groups[key]) {
      groups[key] = [];
    }
    
    // コムデスク用フォーマット行作成（必要カラムのみ抽出）
    groups[key].push([name, phone, pref + city, genre]);
  });
  
  // 各グループごとにCSV文字列を生成
  for (const [filename, records] of Object.entries(groups)) {
    if (records.length === 0) continue;
    
    let csvContent = '店舗名,電話番号,住所,業種\r\n';
    records.forEach(rec => {
      const escapedRow = rec.map(val => `"${String(val).replace(/"/g, '""')}"`);
      csvContent += escapedRow.join(',') + '\r\n';
    });
    
    // Googleドライブに出力保存
    const folder = getOrCreateFolder('Comdesk_CSV_Exports');
    const file = folder.createFile(filename + '.csv', csvContent, MimeType.CSV);
    console.log(`CSV出力完了: ${file.getName()} (URL: ${file.getUrl()})`);
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}
