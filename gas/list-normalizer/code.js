/**
 * HD事業部 CSV統合・正規化・取得漏れ分析・除外理由管理 GAS テンプレート
 *
 * CSV取込後のデータを削除せず、営業対象・確認対象・除外理由を列として残します。
 * Googleマップ拡張機能の取得ステータス列がある場合は、取得漏れ分析にも利用します。
 */

// =====================================================================
// 設定
// =====================================================================
const RAW_SHEET_NAME = '00_IMPORTED';
const NORMALIZED_SHEET_NAME = '01_NORMALIZED';
const DUPLICATE_SHEET_NAME = '02_DUPLICATE_CHECK';
const CHAIN_SHEET_NAME = '03_CHAIN_CHECK';
const FACILITY_SHEET_NAME = '04_FACILITY_CHECK';
const COMDESK_SHEET_NAME = 'コムデスク投入用';
const PROCESS_LOG_SHEET_NAME = 'PROCESS_LOG';

const STANDARD_HEADERS = [
  '店名',
  'ジャンル',
  '検索ジャンル',
  '取得元ジャンル',
  '都道府県',
  '市区町村',
  '住所',
  '電話番号',
  '定休日',
  '営業日',
  '営業開始A',
  '営業終了A',
  '営業開始B',
  '営業終了B',
  '営業時間原文',
  'URL',
  'HP有無',
  '媒体',
  '取得元URL',
  '取得日時',
  '取得ステータス',
  '除外理由',
  '詳細取得リトライ回数',
  '一覧取得順'
];

const ANALYSIS_HEADERS = [
  '正規化電話番号',
  '正規化店名',
  '正規化ジャンル',
  '住所判定',
  'エリア判定',
  'エリア判定理由',
  '基本データ判定',
  '基本データ除外理由',
  '重複判定',
  '重複理由',
  'チェーン判定',
  'チェーン判定理由',
  '施設判定',
  '施設判定理由',
  '営業対象判定',
  '営業対象除外理由'
];

const FINAL_HEADERS = STANDARD_HEADERS.concat(ANALYSIS_HEADERS);

const OLD_GENRE_MAP = {
  '食堂': '定食・食堂',
  'そばうどん': '蕎麦・うどん',
  'バー': 'Bar',
  'パン': 'パン屋',
  'お弁当': '弁当',
  '中華料理': '中華',
  '韓国料理': '韓国',
  'イタリアン': '洋食',
  'フレンチ': '洋食',
  'カレー': '洋食',
  'タピオカ': 'スイーツ',
  'とんかつ': '和食',
  '沖縄料理': '和食'
};

const HD_TARGET_GENRES = [
  'カフェ',
  '定食・食堂',
  '蕎麦・うどん',
  'Bar',
  'パン屋',
  '弁当',
  '中華',
  '韓国',
  '洋食',
  'スイーツ',
  '和食',
  '居酒屋',
  'ラーメン',
  '焼肉',
  '寿司',
  'レストラン'
];

const CAFE_KEYWORDS = [
  'カフェ',
  'Cafe',
  'CAFE',
  'cafe',
  '喫茶',
  '珈琲',
  'コーヒー',
  'coffee',
  'Coffee',
  'COFFEE',
  'コーヒーショップ',
  'カフェテリア',
  'ドッグカフェ',
  'コーヒー焙煎所'
];

const FACILITY_EXCLUDE_KEYWORDS = [
  'イオンモール',
  'イオン',
  'AEON',
  'ららぽーと',
  'アリオ',
  'パルコ',
  'PARCO',
  'ルミネ',
  'LUMINE',
  'アトレ',
  'エキュート',
  'マルイ',
  'OIOI',
  '百貨店',
  '高島屋',
  '伊勢丹',
  '三越',
  'そごう',
  '大丸',
  '松坂屋',
  '阪急',
  '近鉄',
  'ショッピングセンター',
  'ショッピングモール',
  'アウトレット',
  'フードコート',
  '駅ビル',
  'ホテル',
  '病院',
  '大学',
  '学校',
  'スーパー',
  'ホームセンター',
  'ドン・キホーテ',
  'ドンキ',
  'ヨーカドー',
  'イトーヨーカドー',
  'アピタ',
  'ピアゴ',
  '西友',
  'ライフ',
  'マックスバリュ'
];

const FACILITY_REVIEW_KEYWORDS = [
  'ビル',
  'プラザ',
  'タワー',
  'センター',
  'テナント',
  'B1F',
  '1F',
  '2F',
  '3F',
  '4F',
  '5F',
  '階',
  '地下'
];

// =====================================================================
// 一括処理
// =====================================================================
function executeAllProcesses() {
  importCSVFiles();
  executeNormalizeAndValidate();
  executeDuplicateCheck();
  executeChainCheck();
  executeFacilityCheck();
  executeSplitSheets();
  executeGenerateComdeskCsv();
  executeProcessLogSummary();
}

// =====================================================================
// CSV取込
// =====================================================================
function importCSVFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (importSheet && importSheet.getLastRow() > 1) return;

  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() === RAW_SHEET_NAME) return;

  const values = activeSheet.getDataRange().getValues();
  if (values.length === 0 || values[0].length === 0) {
    throw new Error('取込対象データが見つかりません。CSVをシートに貼り付けてから実行してください。');
  }

  writeRowsToSheet(RAW_SHEET_NAME, values[0], values.slice(1));
}

// =====================================================================
// 正規化・基本判定
// =====================================================================
function executeNormalizeAndValidate() {
  const values = readPreferredSource_([RAW_SHEET_NAME]);
  const normalized = normalizeRows_(values);
  writeRowsToSheet(NORMALIZED_SHEET_NAME, FINAL_HEADERS, normalized);
}

function normalizeRows_(values) {
  if (values.length <= 1) return [];

  const sourceHeader = values[0];
  const rows = values.slice(1);

  return rows.map(row => {
    const obj = rowToObject_(sourceHeader, row);
    fillAddressParts_(obj);

    const normalizedName = normalizeName_(obj['店名']);
    const normalizedPhone = normalizePhoneNumber(obj['電話番号']);
    const normalizedGenre = normalizeSystemGenre(obj['ジャンル'], obj['検索ジャンル'], obj['取得元ジャンル'], obj['店名']);
    const addressStatus = judgeAddressStatus_(obj);
    const area = judgeAreaStatus(obj['住所'], obj['都道府県'], obj['市区町村']);
    const basic = judgeBasicData_(obj, normalizedGenre, normalizedPhone, addressStatus);

    obj['ジャンル'] = normalizedGenre || obj['ジャンル'] || '';
    obj['正規化電話番号'] = normalizedPhone;
    obj['正規化店名'] = normalizedName;
    obj['正規化ジャンル'] = normalizedGenre;
    obj['住所判定'] = addressStatus.status;
    obj['エリア判定'] = area.status;
    obj['エリア判定理由'] = area.reason;
    obj['基本データ判定'] = basic.status;
    obj['基本データ除外理由'] = mergeReasons_(obj['除外理由'], basic.reason);
    obj['重複判定'] = obj['重複判定'] || '';
    obj['重複理由'] = obj['重複理由'] || '';
    obj['チェーン判定'] = obj['チェーン判定'] || '';
    obj['チェーン判定理由'] = obj['チェーン判定理由'] || '';
    obj['施設判定'] = obj['施設判定'] || '';
    obj['施設判定理由'] = obj['施設判定理由'] || '';
    obj['営業対象判定'] = obj['営業対象判定'] || '';
    obj['営業対象除外理由'] = obj['営業対象除外理由'] || '';

    return FINAL_HEADERS.map(header => obj[header] || '');
  });
}

function judgeBasicData_(obj, normalizedGenre, normalizedPhone, addressStatus) {
  const reasons = [];
  const name = text_(obj['店名']);
  const fetchStatus = text_(obj['取得ステータス']);

  if (!name) reasons.push('店名なし');
  if (addressStatus.status !== '住所あり') reasons.push(addressStatus.reason);
  if (!normalizedPhone) reasons.push('電話番号なし');
  if (!normalizedGenre) reasons.push('ジャンル確認');
  if (normalizedGenre && !isValidHdGenre(normalizedGenre)) reasons.push('ジャンル確認');
  if (fetchStatus === '失敗') reasons.push('取得失敗');

  if (reasons.length === 0) return { status: '対象', reason: '' };
  return { status: '確認対象', reason: uniqueTexts_(reasons).join(' / ') };
}

// =====================================================================
// 重複判定
// =====================================================================
function executeDuplicateCheck() {
  const values = readPreferredSource_([NORMALIZED_SHEET_NAME, RAW_SHEET_NAME]);
  if (values.length <= 1) {
    writeRowsToSheet(DUPLICATE_SHEET_NAME, FINAL_HEADERS, []);
    return;
  }

  const header = values[0];
  const rows = values.slice(1).map(row => rowToObject_(header, row));
  const seenPhones = new Map();
  const seenNameAddresses = new Map();

  rows.forEach((obj, idx) => {
    const phone = text_(obj['正規化電話番号']) || normalizePhoneNumber(obj['電話番号']);
    const name = text_(obj['正規化店名']) || normalizeName_(obj['店名']);
    const address = normalizeAddressText(obj['住所']);
    const nameAddressKey = [name, address].filter(Boolean).join('_');
    const reasons = [];

    if (phone && seenPhones.has(phone)) {
      reasons.push(`電話番号重複: ${seenPhones.get(phone) + 2}行目`);
    } else if (phone) {
      seenPhones.set(phone, idx);
    }

    if (nameAddressKey && seenNameAddresses.has(nameAddressKey)) {
      reasons.push(`店名・住所重複: ${seenNameAddresses.get(nameAddressKey) + 2}行目`);
    } else if (nameAddressKey) {
      seenNameAddresses.set(nameAddressKey, idx);
    }

    obj['重複判定'] = reasons.length ? '重複' : 'ユニーク';
    obj['重複理由'] = reasons.join(' / ');
  });

  writeRowsToSheet(DUPLICATE_SHEET_NAME, FINAL_HEADERS, objectsToRows_(rows));
}

// =====================================================================
// チェーン判定
// =====================================================================
function executeChainCheck() {
  const values = readPreferredSource_([DUPLICATE_SHEET_NAME, NORMALIZED_SHEET_NAME, RAW_SHEET_NAME]);
  if (values.length <= 1) {
    writeRowsToSheet(CHAIN_SHEET_NAME, FINAL_HEADERS, []);
    return;
  }

  const header = values[0];
  const rows = values.slice(1).map(row => rowToObject_(header, row));
  const chainKeywords = loadChainKeywords_();

  rows.forEach(obj => {
    const name = text_(obj['店名']);
    const matched = chainKeywords.find(keyword => name.indexOf(keyword) !== -1);
    obj['チェーン判定'] = matched ? 'チェーン店' : '単独店';
    obj['チェーン判定理由'] = matched ? `チェーンキーワード一致: ${matched}` : '';
  });

  writeRowsToSheet(CHAIN_SHEET_NAME, FINAL_HEADERS, objectsToRows_(rows));
}

function loadChainKeywords_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MASTER_CHAIN');
  if (!sheet || sheet.getLastRow() < 1) return [];

  return sheet.getDataRange().getValues()
    .flat()
    .map(text_)
    .filter(Boolean);
}

// =====================================================================
// 施設判定
// =====================================================================
function executeFacilityCheck() {
  const values = readPreferredSource_([CHAIN_SHEET_NAME, DUPLICATE_SHEET_NAME, NORMALIZED_SHEET_NAME, RAW_SHEET_NAME]);
  if (values.length <= 1) {
    writeRowsToSheet(FACILITY_SHEET_NAME, FINAL_HEADERS, []);
    return;
  }

  const header = values[0];
  const rows = values.slice(1).map(row => rowToObject_(header, row));

  rows.forEach(obj => {
    const facility = judgeFacilityStatus(obj['店名'], obj['住所']);
    const final = judgeSalesTarget_(obj, facility);

    obj['施設判定'] = facility.status;
    obj['施設判定理由'] = facility.reason;
    obj['営業対象判定'] = final.status;
    obj['営業対象除外理由'] = final.reason;
  });

  writeRowsToSheet(FACILITY_SHEET_NAME, FINAL_HEADERS, objectsToRows_(rows));
}

function judgeSalesTarget_(obj, facility) {
  const reasons = [];
  let hasReviewReason = false;
  let hasExcludeReason = false;

  if (text_(obj['取得ステータス']) === '失敗') {
    hasExcludeReason = true;
    reasons.push('取得失敗');
  }
  if (text_(obj['除外理由'])) {
    hasExcludeReason = true;
    reasons.push(text_(obj['除外理由']));
  }
  if (text_(obj['基本データ判定']) !== '対象') {
    hasReviewReason = true;
    reasons.push(text_(obj['基本データ除外理由']) || '基本データ確認');
  }
  if (text_(obj['エリア判定']) === '判定不可') {
    hasReviewReason = true;
    reasons.push(text_(obj['エリア判定理由']) || 'エリア判定不可');
  } else if (text_(obj['エリア判定']) === 'エリア外') {
    hasExcludeReason = true;
    reasons.push(text_(obj['エリア判定理由']) || 'エリア外');
  }
  if (text_(obj['重複判定']) === '重複') {
    hasExcludeReason = true;
    reasons.push(text_(obj['重複理由']) || '重複');
  }
  if (text_(obj['チェーン判定']) === 'チェーン店') {
    hasExcludeReason = true;
    reasons.push(text_(obj['チェーン判定理由']) || 'チェーン店');
  }
  if (facility.status === '確認対象') {
    hasReviewReason = true;
    reasons.push(facility.reason || '施設確認対象');
  } else if (facility.status === '除外') {
    hasExcludeReason = true;
    reasons.push(facility.reason || '施設除外');
  }

  const compactReasons = uniqueTexts_(reasons);
  if (compactReasons.length === 0) return { status: '対象', reason: '' };
  if (hasReviewReason && !hasExcludeReason) {
    return { status: '確認対象', reason: compactReasons.join(' / ') };
  }
  return { status: '除外', reason: compactReasons.join(' / ') };
}

// =====================================================================
// タブ分け・取得漏れ分析
// =====================================================================
function executeSplitSheets() {
  const values = readPreferredSource_([FACILITY_SHEET_NAME, CHAIN_SHEET_NAME, DUPLICATE_SHEET_NAME, NORMALIZED_SHEET_NAME, RAW_SHEET_NAME]);
  if (values.length <= 1) return;

  const header = values[0];
  const rows = values.slice(1).map(row => rowToObject_(header, row));
  const buckets = buildSplitBuckets_(rows);

  Object.keys(buckets).forEach(sheetName => {
    writeRowsToSheet(sheetName, FINAL_HEADERS, objectsToRows_(buckets[sheetName]));
  });

  writeSummarySheet_(rows);
  writeExclusionAnalysisSheets_(rows);
}

function buildSplitBuckets_(rows) {
  const buckets = {
    '営業対象': [],
    '確認対象': [],
    '重複': [],
    'チェーン店除外': [],
    'ビル管理除外': [],
    'エリア外': [],
    '住所未取得': [],
    '電話番号なし': [],
    'ジャンル確認': [],
    '取得失敗': []
  };

  rows.forEach(obj => {
    const target = text_(obj['営業対象判定']);
    const reasonText = [
      obj['営業対象除外理由'],
      obj['基本データ除外理由'],
      obj['除外理由'],
      obj['施設判定理由'],
      obj['エリア判定理由']
    ].map(text_).join(' / ');

    if (target === '対象') buckets['営業対象'].push(obj);
    if (target === '確認対象') buckets['確認対象'].push(obj);
    if (text_(obj['重複判定']) === '重複') buckets['重複'].push(obj);
    if (text_(obj['チェーン判定']) === 'チェーン店') buckets['チェーン店除外'].push(obj);
    if (text_(obj['施設判定']) === '除外') buckets['ビル管理除外'].push(obj);
    if (text_(obj['エリア判定']) === 'エリア外') buckets['エリア外'].push(obj);
    if (text_(obj['住所判定']) !== '住所あり' || reasonText.indexOf('住所未取得') !== -1) buckets['住所未取得'].push(obj);
    if (!text_(obj['正規化電話番号']) && !text_(obj['電話番号'])) buckets['電話番号なし'].push(obj);
    if (reasonText.indexOf('ジャンル確認') !== -1 || !isValidHdGenre(obj['正規化ジャンル'])) buckets['ジャンル確認'].push(obj);
    if (text_(obj['取得ステータス']) === '失敗' || reasonText.indexOf('取得失敗') !== -1) buckets['取得失敗'].push(obj);
  });

  return buckets;
}

function writeSummarySheet_(rows) {
  const summaryHeader = ['媒体', '都道府県', '市区町村', '検索ジャンル', 'ジャンル', '取得ステータス', '除外理由', '件数'];
  const counts = new Map();

  rows.forEach(obj => {
    const keyValues = [
      text_(obj['媒体']),
      text_(obj['都道府県']),
      text_(obj['市区町村']),
      text_(obj['検索ジャンル']),
      text_(obj['ジャンル'] || obj['正規化ジャンル']),
      text_(obj['取得ステータス']) || '未設定',
      mergeReasons_(obj['除外理由'], obj['営業対象除外理由'])
    ];
    const key = keyValues.join('\u0001');
    counts.set(key, (counts.get(key) || 0) + 1);
  });

  const summaryRows = Array.from(counts.entries()).map(([key, count]) => key.split('\u0001').concat([count]));
  writeRowsToSheet('取得状況サマリー', summaryHeader, summaryRows);
}

function writeExclusionAnalysisSheets_(rows) {
  const analysisHeader = [
    '店名',
    'ジャンル',
    '検索ジャンル',
    '取得元ジャンル',
    '都道府県',
    '市区町村',
    '住所',
    '電話番号',
    'URL',
    '媒体',
    '取得ステータス',
    '除外理由',
    '営業対象判定',
    '営業対象除外理由'
  ];

  const exclusionRows = rows.filter(obj => {
    return text_(obj['除外理由']) || text_(obj['営業対象判定']) === '除外' || text_(obj['取得ステータス']) === '失敗';
  });
  const genreMismatchRows = rows.filter(obj => {
    const searchGenre = text_(obj['検索ジャンル']);
    const normalizedGenre = text_(obj['正規化ジャンル']);
    return normalizedGenre === 'その他' || !isValidHdGenre(normalizedGenre) || (searchGenre && normalizedGenre && searchGenre !== normalizedGenre);
  });
  const missingAddressRows = rows.filter(obj => text_(obj['住所判定']) !== '住所あり');
  const detailFailureRows = rows.filter(obj => {
    const reasons = [obj['除外理由'], obj['営業対象除外理由']].map(text_).join(' / ');
    return text_(obj['取得ステータス']) === '失敗' || reasons.indexOf('詳細取得失敗') !== -1 || (!text_(obj['店名']) && text_(obj['URL']));
  });

  writeRowsToSheet('除外理由別一覧', analysisHeader, toAnalysisRows_(exclusionRows, analysisHeader));
  writeRowsToSheet('ジャンル不一致一覧', analysisHeader, toAnalysisRows_(genreMismatchRows, analysisHeader));
  writeRowsToSheet('住所未取得一覧', analysisHeader, toAnalysisRows_(missingAddressRows, analysisHeader));
  writeRowsToSheet('詳細取得失敗一覧', analysisHeader, toAnalysisRows_(detailFailureRows, analysisHeader));
}

function toAnalysisRows_(rows, header) {
  return rows.map(obj => header.map(key => obj[key] || ''));
}

// =====================================================================
// コムデスク投入用
// =====================================================================
function executeGenerateComdeskCsv() {
  const values = readPreferredSource_([FACILITY_SHEET_NAME, CHAIN_SHEET_NAME, DUPLICATE_SHEET_NAME, NORMALIZED_SHEET_NAME, RAW_SHEET_NAME]);
  if (values.length <= 1) {
    writeRowsToSheet(COMDESK_SHEET_NAME, ['店舗名', '電話番号', '住所', '業種'], []);
    return;
  }

  const header = values[0];
  const rows = values.slice(1).map(row => rowToObject_(header, row));
  const targets = rows.filter(isComdeskTarget_);
  const outputRows = targets.map(obj => [
    obj['店名'] || '',
    obj['正規化電話番号'] || normalizePhoneNumber(obj['電話番号']),
    obj['住所'] || [obj['都道府県'], obj['市区町村']].join(''),
    obj['正規化ジャンル'] || obj['ジャンル'] || ''
  ]);

  writeRowsToSheet(COMDESK_SHEET_NAME, ['店舗名', '電話番号', '住所', '業種'], outputRows);
}

function isComdeskTarget_(obj) {
  const phone = text_(obj['正規化電話番号']) || normalizePhoneNumber(obj['電話番号']);
  const genre = text_(obj['正規化ジャンル']) || text_(obj['ジャンル']);

  return text_(obj['営業対象判定']) === '対象' &&
    text_(obj['重複判定']) === 'ユニーク' &&
    text_(obj['チェーン判定']) === '単独店' &&
    text_(obj['施設判定']) === '対象' &&
    text_(obj['エリア判定']) === 'エリア内' &&
    !!phone &&
    !!text_(obj['住所']) &&
    isValidHdGenre(genre);
}

// 既存名との互換用
function executeGenerateSalesList() {
  executeGenerateComdeskCsv();
}

function exportForComdesk() {
  executeGenerateComdeskCsv();
}

// =====================================================================
// ログ
// =====================================================================
function executeProcessLogSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [
    '営業対象',
    '確認対象',
    '重複',
    'チェーン店除外',
    'ビル管理除外',
    'エリア外',
    '住所未取得',
    '電話番号なし',
    'ジャンル確認',
    '取得失敗',
    COMDESK_SHEET_NAME
  ];
  const counts = {};
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    counts[name] = sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;
  });

  const importedSheet = ss.getSheetByName(RAW_SHEET_NAME) || ss.getSheetByName(NORMALIZED_SHEET_NAME);
  appendProcessLog({
    processName: 'executeAllProcesses',
    importedCount: importedSheet ? Math.max(importedSheet.getLastRow() - 1, 0) : 0,
    targetCount: counts['営業対象'],
    reviewCount: counts['確認対象'],
    duplicateCount: counts['重複'],
    chainCount: counts['チェーン店除外'],
    facilityCount: counts['ビル管理除外'],
    areaOutCount: counts['エリア外'],
    missingAddressCount: counts['住所未取得'],
    missingPhoneCount: counts['電話番号なし'],
    genreReviewCount: counts['ジャンル確認'],
    fetchFailureCount: counts['取得失敗'],
    csvCount: counts[COMDESK_SHEET_NAME],
    message: '処理完了'
  });
}

function appendProcessLog(log) {
  const header = [
    '実行日時',
    '処理名',
    '取込件数',
    '営業対象件数',
    '確認対象件数',
    '重複件数',
    'チェーン除外件数',
    'ビル管理除外件数',
    'エリア外件数',
    '住所未取得件数',
    '電話番号なし件数',
    'ジャンル確認件数',
    '取得失敗件数',
    'CSV出力件数',
    'メッセージ'
  ];
  const row = [
    new Date(),
    log.processName || '',
    log.importedCount || 0,
    log.targetCount || 0,
    log.reviewCount || 0,
    log.duplicateCount || 0,
    log.chainCount || 0,
    log.facilityCount || 0,
    log.areaOutCount || 0,
    log.missingAddressCount || 0,
    log.missingPhoneCount || 0,
    log.genreReviewCount || 0,
    log.fetchFailureCount || 0,
    log.csvCount || 0,
    log.message || ''
  ];

  const sheet = getOrCreateSheet_(PROCESS_LOG_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(header);
  sheet.appendRow(row);
}

// =====================================================================
// 判定・正規化ヘルパー
// =====================================================================
function normalizeSystemGenre(genre, searchGenre, sourceGenre, storeName) {
  const rawGenre = text_(genre);
  const mapped = OLD_GENRE_MAP[rawGenre] || rawGenre;
  const search = text_(searchGenre);
  const haystack = [sourceGenre, storeName, rawGenre].map(text_).join(' ');

  if (search === 'カフェ' && CAFE_KEYWORDS.some(keyword => haystack.indexOf(keyword) !== -1)) {
    return 'カフェ';
  }

  return mapped;
}

function isValidHdGenre(genre) {
  return HD_TARGET_GENRES.indexOf(text_(genre)) !== -1;
}

function normalizePhoneNumber(phone) {
  if (!phone) return '';
  let clean = String(phone).replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 65248);
  }).replace(/[ー－―‐]/g, '-');
  clean = clean.replace(/[^0-9]/g, '');

  if (!clean) return '';
  if (clean.startsWith('090') || clean.startsWith('080') || clean.startsWith('070')) {
    return clean.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  }
  if (clean.startsWith('03') || clean.startsWith('06')) {
    return clean.replace(/(\d{2})(\d{4})(\d{4})/, '$1-$2-$3');
  }
  if (clean.length === 10) {
    return clean.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  }
  if (clean.length === 11) {
    return clean.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  }
  return clean;
}

function normalizeAddressText(address) {
  return text_(address)
    .replace(/(?:〒\d{3}-\d{4}\s*|日本、\s*)/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function normalizeName_(name) {
  return text_(name).replace(/\s+/g, '').trim();
}

function judgeAreaStatus(address, prefecture, city) {
  const normalizedAddress = normalizeAddressText(address);
  const pref = text_(prefecture);
  const cityText = text_(city);

  if (!normalizedAddress) return { status: '判定不可', reason: '住所未取得' };
  if (!pref || !cityText) return { status: '判定不可', reason: '都道府県または市区町村が空' };
  if (isSimpleAddress_(normalizedAddress, cityText)) return { status: '判定不可', reason: '簡易住所' };
  if (normalizedAddress.indexOf(pref) === -1) return { status: 'エリア外', reason: `都道府県不一致: ${pref}` };
  if (normalizedAddress.indexOf(cityText) === -1) return { status: 'エリア外', reason: `市区町村不一致: ${cityText}` };
  return { status: 'エリア内', reason: '' };
}

function judgeFacilityStatus(storeName, address) {
  const haystack = [storeName, address].map(text_).join(' ');
  const excludeKeyword = FACILITY_EXCLUDE_KEYWORDS.find(keyword => haystack.indexOf(keyword) !== -1);
  if (excludeKeyword) {
    return { status: '除外', reason: `完全除外キーワード一致: ${excludeKeyword}` };
  }

  const reviewKeyword = FACILITY_REVIEW_KEYWORDS.find(keyword => haystack.indexOf(keyword) !== -1);
  if (reviewKeyword) {
    return { status: '確認対象', reason: `確認対象キーワード一致: ${reviewKeyword}` };
  }

  return { status: '対象', reason: '' };
}

function judgeAddressStatus_(obj) {
  const address = normalizeAddressText(obj['住所']);
  const pref = text_(obj['都道府県']);
  const city = text_(obj['市区町村']);

  if (!address) return { status: '住所未取得', reason: '住所未取得' };
  if (!pref || !city) return { status: '住所未取得', reason: '都道府県または市区町村が空' };
  if (isSimpleAddress_(address, city)) return { status: '住所未取得', reason: '簡易住所' };
  return { status: '住所あり', reason: '' };
}

function isSimpleAddress_(address, city) {
  return !/[都道府県]/.test(address) || (city && address.indexOf(city) === -1);
}

function parseAddress(address) {
  if (!address) return { prefecture: '', city: '' };
  const cleanAddress = normalizeAddressText(address);
  const regex = /^((?:北海道|東京都|大阪府|京都府|.{2,3}県))?((?:.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村]))?(.+)?$/;
  const match = cleanAddress.match(regex);
  if (!match) return { prefecture: '', city: '' };
  return {
    prefecture: match[1] || '',
    city: match[2] || ''
  };
}

function fillAddressParts_(obj) {
  if (text_(obj['都道府県']) && text_(obj['市区町村'])) return;
  const parsed = parseAddress(obj['住所']);
  obj['都道府県'] = text_(obj['都道府県']) || parsed.prefecture;
  obj['市区町村'] = text_(obj['市区町村']) || parsed.city;
}

// =====================================================================
// シートIOヘルパー
// =====================================================================
function writeRowsToSheet(sheetName, header, rows) {
  const sheet = getOrCreateSheet_(sheetName);
  sheet.clearContents();

  const values = [header].concat(rows || []);
  if (values.length === 0) return;

  sheet.getRange(1, 1, values.length, header.length).setValues(values);
  sheet.setFrozenRows(1);
}

function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function readPreferredSource_(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 0; i < sheetNames.length; i++) {
    const sheet = ss.getSheetByName(sheetNames[i]);
    if (sheet && sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      return sheet.getDataRange().getValues();
    }
  }
  const activeSheet = ss.getActiveSheet();
  return activeSheet.getDataRange().getValues();
}

function rowToObject_(header, row) {
  const obj = {};
  FINAL_HEADERS.forEach(key => {
    obj[key] = '';
  });
  header.forEach((key, idx) => {
    obj[text_(key)] = row[idx] == null ? '' : row[idx];
  });
  return obj;
}

function objectsToRows_(objects) {
  return objects.map(obj => FINAL_HEADERS.map(header => obj[header] || ''));
}

function text_(value) {
  return value == null ? '' : String(value).trim();
}

function mergeReasons_() {
  return uniqueTexts_(Array.prototype.slice.call(arguments).join(' / ').split('/')).join(' / ');
}

function uniqueTexts_(values) {
  const seen = {};
  return values.map(text_).filter(value => {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

// =====================================================================
// APIキー・Drive互換ヘルパー
// =====================================================================
function getApiKey() {
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty('EXTERNAL_SEARCH_API_KEY');
  if (!apiKey) {
    console.warn('警告: EXTERNAL_SEARCH_API_KEY がスクリプトプロパティに設定されていません。');
  }
  return apiKey;
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}
