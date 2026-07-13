/**
 * 店舗管理システム - 統合処理モジュール Ver9.9.1 (飲食店・最強版 ＋ 超・重複排除機能)
 * (完成版CSVファイル全自動エクスポート ➔ 31項目新フォーマット ➔ BP検索確定値化 ➔ 時間の全角対応＆通し営業時間自動スライド ➔ 曜日除外処理 ➔ 表記揺れ完全吸収)
 *
 * Ver9.9.1 変更点（Ver9.9.0からの差分）
 * 1. Ver9.9.0で導入した「架電グループ（居酒屋系／カフェ系／その他飲食）への
 *    04_SALESタブ集約」を取り消し、Ver9.8.1のジャンル別タブ構成（20分割）に戻した。
 *    実際のコムデスク側ワークグループが「居酒屋」「焼き鳥」「焼肉」等、
 *    細かいジャンル単位でそのまま存在しているため、GAS側で先に集約してしまうと
 *    コムデスク投入時に再度手動で振り分け直す必要が生じ、かえって非効率だった。
 * 2. 「00_件数サマリー」に、ジャンル×時間帯（午前始／午後始）の件数クロス集計を追加。
 *    実際の架電運用は「午前始を11で検索」「午後始を15→16→17→18で検索」のように
 *    時間帯を絞って架電するため、架電を始める前に「今この時間帯はどのジャンルが
 *    何件あるか」を事前に把握し、居酒屋→お好み焼き→焼肉→焼き鳥のように
 *    フォールバック順を判断しやすくする。各行への時間帯タグ付与は行わず、
 *    あくまで件数サマリータブ内の集計表として追加するのみ（コムデスク投入データ
 *    自体には影響しない）。
 *    時間帯の列は「11時・15時・16時・17時・18時」のように固定せず、実際に
 *    01_営業対象に存在する午前始／午後始の値から動的に生成する。
 *
 * Ver9.8.1 変更点（Ver9.8.0からの差分）
 * 1. 05_コムデスク投入用（全ジャンル統合タブ）の自動生成を廃止。
 *    コムデスク側はジャンルごと（04_SALES_〇〇）にワークグループを分けて運用しているため、
 *    全ジャンル統合版は不要と判断。既存の05_コムデスク投入用タブは実行時に自動削除される。
 * 2. 検索ジャンル（Googleマップ拡張機能のCSVに含まれる列）をジャンル正規化の
 *    一般フォールバックとして使用するように変更。
 *
 * Ver9.8.0 変更点（Ver9.7.0からの差分）
 * 1. ジャンルを拡張機能側と統一（20種類）。「喫茶店」は独立ジャンルを廃止し「カフェ」へ統合。「美容院」は本リストの対象外（飲食店専用）。
 *    ※コムデスク側に「カフェ」「喫茶店」が別ワークグループとして残っている場合は、
 *      この統合が実態と合わなくなるため、その際は要連絡（HD_GENRE_MAPの
 *      「喫茶店」「喫茶」エントリを削除するだけで独立ジャンルに戻せる）。
 * 2. 最終シート構成を「コアタブのみ」に整理。実行のたびに大量生成されていた
 *    分析_*／確認_*／除外_*／CHAIN_CANDIDATE／PROCESS_LOG タブは自動生成を停止。
 *    残るのは処理パイプライン（01_NORMALIZED〜04_FACILITY_CHECK）と
 *    ワークフロー4区分（01_営業対象／02_確認対象／03_除外対象／04_取得失敗）、
 *    04_SALES_ジャンル別タブのみ。
 */

// =====================================================================
// ⚙️ 設定エリア（あなたのご自身のSerpAPIのAPIキーをここに貼り付けてください）
// =====================================================================
const SERPAPI_KEY = "df3a6435db2e72dfc9431ac13e752eaf5d4bc60e68a50617471bd09a0d73696f";

// =====================================================================
// スプレッドシートを開いた時に自動でオリジナルメニューを作る関数
// =====================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ HD店舗管理システム")
    .addItem("🚀 すべての一括処理を実行", "executeAllProcesses")
    .addSeparator()
    .addItem("📁 1. CSVを一括取り込み", "importCSVFiles")
    .addItem("2. 正規化・基本判定", "executeNormalizeAndValidate")
    .addItem("3. 重複判定", "executeDuplicateCheck")
    .addItem("4. チェーン判定", "executeChainCheck")
    .addItem("5. 施設判定", "executeFacilityCheck")
    .addItem("6. ワークフロー分類", "executeWorkflowGrouping")
    .addItem("7. タブ分け", "executeSplitSheets")
    .addItem("8. 04_SALES_ジャンル別タブ生成", "executeGenerateSalesGenreSheets")
    .addItem("9. 04_SALES_CSVをDrive出力", "executeExportSalesGenreCsvFiles")
    .addSeparator()
    .addItem("📊 件数サマリーを更新", "executeCountSummary")
    .addItem("🔧 チェーンマスタの不足キーワードを追加", "fixKnownChainMasterGaps")
    .addToUi();
}

// =====================================================================
// すべての処理を完全自動で連鎖させる一括実行関数
// =====================================================================
function executeAllProcesses() {
  importCSVFiles();
  executeNormalizeAndValidate();
  executeDuplicateCheck();
  executeChainCheck();
  executeFacilityCheck();
  executeWorkflowGrouping();
  executeSplitSheets();
  executeGenerateSalesGenreSheets();
  executeExportSalesGenreCsvFiles();
  const summary = executeCountSummary();
  SpreadsheetApp.getUi().alert(
    `HDリスト処理が完了しました。04_SALES_ジャンル別タブとCSVを確認してください。\n\n` +
    `営業対象: ${summary.total営業対象}件（ジャンル別・時間帯別の内訳は「00_件数サマリー」タブ参照）`
  );
}

// =====================================================================
// 処理0: 文字コード自動判別 ＆ 不可視ゴミ(BOM)自動抹殺インポート
// Ver9.8.1: 列数・列順が異なる複数CSV（食べログ/Googleマップ拡張機能や
// 旧フォーマットのCSVが混在）を安全に合体できるよう、最初のファイルの
// 列数に固定するのではなく、全ファイルのヘッダーを統合してから
// 各行を「列名基準」で位置合わせする方式に変更。
// =====================================================================
function importCSVFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = getOrCreateSheet(ss, "01_NORMALIZED");
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  let importFolder = getOrCreateFolder(parentFolder, "CSV投入フォルダ");
  let processedFolder = getOrCreateFolder(parentFolder, "処理済みフォルダ");
  const files = importFolder.getFilesByType(MimeType.CSV);

  // 1パス目: 全ファイルを読み込みつつ、ヘッダーの和集合（列名の統合リスト）を作る
  const parsedFiles = [];
  const unifiedHeader = [];
  const unifiedHeaderSet = new Set();

  while (files.hasNext()) {
    const file = files.next();
    try {
      ss.toast(`ファイル「${file.getName()}」を解析中...`, "📁 CSV一括取り込み");

      let blob = file.getBlob();
      let csvText = blob.getDataAsString("UTF-8").replace(/^\uFEFF/, "");

      if (!csvText.includes("店名")) {
        csvText = blob.getDataAsString("MS932").replace(/^\uFEFF/, "");
      }

      const parsedCsv = Utilities.parseCsv(csvText);
      if (parsedCsv.length === 0) { file.moveTo(processedFolder); continue; }

      const fileHeader = parsedCsv[0].map(h => String(h).replace(/^\uFEFF/, "").trim());
      const dataRows = parsedCsv.slice(1).filter(row => row.join("").trim() !== "");

      fileHeader.forEach(col => {
        if (col && !unifiedHeaderSet.has(col)) {
          unifiedHeaderSet.add(col);
          unifiedHeader.push(col);
        }
      });

      parsedFiles.push({ name: file.getName(), header: fileHeader, rows: dataRows });
      file.moveTo(processedFolder);
    } catch (e) {
      Logger.log(`[CSV取込エラー] ${file.getName()}: ${e.message}`);
      ss.toast(`「${file.getName()}」の読み込みに失敗しました: ${e.message}`, "⚠️ CSV取込エラー");
    }
  }

  if (parsedFiles.length === 0) return;

  // 2パス目: 統合ヘッダーの列位置に合わせて各行を並べ直す（無い列は空欄で埋める）
  const combinedData = [unifiedHeader];
  parsedFiles.forEach(pf => {
    const colIndexInFile = pf.header.map(col => unifiedHeader.indexOf(col));
    if (pf.header.length !== unifiedHeader.length) {
      ss.toast(`「${pf.name}」は列構成が異なるため位置合わせして取り込みます`, "📁 CSV一括取り込み");
    }
    pf.rows.forEach(row => {
      const alignedRow = new Array(unifiedHeader.length).fill("");
      row.forEach((value, i) => {
        const targetIdx = colIndexInFile[i];
        if (targetIdx !== undefined && targetIdx !== -1) {
          alignedRow[targetIdx] = value;
        }
      });
      combinedData.push(alignedRow);
    });
  });

  if (combinedData.length > 1) {
    targetSheet.clear();
    const range = targetSheet.getRange(1, 1, combinedData.length, combinedData[0].length);
    range.setNumberFormat("@");
    range.setValues(combinedData);
  }
}

// =====================================================================
// 処理1: 全列自動探索型・重複判定システム（🌟超強化版）
// =====================================================================
function executeDuplicateCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("01_NORMALIZED");
  const targetSheet = getOrCreateSheet(ss, "02_DUPLICATE_CHECK");

  if (!sourceSheet) return;
  const values = sourceSheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = values[0].map(h => String(h).replace(/^\uFEFF/, "").trim());
  const nameIdx = header.indexOf("店名");
  const phoneIdx = header.indexOf("電話番号");

  if (nameIdx === -1 || phoneIdx === -1) {
    SpreadsheetApp.getUi().alert("01_NORMALIZEDシート内に「店名」または「電話番号」の列が見つかりません。");
    return;
  }

  const dataRows = values.slice(1);
  const seenPhones = new Set();
  const seenNames = new Set();
  const outputRows = [];

  dataRows.forEach((row, index) => {
    if (index % 100 === 0) {
      ss.toast(`${index}件目の重複を判定中...`, "🔍 重複チェック");
    }
    const rawName = String(row[nameIdx]).normalize("NFC").trim();

    // 電話番号の表記揺れ（全角数字やハイフン）を完全に統一して数字のみにする
    let rawPhone = String(row[phoneIdx]).replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    rawPhone = rawPhone.replace(/[^\d]/g, "").trim();

    const cleanName = simplifyStoreName(rawName);
    let isDuplicate = false;
    let duplicateReason = "ユニーク";

    // 1. 電話番号での重複チェック（最優先：漢字とひらがなで名前が違っても電話番号が同じなら弾く）
    if (rawPhone !== "" && rawPhone.length >= 9) {
      if (seenPhones.has(rawPhone)) {
        isDuplicate = true;
        duplicateReason = "重複（電話番号一致）";
      } else {
        seenPhones.add(rawPhone);
      }
    }

    // 2. 店舗名での重複チェック（ひらがな・カタカナ・全角半角・記号無視の強力な表記揺れ吸収）
    if (!isDuplicate && cleanName !== "") {
      if (seenNames.has(cleanName)) {
        isDuplicate = true;
        duplicateReason = "重複（店舗名一致）";
      } else {
        seenNames.add(cleanName);
      }
    }

    const newRow = [...row];
    newRow.push(duplicateReason);
    outputRows.push(newRow);
  });

  targetSheet.clear();
  const newHeader = [...header];
  newHeader.push("重複判定");

  const finalOutput = [newHeader, ...outputRows];
  const range = targetSheet.getRange(1, 1, finalOutput.length, finalOutput[0].length);
  range.setNumberFormat("@");
  range.setValues(finalOutput);
}

// =====================================================================
// 処理2: 業種対応・チェーン判定システム
// MASTER_CHAIN 列構成: チェーン名 | 除外キーワード | 業種(飲食/美容/共通) | 除外対象 | メモ
// =====================================================================
function executeChainCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("MASTER_CHAIN");
  const sourceSheet = ss.getSheetByName("02_DUPLICATE_CHECK");
  const targetSheet = getOrCreateSheet(ss, "03_CHAIN_CHECK");

  if (!masterSheet || !sourceSheet) return;
  const masterValues = masterSheet.getDataRange().getValues();
  const chainMaster = [];
  for (let i = 1; i < masterValues.length; i++) {
    const chainName = String(masterValues[i][0]).normalize("NFC").trim();
    const keyword   = String(masterValues[i][1]).normalize("NFC").trim();
    const industry  = String(masterValues[i][2]).normalize("NFC").trim() || "共通"; // 業種: 飲食/美容/共通 (空=共通扱い)
    const isValid   = masterValues[i][3]; // 除外対象
    if (isValid && keyword !== "") {
      chainMaster.push({ chainName: chainName, keyword: keyword, industry: industry, normalizedKeyword: simplifyStoreName(keyword) });
    }
  }

  const sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length <= 1) return;

  const headerRow = sourceValues[0].map(h => String(h).replace(/^\uFEFF/, "").trim());
  const nameIdx = headerRow.indexOf("店名");

  if (nameIdx === -1) {
    SpreadsheetApp.getUi().alert("02_DUPLICATE_CHECKシート内に「店名」の列が見つかりません。");
    return;
  }

  const dataRows = sourceValues.slice(1);
  const brandCounter = {};
  dataRows.forEach(row => {
    const rawName = String(row[nameIdx]).normalize("NFC").trim();
    const cleanName = simplifyStoreName(rawName);
    if (cleanName) { brandCounter[cleanName] = (brandCounter[cleanName] || 0) + 1; }
  });

  const outputRows = [];
  dataRows.forEach((row, index) => {
    if (index % 100 === 0) {
      ss.toast(`${index}件目のチェーン店をスキャン中...`, "🔍 チェーン店判定");
    }
    const storeName = String(row[nameIdx]).normalize("NFC").trim();
    const cleanName = simplifyStoreName(storeName);
    const storeIndustry = getStoreIndustry(headerRow, row);
    let isChain = false;
    let matchedChainName = "";
    let chainReason = "";

    for (const master of chainMaster) {
      // 業種フィルタ: 共通は全業種に適用、それ以外は店舗業種と一致する行のみ
      if (master.industry !== "共通" && master.industry !== storeIndustry) continue;
      const rawMatch = storeName.includes(master.keyword);
      const normalizedMatch = master.normalizedKeyword && cleanName.includes(master.normalizedKeyword);
      if (rawMatch || normalizedMatch) {
        isChain = true; matchedChainName = master.chainName;
        chainReason = "マスタ合致: キーワード[" + master.keyword + "]" + (normalizedMatch && !rawMatch ? "（表記ゆれ吸収）" : ""); break;
      }
    }
    if (!isChain && cleanName && brandCounter[cleanName] >= 5) {
      isChain = true; matchedChainName = cleanName + "チェーン";
      chainReason = "自動検出: 出現数[" + brandCounter[cleanName] + "]件";
    }
    if (!isChain) { matchedChainName = ""; chainReason = "単独店確認"; }

    const newRow = [...row];
    newRow.push(isChain ? "チェーン店" : "単独店", matchedChainName, chainReason);
    outputRows.push(newRow);
  });

  targetSheet.clear();
  const newHeader = [...headerRow];
  newHeader.push("チェーン判定", "チェーン名", "チェーン理由");

  const finalOutput = [newHeader, ...outputRows];
  const range = targetSheet.getRange(1, 1, finalOutput.length, finalOutput[0].length);
  range.setNumberFormat("@");
  range.setValues(finalOutput);
}

// 店舗の業種を判定する（MASTER_CHAIN の業種フィルタに使用）
function getStoreIndustry(header, row) {
  // 明示的な業種列があればそれを優先
  const explicit = getRowValueByHeader(header, row, "業種");
  if (explicit === "美容" || explicit === "飲食") return explicit;

  // 媒体名にビューティー系ワードが含まれていれば美容
  const media = getRowValueByHeader(header, row, "媒体");
  if (media && (media.includes("ビューティー") || media.includes("BEAUTY") || media.includes("Beauty"))) return "美容";

  return "飲食";
}

// =====================================================================
// 二次元配列を100%安全なCSV文字列に変換するプロ仕様コンバーター
// =====================================================================
function convertArrayToCsvText(array) {
  return array.map(row => {
    return row.map(cell => {
      let str = String(cell).replace(/"/g, '""');
      if (str.includes(",") || str.includes("\n") || str.includes("\r") || str.includes('"')) {
        return '"' + str + '"';
      }
      return str;
    }).join(",");
  }).join("\r\n");
}

function formatToPureTime(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    let hh = Utilities.formatDate(val, "JST", "HH");
    let mm = Utilities.formatDate(val, "JST", "mm");
    let hourInt = parseInt(hh, 10);
    return (mm === "00" || mm === "") ? String(hourInt) : hourInt + ":" + mm;
  }
  let str = String(val).trim();
  if (str === "") return "";
  const timeMatch = str.match(/(\d{1,2}):(\d{2})/);
  if (timeMatch) {
    let hourInt = parseInt(timeMatch[1], 10);
    let minStr = timeMatch[2];
    return minStr === "00" ? String(hourInt) : hourInt + ":" + minStr;
  }
  const numMatch = str.match(/^(\d{1,2})$/);
  return numMatch ? String(parseInt(numMatch[1], 10)) : str;
}

function parseAddressDetails(fullAddress) {
  let pcode = "", pref = "", addr1 = fullAddress.trim();
  const pcodeMatch = addr1.match(/〒?(\d{3}-\d{4})/);
  if (pcodeMatch) { pcode = pcodeMatch[1]; addr1 = addr1.replace(/〒?\d{3}-\d{4}\s*/, ""); }
  addr1 = addr1.replace(/^日本、\s*/, "");
  const prefMatch = addr1.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))/);
  if (prefMatch) { pref = prefMatch[1]; addr1 = addr1.replace(pref, ""); }
  return { pcode: pcode, pref: pref, addr1: addr1 };
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); }
  return sheet;
}

// 🌟【超強化版】店舗名の表記揺れを限界まで無くして統一化するロジック
function simplifyStoreName(name) {
  if (!name) return "";

  // 1. 全角半角・大文字小文字を統一（「ＡＢＣ」→「abc」などに変換）
  let n = name.normalize("NFKC").toLowerCase();

  // 2. ひらがなをカタカナに強制変換（例「まつもと」と「マツモト」を同一視）
  n = n.replace(/[\u3041-\u3096]/g, function(match) {
    return String.fromCharCode(match.charCodeAt(0) + 0x60);
  });

  // 3. 記号・スペース・かっこ類をすべて完全に削除（「A B C」と「ABC」を同一視）
  n = n.replace(/[\s ・、。，．・！？!?()（）【】\[\]「」『』_－\-〜~]/g, "");

  // 4. 重複判定の邪魔になる末尾の単語を消す
  n = n.replace(/(店|駅前店|北口店|南口店|東口店|西口店|インター店|SA店|PA店|飯能店|高麗川店|本店|支店|営業所)$/, "");

  return n.trim();
}

// =====================================================================
// ジャンル正規化・施設判定 共通定義
// =====================================================================
const GOOGLEMAP_OPTIONAL_HEADERS = ["検索ジャンル", "取得ステータス", "除外理由", "詳細取得リトライ回数", "一覧取得順", "取得元ジャンル", "架電グループ"];
const ANALYSIS_EXTRA_HEADERS = [
  "正規化電話番号",
  "正規化店名",
  "正規化ジャンル",
  "住所判定",
  "エリア判定",
  "エリア判定理由",
  "基本データ判定",
  "基本データ除外理由",
  "施設判定",
  "施設判定理由",
  "営業対象判定",
  "営業対象除外理由",
  "ワークフローグループ",
  "ワークフロー項目",
  "対応ステータス",
  "次アクション"
];

// 旧ジャンル表記 → 統一ジャンルへのマッピング
// 「喫茶店」は独立ジャンルを廃止し「カフェ」に統合（Ver9.8.0）
// ※コムデスク側で「カフェ」「喫茶店」を別ワークグループとして分けて運用している
//   場合は、この統合が実態と合わなくなる。その場合は下の2行
//   （"喫茶店": "カフェ", "喫茶": "カフェ",）を削除し、HD_TARGET_GENRESに
//   "喫茶店" を追加すれば独立ジャンルに戻せる。
const HD_GENRE_MAP = {
  "喫茶店": "カフェ",
  "喫茶": "カフェ",
  "食堂": "定食・食堂",
  "定食": "定食・食堂",
  "料理旅館": "和食",
  "旅館": "和食",
  "ビュッフェ・バイキング": "和食",
  "ビュッフェ": "和食",
  "バイキング": "和食",
  "そばうどん": "蕎麦・うどん",
  "そば": "蕎麦・うどん",
  "蕎麦": "蕎麦・うどん",
  "うどん": "蕎麦・うどん",
  "バー": "Bar",
  "軽食店": "スナック",
  "パン": "パン屋",
  "ベーカリー": "パン屋",
  "お弁当": "テイクアウト専門店",
  "弁当屋": "テイクアウト専門店",
  "中華料理": "中華",
  "餃子": "中華",
  "韓国料理": "韓国",
  "イタリアン": "洋食",
  "フレンチ": "洋食",
  "カレー": "洋食",
  "ステーキ": "洋食",
  "ハンバーグ": "洋食",
  "パスタ": "洋食",
  "ピザ": "洋食",
  "タピオカ": "スイーツ",
  "ケーキ": "スイーツ",
  "和菓子": "スイーツ",
  "洋菓子": "スイーツ",
  "とんかつ": "和食",
  "沖縄料理": "和食",
  "しゃぶしゃぶ": "和食",
  "日本料理": "和食",
  "海鮮": "和食",
  "魚介": "和食",
  "うなぎ": "和食",
  "天ぷら": "和食",
  "焼鳥": "焼き鳥",
  "焼きとり": "焼き鳥",
  "ホルモン": "焼肉",
  "テイクアウト": "テイクアウト専門店"
};

// 最終的に採用する統一ジャンル（20種類・拡張機能側 HD_POPULAR_GENRES と完全一致させること）
const HD_TARGET_GENRES = [
  "カフェ",
  "居酒屋",
  "スナック",
  "Bar",
  "パン屋",
  "焼き鳥",
  "お好み焼き",
  "焼肉",
  "スイーツ",
  "中華",
  "ハンバーガー",
  "蕎麦・うどん",
  "寿司",
  "和食",
  "洋食",
  "定食・食堂",
  "韓国",
  "テイクアウト専門店",
  "ラーメン"
];

const CAFE_KEYWORDS = [
  "カフェ", "Cafe", "CAFE", "cafe", "喫茶", "珈琲", "コーヒー", "coffee", "Coffee", "COFFEE",
  "コーヒーショップ", "カフェテリア", "ドッグカフェ", "コーヒー焙煎所"
];

// 生ジャンルから統一ジャンルが決まらなかった場合に、検索ジャンルへのフォールバックへ
// 進む前に店名から優先的に拾うジャンルキーワード。具体的な複合語ほど誤爆しにくいので
// 先に判定させる（例:「中華そば」「タンメン」等はラーメンの一種であり「そば」単体より先に見る）。
// 実データで確認された誤爆（ラーメン屋・うどん屋がカフェ/居酒屋タブに混入等）を防ぐための対応。
const NAME_GENRE_PRIORITY_LIST = [
  ["支那そば", "ラーメン"],
  ["中華そば", "ラーメン"],
  ["油そば", "ラーメン"],
  ["まぜそば", "ラーメン"],
  ["つけ麺", "ラーメン"],
  ["拉麺", "ラーメン"],
  ["タンメン", "ラーメン"],
  ["ラーメン", "ラーメン"],
  ["らーめん", "ラーメン"],
  ["麺屋", "ラーメン"],
  ["麺や", "ラーメン"],
  ["讃岐", "蕎麦・うどん"],
  ["うどん", "蕎麦・うどん"],
  ["蕎麦", "蕎麦・うどん"],
  ["そば", "蕎麦・うどん"],
  ["お好み焼き", "お好み焼き"],
  ["もんじゃ", "お好み焼き"],
  ["焼き鳥", "焼き鳥"],
  ["焼鳥", "焼き鳥"],
  ["焼き肉", "焼肉"],
  ["焼肉", "焼肉"],
  ["スナック", "スナック"],
  ["寿司", "寿司"],
  ["鮨", "寿司"],
  ["すし", "寿司"],
  ["とんかつ", "和食"],
  ["とん㐂", "和食"],
  ["うなぎ", "和食"],
  ["鰻", "和食"],
  ["天ぷら", "和食"],
  ["中華", "中華"],
  ["餃子", "中華"],
  ["韓国", "韓国"],
  ["ハンバーガー", "ハンバーガー"],
  ["ベーカリー", "パン屋"],
  ["パン屋", "パン屋"],
  ["スイーツ", "スイーツ"],
  ["弁当", "テイクアウト専門店"]
];

// 店名から上記優先ジャンルのいずれかに一致するか調べる（一致しなければ空文字）
function findGenreFromStoreName(storeName) {
  const nameText = textValue(storeName);
  if (!nameText) return "";
  const hit = NAME_GENRE_PRIORITY_LIST.find(pair => nameText.indexOf(pair[0]) !== -1);
  return hit ? hit[1] : "";
}

const FACILITY_EXCLUDE_KEYWORDS = [
  "イオンモール", "イオン", "AEON", "ららぽーと", "アリオ", "パルコ", "PARCO", "ルミネ", "LUMINE",
  "アトレ", "エキュート", "マルイ", "OIOI", "百貨店", "高島屋", "伊勢丹", "三越", "そごう",
  "大丸", "松坂屋", "阪急", "近鉄", "ショッピングセンター", "ショッピングモール", "モール", "アウトレット",
  "フードコート", "駅ビル", "ホテル", "病院", "クリニック", "医院", "歯科", "大学", "学校", "スーパー", "ホームセンター",
  "ドン・キホーテ", "ドンキ", "ヨーカドー", "イトーヨーカドー", "アピタ", "ピアゴ", "西友", "ライフ", "マックスバリュ",
  "東武ストア", "タイヨー", "ランドローム", "生鮮市場", "食鮮館",
  "ビリヤード", "ボウリング", "カラオケ", "ゲームセンター", "パチンコ", "スロット", "雀荘", "麻雀",
  "スーパー銭湯", "温浴施設", "銭湯", "フィットネス", "スポーツジム", "ジム", "映画館", "ネットカフェ", "漫画喫茶",
  "農園", "時計", "時計店", "営業所", "総合公園", "都市公園",
  // ★追加: ゴルフ場・カントリークラブ（クラブハウス内レストランがGoogle/食べログ側で
  // 「ハンバーガー」等の飲食ジャンルとして誤って単独取得され、実際には独立した
  // 飲食店ではないゴルフ場全体が営業リストに混入する不具合を確認したため追加。
  // 香取市データで「香取カントリークラブ」がハンバーガータブに混入していたケース）
  // 「ゴルフ」を含む語（ゴルフ場・ゴルフクラブ・ゴルフ倶楽部・パブリックゴルフ等）は
  // この1語で幅広く弾けるため個別列挙は最小限にしている。なお「CC」「GC」
  // （Country Club/Golf Clubの略称のみの表記）は他の一般的な単語・型番等にも
  // 出現しうり誤爆リスクが高いため、あえて追加していない。もしこの略称表記の
  // ゴルフ場が混入するケースが見つかった場合は、店名を教えてもらえれば
  // 個別のチェーン名・施設名として追加する。
  "ゴルフ", "ゴルフ倶楽部", "ゴルフ練習場", "打ちっぱなし", "カントリークラブ", "カントリー倶楽部", "ゴルフ場", "ゴルフクラブ",
  // ★追加: 道の駅・川の駅など、単独の飲食店ではなく複数テナントが入る
  // 公共の休憩施設。同様の理由でクラブハウス的な扱いとして除外
  "道の駅", "川の駅",
  // ★追加: 天然温泉（日帰り温泉施設・入浴施設）。「スーパー銭湯」「温浴施設」「銭湯」と
  // 同じ理由（施設内レストランが単独の飲食店として誤って混入するため）で除外
  "天然温泉"
];

const FACILITY_REVIEW_KEYWORDS = ["ビル", "プラザ", "タワー", "センター", "テナント", "B1F", "1F", "2F", "3F", "4F", "5F", "階", "地下"];

// 店名の末尾に「店名（ジャンル）」のように既知ジャンル名が括弧書きで
// 紛れ込んでいる場合に取り除く（食べログ<title>由来のスクレイピング崩れ対策）。
// 括弧内が既知ジャンル語彙と一致する場合のみ除去し、店舗名の一部としての
// 括弧（支店名など）は誤って削らないようにする。
function stripGenreSuffixFromName(name) {
  const raw = textValue(name);
  const match = raw.match(/^(.*?)\s*[（(]([^）)]{1,20})[）)]\s*$/);
  if (!match) return raw;
  const inner = textValue(match[2]);
  const knownGenreWords = HD_TARGET_GENRES
    .concat(Object.keys(HD_GENRE_MAP))
    .concat(CAFE_KEYWORDS);
  if (knownGenreWords.indexOf(inner) !== -1) {
    return textValue(match[1]);
  }
  return raw;
}

function executeNormalizeAndValidate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("01_NORMALIZED");
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = normalizeHeaderRow(values[0]);
  const rows = values.slice(1).map(row => row.slice());
  ensureHeaderColumns(header, rows, GOOGLEMAP_OPTIONAL_HEADERS.concat(ANALYSIS_EXTRA_HEADERS));

  const indexes = buildHeaderIndex(header);
  rows.forEach(row => {
    fillPrefCityFromAddress(header, row);

    const rawStoreName = getRowValueByHeader(header, row, "店名");
    const storeName = stripGenreSuffixFromName(rawStoreName);
    if (storeName !== rawStoreName) setRowValueByHeader(header, row, "店名", storeName);

    const genre = getRowValueByHeader(header, row, "ジャンル");
    const searchGenre = getRowValueByHeader(header, row, "検索ジャンル");
    const sourceGenre = getRowValueByHeader(header, row, "取得元ジャンル");
    const address = getRowValueByHeader(header, row, "住所");
    const pref = getRowValueByHeader(header, row, "都道府県");
    const city = getRowValueByHeader(header, row, "市区町村");
    const phone = getRowValueByHeader(header, row, "電話番号");
    const fetchStatus = getRowValueByHeader(header, row, "取得ステータス");

    const normalizedGenre = normalizeSystemGenre(genre, searchGenre, sourceGenre, storeName);
    const normalizedPhone = normalizePhoneNumberForAnalysis(phone);
    const addressStatus = judgeAddressStatus(address, pref, city);
    const areaStatus = judgeAreaStatus(address, pref, city);
    const basicReasons = [];

    if (!textValue(storeName)) basicReasons.push("店名なし");
    if (addressStatus.status !== "住所あり") basicReasons.push(addressStatus.reason);
    if (!textValue(normalizedPhone)) basicReasons.push("電話番号なし");
    if (!textValue(normalizedGenre) || !isValidHdGenre(normalizedGenre)) basicReasons.push("ジャンル確認");
    if (textValue(fetchStatus) === "失敗") basicReasons.push("取得失敗");

    row[indexes["ジャンル"]] = normalizedGenre || genre;
    setRowValueByHeader(header, row, "正規化電話番号", normalizedPhone);
    setRowValueByHeader(header, row, "正規化店名", simplifyStoreName(storeName));
    setRowValueByHeader(header, row, "正規化ジャンル", normalizedGenre);
    setRowValueByHeader(header, row, "住所判定", addressStatus.status);
    setRowValueByHeader(header, row, "エリア判定", areaStatus.status);
    setRowValueByHeader(header, row, "エリア判定理由", areaStatus.reason);
    setRowValueByHeader(header, row, "基本データ判定", basicReasons.length === 0 ? "対象" : "確認対象");
    setRowValueByHeader(header, row, "基本データ除外理由", uniqueTextList(basicReasons).join(" / "));
  });

  writeRowsToExistingSheet(sheet, header, rows);
}

function executeFacilityCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("03_CHAIN_CHECK") || ss.getSheetByName("02_DUPLICATE_CHECK") || ss.getSheetByName("01_NORMALIZED");
  const targetSheet = getOrCreateSheet(ss, "04_FACILITY_CHECK");
  if (!sourceSheet) return;

  const values = sourceSheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = normalizeHeaderRow(values[0]);
  const rows = values.slice(1).map(row => row.slice());
  ensureHeaderColumns(header, rows, GOOGLEMAP_OPTIONAL_HEADERS.concat(ANALYSIS_EXTRA_HEADERS));

  rows.forEach(row => {
    const facility = judgeFacilityStatus(getRowValueByHeader(header, row, "店名"), getRowValueByHeader(header, row, "住所"), getRowValueByHeader(header, row, "ジャンル"));
    const sales = judgeSalesTargetStatus(header, row, facility);
    setRowValueByHeader(header, row, "施設判定", facility.status);
    setRowValueByHeader(header, row, "施設判定理由", facility.reason);
    setRowValueByHeader(header, row, "営業対象判定", sales.status);
    setRowValueByHeader(header, row, "営業対象除外理由", sales.reason);
  });

  writeRowsToExistingSheet(targetSheet, header, rows);
}

function executeWorkflowGrouping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("04_FACILITY_CHECK");
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = normalizeHeaderRow(values[0]);
  const rows = values.slice(1).map(row => row.slice());
  ensureHeaderColumns(header, rows, ["ワークフローグループ", "ワークフロー項目", "対応ステータス", "次アクション"]);

  rows.forEach(row => {
    const workflow = judgeWorkflowGroup(header, row);
    setRowValueByHeader(header, row, "ワークフローグループ", workflow.group);
    setRowValueByHeader(header, row, "ワークフロー項目", workflow.item);
    setRowValueByHeader(header, row, "対応ステータス", workflow.status);
    setRowValueByHeader(header, row, "次アクション", workflow.nextAction);
  });

  writeRowsToExistingSheet(sheet, header, rows);
}

// =====================================================================
// 処理7: タブ分け（Ver9.8.0：コアの4区分のみ生成。分析/確認の細分化タブは廃止）
// =====================================================================
function executeSplitSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("04_FACILITY_CHECK") || ss.getSheetByName("03_CHAIN_CHECK");
  if (!sourceSheet) return;

  const values = sourceSheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = normalizeHeaderRow(values[0]);
  const rows = values.slice(1);
  const buckets = {
    "01_営業対象": [],
    "02_確認対象": [],
    "03_除外対象": [],
    "04_取得失敗": []
  };

  rows.forEach(row => {
    const workflowGroup = getRowValueByHeader(header, row, "ワークフローグループ") || judgeWorkflowGroup(header, row).group;
    if (buckets[workflowGroup]) buckets[workflowGroup].push(row);
  });

  Object.keys(buckets).forEach(name => {
    writeRowsToSheetByName(ss, name, header, buckets[name]);
  });
}

// =====================================================================
// 処理8: 01_営業対象 → 04_SALES_ジャンル別タブを生成
// =====================================================================
function executeGenerateSalesGenreSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("01_営業対象");

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("01_営業対象シートがありません。先に一括処理を実行してください。");
    return;
  }

  const values = sourceSheet.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = normalizeHeaderRow(values[0]);
  const rows = values.slice(1);
  const finalHeader = getComdeskHeader();

  const genreContainers = {};
  HD_TARGET_GENRES.forEach(genre => {
    genreContainers[genre] = [];
  });

  rows.forEach(row => {
    if (!isComdeskTargetRow(header, row)) return;
    const genre = getFinalHdGenre(header, row);
    if (!genre || !genreContainers[genre]) return;
    genreContainers[genre].push(buildComdeskRow(header, row));
  });

  const sheetsToDelete = ss.getSheets().filter(s => s.getName().startsWith("04_SALES_"));
  sheetsToDelete.forEach(s => {
    if (ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  HD_TARGET_GENRES.forEach(genre => {
    const genreRows = genreContainers[genre];
    if (!genreRows || genreRows.length === 0) return;
    ss.toast(`シート「04_SALES_${genre}」を生成中...`, "📊 ジャンル別タブ生成");
    writeRowsToExistingSheet(getOrCreateSheet(ss, "04_SALES_" + genre), finalHeader, genreRows);
  });

  // コムデスク側はジャンルごと（04_SALES_〇〇）にワークグループを分けて運用しているため、
  // 全ジャンル統合タブは不要。既存のものが残っていれば削除する。
  const oldComdeskSheet = ss.getSheetByName("05_コムデスク投入用");
  if (oldComdeskSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(oldComdeskSheet);
  }
}

function getFinalHdGenre(header, row) {
  const normalizedGenre = getRowValueByHeader(header, row, "正規化ジャンル");
  const genre = getRowValueByHeader(header, row, "ジャンル");
  const searchGenre = getRowValueByHeader(header, row, "検索ジャンル");
  const sourceGenre = getRowValueByHeader(header, row, "取得元ジャンル");
  const storeName = getRowValueByHeader(header, row, "店名");

  const finalGenre = normalizeSystemGenre(
    normalizedGenre || genre || searchGenre || sourceGenre,
    searchGenre,
    sourceGenre,
    storeName
  );

  return HD_TARGET_GENRES.indexOf(finalGenre) !== -1 ? finalGenre : "";
}

// =====================================================================
// 処理9: 04_SALES_ジャンル別タブ → 完成版CSVエクスポートフォルダへ保存
// =====================================================================
function executeExportSalesGenreCsvFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();
  const exportFolder = getOrCreateFolder(parentFolder, "完成版CSVエクスポート");
  const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (!sheetName.startsWith("04_SALES_")) return;

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return;

    const genre = sheetName.replace("04_SALES_", "");
    const areaName = detectAreaNameFromComdeskRows(values) || "ダウンロードリスト";
    const fileName = `【営業リスト】${areaName}_${genre}_${formattedDate}.csv`;

    ss.toast(`CSVファイル「${fileName}」をDriveへ保存中...`, "📂 CSV出力");

    const bom = "\uFEFF";
    const blob = Utilities.newBlob(bom + convertArrayToCsvText(values), "text/csv", fileName);

    const existingFiles = exportFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    exportFolder.createFile(blob);
  });
}

// =====================================================================
// 処理10: 件数サマリー（各タブ・ジャンル別の件数 ＋ ジャンル×時間帯クロス集計）
// =====================================================================
function executeCountSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const countRows = sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;
    return Math.max(sheet.getLastRow() - 1, 0);
  };

  const coreCounts = {
    "01_NORMALIZED（取込総数）": countRows("01_NORMALIZED"),
    "01_営業対象": countRows("01_営業対象"),
    "02_確認対象": countRows("02_確認対象"),
    "03_除外対象": countRows("03_除外対象"),
    "04_取得失敗": countRows("04_取得失敗")
  };

  const genreCounts = {};
  let total営業対象InGenres = 0;
  HD_TARGET_GENRES.forEach(genre => {
    const count = countRows("04_SALES_" + genre);
    genreCounts[genre] = count;
    total営業対象InGenres += count;
  });

  const summarySheet = getOrCreateSheet(ss, "00_件数サマリー");
  const rows = [["区分", "件数"]];
  rows.push(["── 全体 ──", ""]);
  Object.keys(coreCounts).forEach(key => rows.push([key, coreCounts[key]]));
  rows.push(["", ""]);
  rows.push(["── 04_SALES_ジャンル別 ──", ""]);
  HD_TARGET_GENRES.forEach(genre => rows.push(["04_SALES_" + genre, genreCounts[genre]]));
  rows.push(["", ""]);
  rows.push(["ジャンル別合計（01_営業対象と一致するはず）", total営業対象InGenres]);
  rows.push(["更新日時", Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss")]);

  summarySheet.clearContents();
  summarySheet.getRange(1, 1, rows.length, 2).setValues(rows);
  summarySheet.getRange(1, 1, 1, 2).setFontWeight("bold");

  // ★NEW(Ver9.9.1): ジャンル×時間帯（午前始／午後始）クロス集計を同じシートの下に追記。
  // 架電を始める前に「今この時間帯はどのジャンルが何件あるか」を事前に把握できるようにする。
  const crossTab = buildGenreTimeCrossTab();
  const crossTabStartRow = rows.length + 2;
  writeGenreTimeCrossTabToSheet(summarySheet, crossTab, crossTabStartRow);

  summarySheet.autoResizeColumns(1, Math.max(2, crossTab.columns.length + 2));

  // 00_件数サマリーを一番左に移動して見やすくする
  ss.setActiveSheet(summarySheet);
  ss.moveActiveSheet(1);

  ss.toast(`営業対象 ${coreCounts["01_営業対象"]}件（ジャンル別合計 ${total営業対象InGenres}件）`, "📊 件数サマリー更新完了");

  return { total営業対象: coreCounts["01_営業対象"], genreCounts, coreCounts };
}

// 01_営業対象からジャンル×時間帯（午前始／午後始）の件数クロス集計を作る。
// 実際に04_SALESタブへ出力される時刻（通し営業時間の自動スライド処理後の値）と
// 数字が一致するよう、buildComdeskRowの計算結果をそのまま使う。
// 時間帯の列は固定せず、実データに存在する時刻から動的に生成する。
function buildGenreTimeCrossTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("01_営業対象");
  const table = {};
  HD_TARGET_GENRES.forEach(g => { table[g] = {}; });

  const emptyResult = { columns: [], table, genres: HD_TARGET_GENRES };
  if (!sheet) return emptyResult;

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return emptyResult;

  const header = normalizeHeaderRow(values[0]);
  const morningHours = new Set();
  const afternoonHours = new Set();

  values.slice(1).forEach(row => {
    if (!isComdeskTargetRow(header, row)) return;
    const genre = getFinalHdGenre(header, row);
    if (!genre) return;

    const salesRow = buildComdeskRow(header, row);
    const openA = textValue(salesRow[27]); // 午前始
    const openB = textValue(salesRow[29]); // 午後始

    const morningHour = extractHourLabel(openA);
    if (morningHour) {
      morningHours.add(Number(morningHour));
      const key = "午前始" + morningHour + "時台";
      table[genre][key] = (table[genre][key] || 0) + 1;
    }

    const afternoonHour = extractHourLabel(openB);
    if (afternoonHour) {
      afternoonHours.add(Number(afternoonHour));
      const key = "午後始" + afternoonHour + "時台";
      table[genre][key] = (table[genre][key] || 0) + 1;
    }
  });

  const morningCols = Array.from(morningHours).sort((a, b) => a - b).map(h => "午前始" + h + "時台");
  const afternoonCols = Array.from(afternoonHours).sort((a, b) => a - b).map(h => "午後始" + h + "時台");

  return { columns: morningCols.concat(afternoonCols), table, genres: HD_TARGET_GENRES };
}

// "11" や "11:30" のような時刻文字列から時（hour）部分だけを取り出す
function extractHourLabel(timeText) {
  const t = textValue(timeText);
  if (!t) return "";
  const match = t.match(/^(\d{1,2})/);
  return match ? String(parseInt(match[1], 10)) : "";
}

// ジャンル×時間帯クロス集計を「00_件数サマリー」シートの指定行以降に書き込む
function writeGenreTimeCrossTabToSheet(sheet, crossTab, startRow) {
  const columns = crossTab.columns;

  sheet.getRange(startRow, 1).setValue("── ジャンル×時間帯（午前始／午後始）件数クロス集計 ──");
  sheet.getRange(startRow, 1).setFontWeight("bold");

  if (columns.length === 0) {
    sheet.getRange(startRow + 1, 1).setValue("対象データなし（営業開始時刻が未取得のため集計できません）");
    return;
  }

  const headerRow = ["ジャンル", "合計"].concat(columns);
  const dataRows = crossTab.genres
    .map(genre => {
      const genreRow = crossTab.table[genre] || {};
      const total = Object.values(genreRow).reduce((sum, n) => sum + n, 0);
      const cells = columns.map(col => genreRow[col] || 0);
      return [genre, total].concat(cells);
    })
    // 件数0のジャンル行は表示を省略して見やすくする（架電判断に不要なため）
    .filter(row => row[1] > 0);

  const output = [headerRow].concat(dataRows);
  sheet.getRange(startRow + 1, 1, output.length, headerRow.length).setValues(output);
  sheet.getRange(startRow + 1, 1, 1, headerRow.length).setFontWeight("bold");
}

// =====================================================================
// 処理11: チェーンマスタの既知の抜け漏れを自動追加
// 「storeName.includes(keyword)」の単純一致なので、キーワードが実際の
// 店名表記（カタカナ表記・支店名の付き方など）と1文字でも違うとスルー
// されてしまう。実データで漏れが確認できたチェーンをここに追記していく。
// 既存の行と重複するキーワードはスキップするので、何度実行しても安全。
// =====================================================================
function fixKnownChainMasterGaps() {
  // [チェーン名, キーワード, 業種, 有効フラグ, 備考]
  const KNOWN_GAPS = [
    ["スターバックス", "スターバックス", "飲食", true, "カタカナ表記対策（英語表記Starbucksのみでは実店名に一致しないため追加）"],
    ["コメダ珈琲店", "コメダ珈琲店", "飲食", true, "接頭辞なし表記対策（「珈琲所コメダ珈琲店」だと接頭辞なしの実店名に一致しないため追加）"],
    ["星乃珈琲店", "星乃珈琲店", "飲食", true, "未登録だったため新規追加"],
    ["マクドナルド", "マクドナルド", "飲食", true, "カタカナ表記対策（英語表記McDonald'sのみでは実店名に一致しないため追加）"],
    ["ドミノ・ピザ", "ドミノピザ", "飲食", true, "未登録だったため新規追加"],
    ["ドミノ・ピザ", "ドミノ・ピザ", "飲食", true, "中黒あり表記対策"],
    ["スシロー", "スシロー", "飲食", true, "回転寿司チェーン。木更津市データで未除外を確認したため追加"],
    ["はま寿司", "はま寿司", "飲食", true, "回転寿司チェーン。木更津市データで未除外を確認したため追加"],
    ["くら寿司", "くら寿司", "飲食", true, "回転寿司チェーン（「無添くら寿司」表記にも部分一致）。木更津市データで未除外を確認したため追加"],
    ["すし銚子丸", "銚子丸", "飲食", true, "回転寿司チェーン。木更津市データで未除外を確認したため追加"],
    ["元気寿司", "元気寿司", "飲食", true, "回転寿司チェーン。表記ゆれ対策のため予防的に追加"],
    ["回転寿司 やまと", "回転寿司 やまと", "飲食", true, "回転寿司チェーン。「やまと」単独だと一般的すぎて誤爆するため店名全体をキーワードに設定"],
    ["モスバーガー", "モスバーガー", "飲食", true, "ハンバーガーチェーン。木更津市データで未除外を確認したため追加"],
    ["バーガーキング", "BURGER KING", "飲食", true, "ハンバーガーチェーン（英語表記）。木更津市データで未除外を確認したため追加"],
    ["シャトレーゼ", "シャトレーゼ", "飲食", true, "洋菓子チェーン。木更津市データで未除外を確認したため追加"],
    ["ガスト", "ガスト", "飲食", true, "ファミレスチェーン。印西市データで未除外を確認したため追加"],
    ["ピザハット", "ピザハット", "飲食", true, "ピザ宅配チェーン。印西市データで未除外を確認したため追加"],
    ["カプリチョーザ", "カプリチョーザ", "飲食", true, "イタリアンファミレスチェーン。印西市データで未除外を確認したため追加"],
    ["ココス", "ココス", "飲食", true, "ファミレスチェーン。印西市データで未除外を確認したため追加"],
    ["サイゼリヤ", "サイゼリヤ", "飲食", true, "イタリアンファミレスチェーン。印西市データで未除外を確認したため追加"],
    ["ゆで太郎", "ゆで太郎", "飲食", true, "そば・うどんチェーン。木更津市データで未除外を確認したため追加"],
    ["夢庵", "夢庵", "飲食", true, "和食ファミレスチェーン。木更津市データで未除外を確認したため追加"],
    ["しゃぶ葉", "しゃぶ葉", "飲食", true, "しゃぶしゃぶチェーン。木更津市データで未除外を確認したため追加"],
    ["蟹工船", "蟹工船", "飲食", true, "居酒屋チェーン。木更津市データで未除外を確認したため追加"],
    ["ピザーラ", "ピザーラ", "飲食", true, "ピザ宅配チェーン。木更津市データで未除外を確認したため追加"],
    ["デニーズ", "デニーズ", "飲食", true, "ファミレスチェーン。木更津市データで未除外を確認したため追加"],

    // ↓ここから: MASTER_CHAINに英語表記のキーワードしか登録されておらず、
    // 実際の店名（日本語表記）に一致しないチェーンへの対策。
    // 例:「吉野家」の登録キーワードが「YOSHINOYA」のみだったため、
    // 実店名「吉野家 木更津駅前店」に一致せず未除外だった不具合を確認したため、
    // 主要チェーンの日本語表記キーワードをまとめて追加する。
    ["吉野家", "吉野家", "飲食", true, "日本語表記対策（登録済みキーワード「YOSHINOYA」が実店名に一致しないため追加）"],
    ["すき家", "すき家", "飲食", true, "日本語表記対策（登録済みキーワード「SUKIYA」が実店名に一致しないため追加）"],
    ["なか卯", "なか卯", "飲食", true, "日本語表記対策（登録済みキーワード「NAKAU」が実店名に一致しないため追加）"],
    ["ケンタッキー", "ケンタッキーフライドチキン", "飲食", true, "日本語表記対策（登録済みキーワード「KFC」が実店名に一致しないため追加）"],
    ["ケンタッキー", "ケンタッキー", "飲食", true, "日本語表記対策（短縮表記対応）"],
    ["ロイヤルホスト", "ロイヤルホスト", "飲食", true, "日本語表記対策（登録済みキーワード「Royal Host」が実店名に一致しないため追加）"],
    ["ジョイフル", "ジョイフル", "飲食", true, "日本語表記対策（登録済みキーワード「Joyfull」が実店名に一致しないため追加）"],
    ["バーミヤン", "バーミヤン", "飲食", true, "日本語表記対策（登録済みキーワード「Bamiyan」が実店名に一致しないため追加）"],
    ["フレッシュネスバーガー", "フレッシュネスバーガー", "飲食", true, "日本語表記対策（登録済みキーワード「FRESHNESS BURGER」が実店名に一致しないため追加）"],
    ["サーティワンアイスクリーム", "サーティワンアイスクリーム", "飲食", true, "日本語表記対策（登録済みキーワード「31アイス」が実店名の正式表記に一致しないため追加）"],
    ["いきなりステーキ", "いきなりステーキ", "飲食", true, "日本語表記対策（登録済みキーワード「IKINARI STEAK」が実店名に一致しないため追加）"],
    ["ペッパーランチ", "ペッパーランチ", "飲食", true, "日本語表記対策（登録済みキーワード「Pepper Lunch」が実店名に一致しないため追加）"],
    ["アンデルセン", "アンデルセン", "飲食", true, "日本語表記対策（登録済みキーワード「ANDERSEN」が実店名に一致しないため追加）"],
    ["ブロンコビリー", "ブロンコビリー", "飲食", true, "日本語表記対策（登録済みキーワード「BRONCO BILLY」が実店名に一致しないため追加）"],
    ["ウェンディーズ", "ウェンディーズ", "飲食", true, "日本語表記対策（登録済みキーワード「Wendy」が実店名に一致しないため追加）"],
    ["ファーストキッチン", "ファーストキッチン", "飲食", true, "日本語表記対策（登録済みキーワード「First Kitchen」が実店名に一致しないため追加）"],
    ["かつや", "かつや", "飲食", true, "とんかつチェーン。木更津市データで未除外を確認したため追加"],
    ["ヤマザキデイリーストア", "ヤマザキデイリーストア", "飲食", true, "コンビニチェーン。木更津市データで未除外を確認したため追加"],
    ["元祖からあげ本舗 だるま", "からあげ本舗 だるま", "飲食", true, "唐揚げチェーン。「だるま」単独だと一般的すぎて誤爆するため店名の特徴的な部分をキーワードに設定。木更津市データで2店舗(未除外)を確認したため追加"],
    ["ビッグボーイ", "ビッグボーイ", "飲食", true, "ステーキ・ハンバーグファミレスチェーン。白井市データでBarタブへの誤混入を確認したため追加"],
    ["ステーキハンバーグ＆サラダバーけん", "サラダバーけん", "飲食", true, "ステーキ・ハンバーグチェーン。「けん」単独だと一般的すぎて誤爆するため店名の特徴的な部分をキーワードに設定。白井市データでBarタブへの誤混入を確認したため追加"],

    // ↓ここから: 「餃子の王将」報告をきっかけに、他の主要全国チェーンで
    // MASTER_CHAINに未登録、または登録キーワードが実店名と一致しない
    // ものがないか横断チェックして見つかった追加分。
    ["餃子の王将", "餃子の王将", "飲食", true, "中華チェーン。登録済みキーワード「王将」は誤判定注意でFALSE、「京都王将」も実店名「餃子の王将 西白井店」に一致しないため正式名称で追加"],
    ["丸亀製麺", "丸亀製麺", "飲食", true, "うどんチェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["鳥貴族", "鳥貴族", "飲食", true, "焼き鳥居酒屋チェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["はなまるうどん", "はなまるうどん", "飲食", true, "うどんチェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["リンガーハット", "リンガーハット", "飲食", true, "ちゃんぽんチェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["山内農場", "山内農場", "飲食", true, "居酒屋チェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["デイリーヤマザキ", "デイリーヤマザキ", "飲食", true, "コンビニチェーン。登録済みキーワード「ヤマザキデイリーストア」は語順が異なり実店名「デイリーヤマザキ 木更津駅前店」に一致しないため追加"],
    ["函館函太郎", "函館函太郎", "飲食", true, "回転寿司チェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["金子半之助", "金子半之助", "飲食", true, "天丼チェーン。MASTER_CHAINに未登録だったため新規追加"],
    ["まごころ弁当", "まごころ弁当", "飲食", true, "宅配弁当チェーン。我孫子市データで「まごころ弁当 我孫子店」「まごころ弁当 千葉NT店」の2店舗が未除外だったため追加"],
    ["鮒忠", "鮒忠", "飲食", true, "やきとり・うなぎチェーン（関東広域展開）。香取市データで「鮒忠佐原店」がハンバーガータブに混入していたため追加"],
    ["カラオケまねきねこ", "まねきねこ", "共通", true, "全国チェーンのカラオケ店（飲食店ではない）。香取市データで「まねきねこ」が弁当タブに混入していたため追加。店名に「カラオケ」が付かない表記のためFACILITY_EXCLUDE_KEYWORDSの「カラオケ」だけでは弾けず、チェーンマスタ側で対応"]
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("MASTER_CHAIN");
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("MASTER_CHAINシートが見つかりません。");
    return;
  }

  const existingValues = masterSheet.getDataRange().getValues();
  const existingKeywords = new Set(
    existingValues.slice(1).map(r => String(r[1]).normalize("NFC").trim())
  );

  const newRows = KNOWN_GAPS.filter(row => !existingKeywords.has(String(row[1]).normalize("NFC").trim()));

  if (newRows.length === 0) {
    SpreadsheetApp.getUi().alert("追加対象がありません（すべて登録済みです）。");
    return;
  }

  const startRow = masterSheet.getLastRow() + 1;
  masterSheet.getRange(startRow, 1, newRows.length, 5).setNumberFormat("@").setValues(newRows);
  masterSheet.getRange(startRow, 4, newRows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );

  SpreadsheetApp.getUi().alert(
    `【完了】${newRows.length}件のキーワードをMASTER_CHAINへ追加しました。\n` +
    newRows.map(r => `・${r[0]}（${r[1]}）`).join("\n") +
    `\n\n反映するには「3. チェーン判定」以降を再実行してください。`
  );
}

function detectAreaNameFromComdeskRows(values) {
  const header = values[0];
  const prefIdx = header.indexOf("都道府県");
  const addrIdx = header.indexOf("住所１");

  if (prefIdx === -1 || addrIdx === -1) return "";

  for (let i = 1; i < values.length; i++) {
    const pref = textValue(values[i][prefIdx]);
    const addr = textValue(values[i][addrIdx]);

    if (!pref || !addr) continue;

    const cityMatch = addr.match(/^(.+?[市区町村])/);
    const city = cityMatch ? cityMatch[1] : "";

    if (city) return pref + city;
    return pref;
  }

  return "";
}

function getComdeskHeader() {
  return [
    "UUID", "種別", "名前", "カナ", "郵便番号", "都道府県", "住所１", "住所２", "住所カナ",
    "Tel1", "Tel2", "Tel3", "Tel4", "FAX", "URL", "備考", "旧社名", "リードソース",
    "旧進捗", "履歴", "オーナー名", "HPある？", "BP検索", "アポ済商材", "最新履歴", "営業曜日", "休業曜日",
    "午前始", "午前終", "午後始", "午後終"
  ];
}

function buildComdeskRow(header, row) {
  const storeName = getRowValueByHeader(header, row, "店名");
  const fullAddr = getRowValueByHeader(header, row, "住所");
  const addrDetails = parseAddressDetails(fullAddr);
  const pref = getRowValueByHeader(header, row, "都道府県") || addrDetails.pref;
  const city = getRowValueByHeader(header, row, "市区町村");
  const phone = getRowValueByHeader(header, row, "電話番号");
  const formattedPhone = normalizePhoneNumberForAnalysis(phone);
  const cleanPhone = formattedPhone.replace(/[^\d]/g, "");
  const media = getRowValueByHeader(header, row, "媒体");
  const url = getRowValueByHeader(header, row, "URL");
  const hpHave = getRowValueByHeader(header, row, "HP有無");
  const hpStatus = (hpHave.indexOf("有") !== -1 || hpHave === "1" || hpHave.toLowerCase() === "true") ? "1" : "0";
  const bizDaysVal = removeHolidayFromBizDays(getRowValueByHeader(header, row, "営業日"), getRowValueByHeader(header, row, "定休日"));
  const holidayVal = getRowValueByHeader(header, row, "定休日");
  const rawOpenA = getRowValueByHeader(header, row, "営業開始A") || getRowValueByHeader(header, row, "営業開始");
  const openAVal = formatToPureTime(toHalfWidthForTime(rawOpenA));
  const closeAVal = formatToPureTime(toHalfWidthForTime(getRowValueByHeader(header, row, "営業終了A") || getRowValueByHeader(header, row, "営業終了")));
  const openBVal = formatToPureTime(toHalfWidthForTime(getRowValueByHeader(header, row, "営業開始B")));
  const closeBVal = formatToPureTime(toHalfWidthForTime(getRowValueByHeader(header, row, "営業終了B")));
  const timeValues = normalizeBusinessTimeValues(rawOpenA, openAVal, closeAVal, openBVal, closeBVal);
  // addrDetails.addr1 は郵便番号・都道府県を除去済みの住所なので、これを土台にする
  // （fullAddrをそのまま使うと「◯◯市〒123-4567△△町」のように郵便番号が住所の途中に残ってしまう）
  const cleanAddr1 = city ? addrDetails.addr1.replace(city, "") : addrDetails.addr1;
  const address1 = city ? city + cleanAddr1 : addrDetails.addr1;
  const areaText = pref + city;

  const salesRow = Array(31).fill("");
  salesRow[2] = storeName;
  salesRow[4] = addrDetails.pcode;
  salesRow[5] = pref;
  salesRow[6] = address1;
  salesRow[9] = formattedPhone || phone;
  salesRow[14] = url;
  salesRow[15] = "";
  salesRow[17] = media;
  salesRow[21] = hpStatus;
  salesRow[22] = `${areaText}tel${cleanPhone}`;
  salesRow[25] = bizDaysVal;
  salesRow[26] = holidayVal;
  salesRow[27] = timeValues.openA;
  salesRow[28] = timeValues.closeA;
  salesRow[29] = timeValues.openB;
  salesRow[30] = timeValues.closeB;
  return salesRow;
}

function toHalfWidthForTime(str) {
  return textValue(str).replace(/[０-９：]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
}

function normalizeBusinessTimeValues(rawOpenA, openAVal, closeAVal, openBVal, closeBVal) {
  const result = { openA: openAVal, closeA: closeAVal, openB: openBVal, closeB: closeBVal };
  if (!openBVal && !closeBVal) {
    const timeMatches = toHalfWidthForTime(rawOpenA).match(/(\d{1,2}:\d{2})/g);
    if (timeMatches && timeMatches.length >= 2) {
      result.openA = formatToPureTime(timeMatches[0]);
      result.closeA = "";
      result.openB = "";
      result.closeB = formatToPureTime(timeMatches[timeMatches.length - 1]);
    } else if (openAVal) {
      result.closeA = "";
      result.openB = "";
      result.closeB = closeAVal;
    }
  }
  return result;
}

function removeHolidayFromBizDays(bizDays, holiday) {
  let bizDaysVal = textValue(bizDays);
  const holidayVal = textValue(holiday);
  if (!bizDaysVal || !holidayVal) return bizDaysVal;

  ["月", "火", "水", "木", "金", "土", "日", "祝"].forEach(day => {
    if (holidayVal.indexOf(day) !== -1) {
      const regex = new RegExp(day + "[・、/]?|[・、/]?" + day, "g");
      bizDaysVal = bizDaysVal.replace(regex, "");
    }
  });
  return bizDaysVal.replace(/^[・、/]+|[・、/]+$/g, "").replace(/[・、/]{2,}/g, "・");
}

// =====================================================================
// ジャンル正規化（Ver9.8.1：検索ジャンルを一般フォールバックとして活用）
// =====================================================================
function normalizeSystemGenre(genre, searchGenre, sourceGenre, storeName) {
  // ★最優先ルール: 店名に「居酒屋」と明記されている店は、Google/Tabelog側の
  // カテゴリ判定がBar・和食・その他何であっても、必ず「居酒屋」に確定させる。
  // 他の広い受け皿ジャンルに絶対に紛れさせないための最優先チェック
  // （HD_GENRE_MAPによるマッピングより先に判定する）。
  if (textValue(storeName).indexOf("居酒屋") !== -1) return "居酒屋";

  // ★最優先ルール: 店名にカフェ関連キーワード（カフェ/coffee/珈琲/喫茶等）が
  // 含まれる場合も、Google側の生ジャンル抽出結果に関わらず必ず「カフェ」に確定
  // させる。実データで「検索:カフェ」でヒットした店が、抽出不良により
  // 取得元ジャンルが「洋食」等の無関係な値になり、店名が明らかにカフェ
  // （例:「MIFUNEYAMA COFFEE」「cafe Bohemian」等）でも洋食タブに混入する
  // 不具合を88件規模で確認済み。居酒屋と同様、店名の方が生ジャンルより
  // 信頼できるため最優先で判定する。
  if (CAFE_KEYWORDS.some(keyword => textValue(storeName).indexOf(keyword) !== -1)) return "カフェ";

  // ★最優先ルール: 店名にラーメン・蕎麦うどん・寿司・焼肉等の具体的なジャンル語が
  // 含まれる場合も、生ジャンルが既に「有効な別ジャンル」に見えていても店名を優先する。
  // 実データで「横浜ラーメン寺田家」「竹岡式ラーメンたか木屋白井店」等、店名に
  // 「ラーメン」と明記された店が、抽出不良で生ジャンルが「カフェ」等の別の有効な
  // ジャンルになってしまい、!isValidHdGenre(mappedGenre)による従来のフォールバック
  // （生ジャンルが無効な場合のみ店名を見る）では拾えないケースを確認したため、
  // 居酒屋・カフェと同様に無条件の最優先チェックへ格上げする。
  const nameGenrePriority = findGenreFromStoreName(storeName);
  if (nameGenrePriority) return nameGenrePriority;

  const rawGenre = textValue(genre);
  let mappedGenre = HD_GENRE_MAP[rawGenre] || rawGenre;
  if (textValue(searchGenre) === "喫茶店") return "カフェ";

  // ★修正: Googleマップ側のカテゴリ体系には「居酒屋」の粒度が無いことが多く、
  // アルコールを提供する店は広く「Bar」に分類されがちだった
  // （例:「居酒屋 華」がBarタブに混入する不具合）。店名に「居酒屋」と
  // 明記されているケースは上の最優先ルールで既に確定しているので、
  // ここではBar判定時のみ、それ以外の具体的なキーワードを優先する。
  if (mappedGenre === "Bar") {
    const nameText = textValue(storeName);
    if (nameText.indexOf("スナック") !== -1) return "スナック";
    if (nameText.indexOf("焼き鳥") !== -1 || nameText.indexOf("焼鳥") !== -1) return "焼き鳥";
  }

  if (!isValidHdGenre(mappedGenre)) {
    const matchedOldGenre = Object.keys(HD_GENRE_MAP).find(oldGenre => rawGenre.indexOf(oldGenre) !== -1 || textValue(sourceGenre).indexOf(oldGenre) !== -1);
    if (matchedOldGenre) mappedGenre = HD_GENRE_MAP[matchedOldGenre];
  }
  // （店名からのジャンル優先判定は上の最優先ルールに統合済みのため、ここでは行わない）

  // ★方針変更(Ver9.9.2): 以前はここで「店舗自身のジャンル表記からも店名からも
  // 有効な統一ジャンルを決められなかった場合、検索に使ったジャンル語（検索ジャンル列）を
  // 最後のフォールバックとして採用する」処理を行っていたが、これが原因で
  // 郊外エリア（例: 香取市）のように該当ジャンルの店が少ない検索では、Googleマップ側が
  // 検索条件に緩く一致するだけの無関係な店（そば屋・食堂・チェーン店・ゴルフ場の
  // クラブハウス・カラオケ店等）まで結果に含めてしまい、それらが「検索した時の
  // ジャンル」にそのまま割り当てられて営業リストに混入する不具合が繰り返し発生した
  // （香取市データで「香取カントリークラブ」「鮒忠佐原店」「まねきねこ」等が
  // ハンバーガー・弁当タブに混入していたケースで確認）。
  // 店名にも店舗自身のジャンル表記にも手がかりが無い店を「検索ジャンルだから
  // 多分そうだろう」と決めつけるのはリスクが高いため、このフォールバックは廃止する。
  // 結果として、判定できなかった店は isValidHdGenre(mappedGenre) が false のまま
  // 返り、基本データ判定で「確認対象」（ジャンル確認）に回るため、誤って
  // 営業対象へ混入することはなくなる。

  return mappedGenre;
}

function judgeFacilityStatus(storeName, address, genre) {
  const haystack = [storeName, address, genre].map(textValue).join(" ");
  const excludeKeyword = FACILITY_EXCLUDE_KEYWORDS.find(keyword => haystack.indexOf(keyword) !== -1);
  if (excludeKeyword) return { status: "除外", reason: "完全除外キーワード一致: " + excludeKeyword };
  const reviewKeyword = FACILITY_REVIEW_KEYWORDS.find(keyword => haystack.indexOf(keyword) !== -1);
  if (reviewKeyword) return { status: "確認対象", reason: "確認対象キーワード一致: " + reviewKeyword };
  return { status: "対象", reason: "" };
}

function judgeSalesTargetStatus(header, row, facility) {
  const reasons = [];
  let hasReview = false;
  let hasExclude = false;
  const fetchStatus = getRowValueByHeader(header, row, "取得ステータス");
  const externalReason = getRowValueByHeader(header, row, "除外理由");
  const basicStatus = getRowValueByHeader(header, row, "基本データ判定");
  const basicReason = getRowValueByHeader(header, row, "基本データ除外理由");
  const areaStatus = getRowValueByHeader(header, row, "エリア判定");
  const areaReason = getRowValueByHeader(header, row, "エリア判定理由");
  const dupStatus = getRowValueByHeader(header, row, "重複判定");
  const chainStatus = getRowValueByHeader(header, row, "チェーン判定");

  if (fetchStatus === "失敗") { hasExclude = true; reasons.push("取得失敗"); }
  if (externalReason) { hasExclude = true; reasons.push(externalReason); }
  if (basicStatus !== "対象") { hasReview = true; reasons.push(basicReason || "基本データ確認"); }
  if (areaStatus === "判定不可") { hasReview = true; reasons.push(areaReason || "エリア判定不可"); }
  if (areaStatus === "エリア外") { hasExclude = true; reasons.push(areaReason || "エリア外"); }
  if (dupStatus !== "ユニーク") { hasExclude = true; reasons.push(dupStatus || "重複"); }
  if (chainStatus === "チェーン店") { hasExclude = true; reasons.push("チェーン店"); }
  if (facility.status === "確認対象") { hasReview = true; reasons.push(facility.reason); }
  if (facility.status === "除外") { hasExclude = true; reasons.push(facility.reason); }

  const joined = uniqueTextList(reasons).join(" / ");
  if (!joined) return { status: "対象", reason: "" };
  if (hasReview && !hasExclude) return { status: "確認対象", reason: joined };
  return { status: "除外", reason: joined };
}

function judgeWorkflowGroup(header, row) {
  const storeName = getRowValueByHeader(header, row, "店名");
  const address = getRowValueByHeader(header, row, "住所");
  const url = getRowValueByHeader(header, row, "URL");
  const fetchStatus = getRowValueByHeader(header, row, "取得ステータス");
  const externalReason = getRowValueByHeader(header, row, "除外理由");
  const dupStatus = getRowValueByHeader(header, row, "重複判定");
  const chainStatus = getRowValueByHeader(header, row, "チェーン判定");
  const facilityStatus = getRowValueByHeader(header, row, "施設判定");
  const areaStatus = getRowValueByHeader(header, row, "エリア判定");
  const addressStatus = getRowValueByHeader(header, row, "住所判定");
  const normalizedPhone = getRowValueByHeader(header, row, "正規化電話番号") || normalizePhoneNumberForAnalysis(getRowValueByHeader(header, row, "電話番号"));
  const normalizedGenre = getRowValueByHeader(header, row, "正規化ジャンル") || normalizeSystemGenre(getRowValueByHeader(header, row, "ジャンル"), getRowValueByHeader(header, row, "検索ジャンル"), getRowValueByHeader(header, row, "取得元ジャンル"), storeName);
  const failureText = [fetchStatus, externalReason, getRowValueByHeader(header, row, "営業対象除外理由")].join(" / ");

  if (fetchStatus === "失敗" || failureText.indexOf("詳細取得失敗") !== -1 || (!storeName && url)) {
    return { group: "04_取得失敗", item: "詳細取得失敗", status: "未対応", nextAction: "再取得" };
  }

  if (dupStatus !== "ユニーク") return { group: "03_除外対象", item: "重複除外", status: "除外確定", nextAction: "投入しない" };
  if (chainStatus === "チェーン店") return { group: "03_除外対象", item: "チェーン店除外", status: "除外確定", nextAction: "投入しない" };
  if (facilityStatus === "除外") return { group: "03_除外対象", item: "ビル管理除外", status: "除外確定", nextAction: "投入しない" };
  if (areaStatus === "エリア外") return { group: "03_除外対象", item: "エリア外除外", status: "除外確定", nextAction: "投入しない" };
  if (!storeName || (!address && !url) || (addressStatus === "住所未取得" && !url)) {
    return { group: "03_除外対象", item: "住所未取得除外", status: "除外確定", nextAction: "投入しない" };
  }

  if (facilityStatus === "確認対象") return { group: "02_確認対象", item: "小規模ビル確認", status: "未対応", nextAction: "電力契約変更可否を確認" };
  if (!normalizedPhone && storeName && address) return { group: "02_確認対象", item: "電話番号なし確認", status: "未対応", nextAction: "電話番号補完" };
  if (!normalizedGenre || normalizedGenre === "その他" || !isValidHdGenre(normalizedGenre)) return { group: "02_確認対象", item: "ジャンル確認", status: "未対応", nextAction: "ジャンルを目視確認" };
  if (addressStatus !== "住所あり") return { group: "02_確認対象", item: "住所確認", status: "未対応", nextAction: "住所確認" };
  if (areaStatus === "判定不可") return { group: "02_確認対象", item: "住所確認", status: "未対応", nextAction: "住所確認" };

  if (isComdeskTargetRow(header, row)) return { group: "01_営業対象", item: "営業対象", status: "未対応", nextAction: "コムデスク投入" };
  return { group: "02_確認対象", item: "住所確認", status: "未対応", nextAction: "目視確認" };
}

function judgeAreaStatus(address, pref, city) {
  const cleanAddress = normalizeAddressText(address);
  const prefText = textValue(pref);
  const cityText = textValue(city);
  if (!cleanAddress) return { status: "判定不可", reason: "住所未取得" };
  if (!prefText || !cityText) return { status: "判定不可", reason: "都道府県または市区町村が空" };
  if (isSimpleAddress(cleanAddress, cityText)) return { status: "判定不可", reason: "簡易住所" };
  if (cleanAddress.indexOf(prefText) === -1) return { status: "エリア外", reason: "都道府県不一致: " + prefText };
  if (cleanAddress.indexOf(cityText) === -1) return { status: "エリア外", reason: "市区町村不一致: " + cityText };
  return { status: "エリア内", reason: "" };
}

function judgeAddressStatus(address, pref, city) {
  const cleanAddress = normalizeAddressText(address);
  if (!cleanAddress) return { status: "住所未取得", reason: "住所未取得" };
  if (!textValue(pref) || !textValue(city)) return { status: "住所未取得", reason: "都道府県または市区町村が空" };
  if (isSimpleAddress(cleanAddress, textValue(city))) return { status: "住所未取得", reason: "簡易住所" };
  return { status: "住所あり", reason: "" };
}

function isComdeskTargetRow(header, row) {
  const phone = getRowValueByHeader(header, row, "正規化電話番号") || normalizePhoneNumberForAnalysis(getRowValueByHeader(header, row, "電話番号"));
  const genre = getRowValueByHeader(header, row, "正規化ジャンル") || getRowValueByHeader(header, row, "ジャンル");
  return getRowValueByHeader(header, row, "営業対象判定") === "対象" &&
    getRowValueByHeader(header, row, "重複判定") === "ユニーク" &&
    getRowValueByHeader(header, row, "チェーン判定") === "単独店" &&
    getRowValueByHeader(header, row, "施設判定") === "対象" &&
    getRowValueByHeader(header, row, "エリア判定") === "エリア内" &&
    !!phone &&
    !!getRowValueByHeader(header, row, "住所") &&
    isValidHdGenre(genre);
}

function normalizeHeaderRow(header) {
  return header.map(h => String(h).replace(/^\uFEFF/, "").trim());
}

function ensureHeaderColumns(header, rows, columns) {
  columns.forEach(name => {
    if (header.indexOf(name) === -1) {
      header.push(name);
      rows.forEach(row => row.push(""));
    }
  });
}

function buildHeaderIndex(header) {
  const indexes = {};
  header.forEach((name, index) => { indexes[name] = index; });
  return indexes;
}

function getRowValueByHeader(header, row, name) {
  const idx = header.indexOf(name);
  return idx === -1 ? "" : textValue(row[idx]);
}

function setRowValueByHeader(header, row, name, value) {
  const idx = header.indexOf(name);
  if (idx !== -1) row[idx] = value;
}

function writeRowsToExistingSheet(sheet, header, rows) {
  sheet.clearContents();
  const output = [header].concat(rows);
  sheet.getRange(1, 1, output.length, header.length).setNumberFormat("@").setValues(output);
}

function writeRowsToSheetByName(ss, sheetName, header, rows) {
  const sheet = getOrCreateSheet(ss, sheetName);
  writeRowsToExistingSheet(sheet, header, rows);
}

function fillPrefCityFromAddress(header, row) {
  const prefIdx = header.indexOf("都道府県");
  const cityIdx = header.indexOf("市区町村");
  if (prefIdx === -1 || cityIdx === -1) return;
  if (textValue(row[prefIdx]) && textValue(row[cityIdx])) return;
  const parsed = parsePrefCityFromAddress(getRowValueByHeader(header, row, "住所"));
  if (!textValue(row[prefIdx])) row[prefIdx] = parsed.pref;
  if (!textValue(row[cityIdx])) row[cityIdx] = parsed.city;
}

function parsePrefCityFromAddress(address) {
  const cleanAddress = normalizeAddressText(address);
  const match = cleanAddress.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))?((?:.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村]))?/);
  return {
    pref: match && match[1] ? match[1] : "",
    city: match && match[2] ? match[2] : ""
  };
}

function normalizePhoneNumberForAnalysis(phone) {
  const digits = textValue(phone)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[^\d]/g, "");
  if (!digits) return "";
  if (digits.length === 11) return digits.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3");
  if (digits.startsWith("03") || digits.startsWith("06")) return digits.replace(/(\d{2})(\d{4})(\d{4})/, "$1-$2-$3");
  if (digits.length === 10) return digits.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
  return digits;
}

function normalizeAddressText(address) {
  return textValue(address).replace(/(?:〒\d{3}-\d{4}\s*|日本、\s*)/g, "").replace(/\s+/g, "");
}

function isSimpleAddress(address, city) {
  return !/[都道府県]/.test(address) || (city && address.indexOf(city) === -1);
}

function isValidHdGenre(genre) {
  return HD_TARGET_GENRES.indexOf(textValue(genre)) !== -1;
}

function uniqueTextList(values) {
  const seen = {};
  return values.map(textValue).filter(value => {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function textValue(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}