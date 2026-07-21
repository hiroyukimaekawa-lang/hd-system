/**
 * 店舗管理システム - 統合処理モジュール Ver9.7.0 (飲食店・最強版 ＋ 超・重複排除機能)
 * (完成版CSVファイル全自動エクスポート ➔ ログ自動永久保存 ➔ 31項目新フォーマット ➔ BP検索確定値化 ➔ 時間の全角対応＆通し営業時間自動スライド ➔ 曜日除外処理 ➔ 表記揺れ完全吸収)
 */

// =====================================================================
// ⚙️ 設定エリア
// SerpAPIキーはコードに直接書かず、GASのスクリプトプロパティ
// （ファイル > プロジェクトの設定 > スクリプト プロパティ）に
//   SERPAPI_API_KEY = <あなたのAPIキー>
// として登録してください。未設定の場合、SerpAPI補完はスキップされます。
// =====================================================================
function getSerpApiKeyOrEmpty_() {
  try {
    return PropertiesService.getScriptProperties().getProperty("SERPAPI_API_KEY") || "";
  } catch (e) {
    return "";
  }
}

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
  executeProcessLogSummary();
  SpreadsheetApp.getUi().alert("HDリスト処理が完了しました。04_SALES_ジャンル別タブとCSVを確認してください。");
}

// =====================================================================
// 処理0: 文字コード自動判別 ＆ 不可視ゴミ(BOM)自動抹殺インポート
// =====================================================================
function importCSVFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = getOrCreateSheet(ss, "01_NORMALIZED");
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  let importFolder = getOrCreateFolder(parentFolder, "CSV投入フォルダ");
  let processedFolder = getOrCreateFolder(parentFolder, "処理済みフォルダ");
  const files = importFolder.getFilesByType(MimeType.CSV);

  let isFirstFile = true;
  let combinedData = [];
  let headerRow = [];

  while (files.hasNext()) {
    const file = files.next();
    try {
      ss.toast(`ファイル「${file.getName()}」を解析・合体中...`, "📁 CSV一括取り込み");

      let blob = file.getBlob();
      let csvText = blob.getDataAsString("UTF-8").replace(/^\uFEFF/, "");

      if (!csvText.includes("店名")) {
        csvText = blob.getDataAsString("MS932").replace(/^\uFEFF/, "");
      }

      const parsedCsv = Utilities.parseCsv(csvText);

      if (parsedCsv.length > 0) {
        if (isFirstFile) {
          headerRow = parsedCsv[0].map(h => String(h).replace(/^\uFEFF/, "").trim());
          combinedData.push(headerRow);
          isFirstFile = false;
        }
        const dataRows = parsedCsv.slice(1);
        dataRows.forEach(row => {
          if (row.join("").trim() !== "") {
            combinedData.push(row);
          }
        });
      }
      file.moveTo(processedFolder);
    } catch (e) {
      Logger.log(`[CSV取込エラー] ${file.getName()}: ${e.message}`);
    }
  }

  if (combinedData.length > 1) {
    targetSheet.clear();
    const colCount = combinedData[0].length;
    const normalizedData = combinedData.map(row => {
      if (row.length === colCount) return row;
      const padded = row.slice(0, colCount);
      while (padded.length < colCount) padded.push("");
      return padded;
    });
    const range = targetSheet.getRange(1, 1, normalizedData.length, colCount);
    range.setNumberFormat("@");
    range.setValues(normalizedData);
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
    if (isValid && keyword !== "") { chainMaster.push({ chainName: chainName, keyword: keyword, industry: industry }); }
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
  const candidateLog = {};
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
      if (storeName.includes(master.keyword)) {
        isChain = true; matchedChainName = master.chainName;
        chainReason = "マスタ合致: キーワード[" + master.keyword + "]"; break;
      }
    }
    if (!isChain && cleanName && brandCounter[cleanName] >= 5) {
      isChain = true; matchedChainName = cleanName + "チェーン";
      chainReason = "自動検出: 出現数[" + brandCounter[cleanName] + "]件";
      candidateLog[cleanName] = brandCounter[cleanName];
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
  if (Object.keys(candidateLog).length > 0) { writeToChainCandidateHook(ss, candidateLog); }
}

// 店舗の業種を判定する（MASTER_CHAIN の業種フィルタに使用）
function getStoreIndustry(header, row) {
  // 明示的な業種列があればそれを優先
  const explicit = getRowValueByHeader(header, row, "業種");
  if (explicit === "美容" || explicit === "飲食") return explicit;

  // 正規化ジャンルが美容院なら美容
  const genre = getRowValueByHeader(header, row, "正規化ジャンル") || getRowValueByHeader(header, row, "ジャンル");
  if (genre === "美容院") return "美容";

  // 媒体名にビューティー系ワードが含まれていれば美容
  const media = getRowValueByHeader(header, row, "媒体");
  if (media && (media.includes("ビューティー") || media.includes("BEAUTY") || media.includes("Beauty"))) return "美容";

  return "飲食";
}

// =====================================================================
// 処理3: 営業リスト生成 ＆ 31項目新フォーマット・CSV自動エクスポート
// =====================================================================
function executeGenerateSalesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("03_CHAIN_CHECK");
  if (!sourceSheet) return;

  const range = sourceSheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return;

  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();
  let exportFolder = getOrCreateFolder(parentFolder, "完成版CSVエクスポート");

  const srcHeader = values[0].map(h => String(h).replace(/^\uFEFF/, "").trim());

  const nameIdx    = srcHeader.indexOf("店名");
  const genreIdx   = srcHeader.indexOf("ジャンル");
  const prefIdx    = srcHeader.indexOf("都道府県");
  const cityIdx    = srcHeader.indexOf("市区町村");
  const addrIdx    = srcHeader.indexOf("住所");
  const phoneIdx   = srcHeader.indexOf("電話番号");
  const holidayIdx = srcHeader.indexOf("定休日");
  const bizDaysIdx = srcHeader.indexOf("営業日");

  const openAIdx   = srcHeader.indexOf("営業開始A") !== -1 ? srcHeader.indexOf("営業開始A") : srcHeader.indexOf("営業開始");
  const closeAIdx  = srcHeader.indexOf("営業終了A") !== -1 ? srcHeader.indexOf("営業終了A") : srcHeader.indexOf("営業終了");
  const openBIdx   = srcHeader.indexOf("営業開始B");
  const closeBIdx  = srcHeader.indexOf("営業終了B");

  const urlIdx     = srcHeader.indexOf("URL");
  const hpHaveIdx  = srcHeader.indexOf("HP有無");
  const mediaIdx   = srcHeader.indexOf("媒体");
  const dupIdx     = srcHeader.indexOf("重複判定");
  const chainIdx   = srcHeader.indexOf("チェーン判定");

  if (dupIdx === -1 || chainIdx === -1 || phoneIdx === -1 || nameIdx === -1) {
    SpreadsheetApp.getUi().alert("【エラー】最終ステップで見出しの自動特定に失敗しました。システム管理者に連絡してください。");
    return;
  }

  let currentTargetCount = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][dupIdx]).trim() === "ユニーク" && String(values[i][chainIdx]).trim() === "単独店" && String(values[i][phoneIdx]).replace(/[^\d]/g, "").trim() === "") currentTargetCount++;
  }

  let apiCallCount = 0;
  let currentProgress = 0;
  const SERPAPI_KEY = getSerpApiKeyOrEmpty_();
  if (SERPAPI_KEY !== "YOUR_SERPAPI_API_KEY_HERE" && SERPAPI_KEY !== "") {
    for (let i = 1; i < values.length; i++) {
      const isUnique = (String(values[i][dupIdx]).trim() === "ユニーク");
      const isSingleStore = (String(values[i][chainIdx]).trim() === "単独店");
      const currentPhone = String(values[i][phoneIdx]).replace(/[^\d]/g, "").trim();

      if (isUnique && isSingleStore && currentPhone === "") {
        const storeName = String(values[i][nameIdx]).trim();
        currentProgress++;

        ss.toast(`「${storeName}」の電話番号をGoogle検索中... (${currentProgress} / ${currentTargetCount}件中)`, "🚀 SerpAPI自動補完中");

        const address = addrIdx !== -1 ? String(values[i][addrIdx]).trim() : "";
        const searchQuery = `${storeName} ${address} 電話番号`.trim();
        try {
          const url = `https://serpapi.com/search.json?q=${encodeURIComponent(searchQuery)}&hl=ja&gl=jp&api_key=${SERPAPI_KEY}`;
          const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
          const json = JSON.parse(response.getContentText());
          let foundPhone = "";
          if (json.knowledge_graph && json.knowledge_graph.phone) { foundPhone = json.knowledge_graph.phone; }
          if (!foundPhone && json.organic_results && json.organic_results.length > 0) {
            const telRegex = /(?:0\d{1,4}-\d{1,4}-\d{3,4}|\d{10,11})/;
            for (let j = 0; j < Math.min(json.organic_results.length, 3); j++) {
              const snippet = json.organic_results[j].snippet || "";
              const match = snippet.match(telRegex);
              if (match) { foundPhone = match[0]; break; }
            }
          }
          if (foundPhone) { values[i][phoneIdx] = foundPhone.replace(/[^\d\-]/g, "").trim(); apiCallCount++; }
          Utilities.sleep(200);
        } catch (e) { Logger.log(`[SerpAPIエラー] ${storeName}: ${e.message}`); }
      }
    }
    if (apiCallCount > 0) { range.setValues(values); }
  }

  const finalHeader = [
    'UUID', '種別', '名前', 'カナ', '郵便番号', '都道府県', '住所１', '住所２', '住所カナ',
    'Tel1', 'Tel2', 'Tel3', 'Tel4', 'FAX', 'URL', '備考', '旧社名', 'リードソース',
    '旧進捗', '履歴', 'オーナー名', 'HPある？', 'BP検索', 'アポ済商材', '最新履歴', '営業曜日', '休業曜日',
    '午前始', '午前終', '午後始', '午後終'
  ];

  const TARGET_GENRES = HD_TARGET_GENRES;

  const genreContainers = {};
  TARGET_GENRES.forEach(g => { genreContainers[g] = []; });
  genreContainers["その他"] = [];
  const dataRows = values.slice(1);
  const downloadTargetArea = detectDownloadTargetArea(srcHeader, dataRows);

  let detectedAreaName = downloadTargetArea.areaText || "ダウンロードリスト";

  dataRows.forEach(row => {
    const isUnique = (String(row[dupIdx]).trim() === "ユニーク");
    const isSingleStore = (String(row[chainIdx]).trim() === "単独店");

    if (isUnique && isSingleStore) {
      const genreVal = genreIdx !== -1 ? String(row[genreIdx]) : "";
      const storeNameVal = nameIdx !== -1 ? String(row[nameIdx]) : "";
      const fullAddr = addrIdx !== -1 ? String(row[addrIdx]) : "";
      const addrDetails = parseAddressDetails(fullAddr);
      const currentPref = prefIdx !== -1 ? String(row[prefIdx]).trim() : addrDetails.pref;
      const currentCity = cityIdx !== -1 ? String(row[cityIdx]).trim() : "";
      const areaTargetPref = downloadTargetArea.pref || currentPref;
      const areaTargetCity = downloadTargetArea.city || currentCity;
      const areaStatus = judgeAreaStatus(fullAddr, areaTargetPref, areaTargetCity);
      if (areaStatus.status !== "エリア内") return;

      const facilityStatus = judgeFacilityStatus(storeNameVal, fullAddr);
      if (facilityStatus.status === "除外") return;

      const normalizedGenreVal = normalizeSystemGenre(genreVal, getRowValueByHeader(srcHeader, row, "検索ジャンル"), getRowValueByHeader(srcHeader, row, "取得元ジャンル"), storeNameVal);
      const targetGroup = determineGenreGroup(normalizedGenreVal, storeNameVal, TARGET_GENRES);
      if (targetGroup === "SKIP") return;

      const toHalfWidth = (str) => str.replace(/[０-９：]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));

      let rawOpenA = openAIdx !== -1 ? String(row[openAIdx]) : "";
      let openAVal  = formatToPureTime(toHalfWidth(rawOpenA));
      let closeAVal = closeAIdx !== -1 ? formatToPureTime(toHalfWidth(String(row[closeAIdx]))) : "";
      let openBVal  = openBIdx !== -1 ? formatToPureTime(toHalfWidth(String(row[openBIdx]))) : "";
      let closeBVal = closeBIdx !== -1 ? formatToPureTime(toHalfWidth(String(row[closeBIdx]))) : "";

      let finalOpenA = openAVal;
      let finalCloseA = closeAVal;
      let finalOpenB = openBVal;
      let finalCloseB = closeBVal;

      if (!openBVal && !closeBVal) {
          const timeMatches = toHalfWidth(rawOpenA).match(/(\d{1,2}:\d{2})/g);
          if (timeMatches && timeMatches.length >= 2) {
              finalOpenA = formatToPureTime(timeMatches[0]);
              finalCloseA = "";
              finalOpenB = "";
              finalCloseB = formatToPureTime(timeMatches[timeMatches.length - 1]);
          } else if (openAVal) {
              finalOpenA = openAVal;
              finalCloseA = "";
              finalOpenB = "";
              finalCloseB = closeAVal;
          }
      }

      let hpHaveVal = hpHaveIdx !== -1 ? String(row[hpHaveIdx]).trim() : "";
      let hpStatus = "0";
      if (hpHaveVal.includes("有") || hpHaveVal === "1" || hpHaveVal.toLowerCase() === "true") {
        hpStatus = "1";
      }

      let bizDaysVal = bizDaysIdx !== -1 ? String(row[bizDaysIdx]).trim() : "";
      let holidayVal = holidayIdx !== -1 ? String(row[holidayIdx]).trim() : "";

      if (bizDaysVal && holidayVal) {
          const weekDays = ["月", "火", "水", "木", "金", "土", "日", "祝"];
          weekDays.forEach(day => {
              if (holidayVal.includes(day)) {
                  const regex = new RegExp(day + "[・、/]?|[・、/]?" + day, "g");
                  bizDaysVal = bizDaysVal.replace(regex, "");
              }
          });
          bizDaysVal = bizDaysVal.replace(/^[・、/]+|[・、/]+$/g, "").replace(/[・、/]{2,}/g, "・");
      }

      const mediaVal = mediaIdx !== -1 ? String(row[mediaIdx]) : "";
      let outputPref = areaTargetPref || currentPref || addrDetails.pref;
      let outputCity = areaTargetCity || currentCity;
      let areaText = outputPref + outputCity;
      if (areaText && detectedAreaName === "ダウンロードリスト") { detectedAreaName = areaText; }

      let cleanPhone = phoneIdx !== -1 ? String(row[phoneIdx]).replace(/[^\d]/g, "").trim() : "";

      const salesRow = Array(31).fill("");
      salesRow[0]  = "";
      salesRow[1]  = "";
      salesRow[2]  = storeNameVal;
      salesRow[3]  = "";
      salesRow[4]  = addrDetails.pcode;
      salesRow[5]  = outputPref;
      salesRow[6]  = cityIdx !== -1 ? outputCity + String(addrIdx !== -1 ? row[addrIdx] : "").replace(outputPref, "").replace(outputCity, "") : addrDetails.addr1;
      salesRow[7]  = ""; salesRow[8]  = "";
      salesRow[9]  = phoneIdx !== -1 ? String(row[phoneIdx]) : "";
      salesRow[10] = ""; salesRow[11] = ""; salesRow[12] = ""; salesRow[13] = "";
      salesRow[14] = urlIdx !== -1 ? String(row[urlIdx]) : "";
      salesRow[15] = "";
      salesRow[16] = ""; salesRow[17] = mediaVal;
      salesRow[18] = ""; salesRow[19] = ""; salesRow[20] = "";
      salesRow[21] = hpStatus;

      salesRow[22] = `${areaText}tel${cleanPhone}`;

      salesRow[23] = ""; salesRow[24] = "";
      salesRow[25] = bizDaysVal;
      salesRow[26] = holidayVal;

      salesRow[27] = finalOpenA;
      salesRow[28] = finalCloseA;
      salesRow[29] = finalOpenB;
      salesRow[30] = finalCloseB;

      genreContainers[targetGroup].push(salesRow);
    }
  });

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("04_SALES_")) {
      if (ss.getSheets().length > 1) { ss.deleteSheet(sheet); } else { sheet.clearContents(); }
    }
  });

  let formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");

  for (const [genreName, rows] of Object.entries(genreContainers)) {
    if (rows.length > 0) {
      ss.toast(`シート「04_SALES_${genreName}」を生成中...`, "📊 タブ分割処理");
      const sheet = getOrCreateSheet(ss, "04_SALES_" + genreName); sheet.clearContents();
      const finalOutput = [finalHeader, ...rows];
      sheet.getRange(1, 1, finalOutput.length, finalOutput[0].length).setValues(finalOutput);

      let fileName = `【営業リスト】${detectedAreaName}_${genreName}_${formattedDate}.csv`;
      ss.toast(`CSVファイル「${fileName}」をドライブへ自動保存中...`, "📂 ログ保存システム");

      let csvStringText = convertArrayToCsvText(finalOutput);
      let bom = "\uFEFF";
      let blob = Utilities.newBlob(bom + csvStringText, "text/csv", fileName);

      let existingFiles = exportFolder.getFilesByName(fileName);
      while (existingFiles.hasNext()) { existingFiles.next().setTrashed(true); }

      exportFolder.createFile(blob);
    }
  }
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

function determineGenreGroup(genreStr, storeName, targetGenres) {
  if (!genreStr) return "その他";
  const n = String(genreStr).normalize("NFC");
  const s = String(storeName).normalize("NFC");
  const mapped = normalizeSystemGenre(n, "", n, s);
  if (targetGenres.indexOf(mapped) !== -1) return mapped;

  if (n.includes("ファミレス") || n.includes("ファミリーレストラン")) return "SKIP";
  if (s.includes("スナック") || s.includes("ラウンジ") || s.includes("Lounge") || n.includes("スナック")) return "スナック";
  if (n.includes("弁当") || n.includes("べんとう") || n.includes("仕出し") || s.includes("弁当")) return "弁当";
  if (n.includes("テイクアウト") || n.includes("持ち帰り")) return "テイクアウト専門店";
  if (n.includes("ハンバーガー")) return "ハンバーガー";
  if (n.includes("美容院") || n.includes("美容室") || n.includes("ヘアサロン")) return "美容院";
  if (n.includes("とんかつ") || n.includes("沖縄料理") || n.includes("しゃぶしゃぶ")) return "和食";
  if (n.includes("お好み焼き") || n.includes("お好み焼") || n.includes("もんじゃ")) return "お好み焼き";
  if (n.includes("焼き鳥") || n.includes("焼鳥") || n.includes("串揚げ") || n.includes("串かつ") || n.includes("串カツ")) return "焼き鳥";
  if (n.includes("寿司") || n.includes("すし")) return "寿司";
  if (n.includes("海鮮") || n.includes("魚介")) return "和食";
  if (n.includes("焼肉") || n.includes("焼き肉") || n.includes("ホルモン") || n.includes("ステーキ") || n.includes("ハンバーグ") || n.includes("韓国バーベキュー")) return "焼肉";
  if (n.includes("イタリアン") || n.includes("パスタ") || n.includes("ピザ")) return "洋食";
  if (n.includes("フレンチ") || n.includes("フランス料理")) return "洋食";
  if (n.includes("和食") || n.includes("日本料理") || n.includes("天ぷら") || n.includes("うなぎ") || n.includes("鍋") || n.includes("もつ鍋") || n.includes("すき焼き") || n.includes("水炊き")) return "和食";
  if (n.includes("そば") || n.includes("蕎麦") || n.includes("うどん")) return "蕎麦・うどん";
  if (n.includes("ラーメン") || n.includes("つけ麺") || n.includes("油そば") || n.includes("まぜそば")) return "ラーメン";
  if (n.includes("カレー")) return "洋食";
  if (n.includes("中華") || n.includes("餃子") || n.includes("台湾") || n.includes("四川") || n.includes("担々麺") || n.includes("肉まん") || n.includes("飲茶") || n.includes("点心")) return "中華";
  if (n.includes("韓国料理") || n.includes("韓国")) return "韓国";
  if (n.includes("食堂") || n.includes("定食")) return "定食・食堂";
  if (n.includes("居酒屋")) return "居酒屋";
  if (n.includes("バー") || n.includes("酒") || n.includes("バル")) return "Bar";
  if (n.includes("スナック")) return "スナック";
  if (n.includes("喫茶店")) return "喫茶店";
  if (n.includes("カフェ") || n.includes("喫茶") || n.includes("コーヒーショップ")) return "カフェ";
  if (n.includes("パン") || n.includes("ベーカリー")) return "パン屋";
  if (n.includes("スイーツ") || n.includes("デザート") || n.includes("ケーキ") || n.includes("洋菓子") || n.includes("和菓子") || n.includes("タピオカ")) return "スイーツ";
  if (n.includes("洋食") || n.includes("スペイン") || n.includes("ビストロ")) return "洋食";
  return "その他";
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

function detectDownloadTargetArea(header, rows) {
  const prefIdx = header.indexOf("都道府県");
  const cityIdx = header.indexOf("市区町村");
  const addrIdx = header.indexOf("住所");
  const areaCounter = {};

  rows.forEach(row => {
    let pref = prefIdx !== -1 ? textValue(row[prefIdx]) : "";
    let city = cityIdx !== -1 ? textValue(row[cityIdx]) : "";

    if ((!pref || !city) && addrIdx !== -1) {
      const parsed = parsePrefCityFromAddress(row[addrIdx]);
      pref = pref || parsed.pref;
      city = city || parsed.city;
    }

    if (!pref || !city) return;
    const key = pref + "\u0001" + city;
    areaCounter[key] = (areaCounter[key] || 0) + 1;
  });

  let bestKey = "";
  let bestCount = 0;
  Object.keys(areaCounter).forEach(key => {
    if (areaCounter[key] > bestCount) {
      bestKey = key;
      bestCount = areaCounter[key];
    }
  });

  if (!bestKey) return { pref: "", city: "", areaText: "" };
  const parts = bestKey.split("\u0001");
  return { pref: parts[0], city: parts[1], areaText: parts[0] + parts[1] };
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

function writeToChainCandidateHook(ss, candidateLog) {
  const candidateSheet = getOrCreateSheet(ss, "CHAIN_CANDIDATE");
  const existingValues = candidateSheet.getDataRange().getValues();
  let existingBrands = existingValues.length > 0 ? new Set(existingValues.slice(1).map(r => String(r[0]).trim())) : new Set(candidateSheet.appendRow(["候補ブランド名", "データ内出現数", "初回検出日時", "処理ステータス"]));
  const appendRows = [];
  for (const [brand, count] of Object.entries(candidateLog)) { if (!existingBrands.has(brand)) { appendRows.push([brand, count, new Date(), "未追加"]); } }
  if (appendRows.length > 0) { candidateSheet.getRange(candidateSheet.getLastRow() + 1, 1, appendRows.length, appendRows[0].length).setValues(appendRows); }
}

// =====================================================================
// 追加処理: Googleマップ取得漏れ分析・除外理由管理
// 既存のVer9.7.0構造を残し、後段タブとして分析結果を追加する
// =====================================================================
const GOOGLEMAP_OPTIONAL_HEADERS = ["検索ジャンル", "取得ステータス", "除外理由", "詳細取得リトライ回数", "一覧取得順", "取得元ジャンル"];
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

const HD_GENRE_MAP = {
  "食堂": "定食・食堂",
  "定食": "定食・食堂",
  "そばうどん": "蕎麦・うどん",
  "そば": "蕎麦・うどん",
  "蕎麦": "蕎麦・うどん",
  "うどん": "蕎麦・うどん",
  "バー": "Bar",
  "パン": "パン屋",
  "ベーカリー": "パン屋",
  "お弁当": "弁当",
  "弁当屋": "弁当",
  "中華料理": "中華",
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
  "焼鳥": "焼き鳥",
  "焼きとり": "焼き鳥",
  "テイクアウト": "テイクアウト専門店"
};

const HD_TARGET_GENRES = [
  "カフェ",
  "居酒屋",
  "スナック",
  "Bar",
  "パン屋",
  "焼き鳥",
  "喫茶店",
  "お好み焼き",
  "焼肉",
  "スイーツ",
  "美容院",
  "中華",
  "ハンバーガー",
  "蕎麦・うどん",
  "寿司",
  "和食",
  "洋食",
  "定食・食堂",
  "弁当",
  "韓国",
  "テイクアウト専門店",
  "ラーメン"
];

const CAFE_KEYWORDS = [
  "カフェ", "Cafe", "CAFE", "cafe", "喫茶", "珈琲", "コーヒー", "coffee", "Coffee", "COFFEE",
  "コーヒーショップ", "カフェテリア", "ドッグカフェ", "コーヒー焙煎所"
];

const FACILITY_EXCLUDE_KEYWORDS = [
  "イオンモール", "イオン", "AEON", "ららぽーと", "アリオ", "パルコ", "PARCO", "ルミネ", "LUMINE",
  "アトレ", "エキュート", "マルイ", "OIOI", "百貨店", "高島屋", "伊勢丹", "三越", "そごう",
  "大丸", "松坂屋", "阪急", "近鉄", "ショッピングセンター", "ショッピングモール", "アウトレット",
  "フードコート", "駅ビル", "ホテル", "病院", "大学", "学校", "スーパー", "ホームセンター",
  "ドン・キホーテ", "ドンキ", "ヨーカドー", "イトーヨーカドー", "アピタ", "ピアゴ", "西友", "ライフ", "マックスバリュ"
];

const FACILITY_REVIEW_KEYWORDS = ["ビル", "プラザ", "タワー", "センター", "テナント", "B1F", "1F", "2F", "3F", "4F", "5F", "階", "地下"];

function executeGoogleMapLeakAnalysis() {
  executeNormalizeAndValidate();
  executeFacilityCheck();
  executeWorkflowGrouping();
  executeSplitSheets();
  executeGenerateComdeskCsv();
  executeProcessLogSummary();
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

    const storeName = getRowValueByHeader(header, row, "店名");
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
    const facility = judgeFacilityStatus(getRowValueByHeader(header, row, "店名"), getRowValueByHeader(header, row, "住所"));
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
    "04_取得失敗": [],
    "確認_小規模ビル": [],
    "確認_電話番号なし": [],
    "確認_ジャンル確認": [],
    "確認_住所確認": [],
    "除外_重複": [],
    "除外_チェーン店": [],
    "除外_ビル管理": [],
    "除外_エリア外": [],
    "除外_住所未取得": []
  };

  rows.forEach(row => {
    const reasonText = [
      getRowValueByHeader(header, row, "除外理由"),
      getRowValueByHeader(header, row, "基本データ除外理由"),
      getRowValueByHeader(header, row, "営業対象除外理由"),
      getRowValueByHeader(header, row, "施設判定理由")
    ].join(" / ");
    const dupStatus = getRowValueByHeader(header, row, "重複判定");
    const chainStatus = getRowValueByHeader(header, row, "チェーン判定");
    const facilityStatus = getRowValueByHeader(header, row, "施設判定");
    const areaStatus = getRowValueByHeader(header, row, "エリア判定");
    const addressStatus = getRowValueByHeader(header, row, "住所判定");
    const normalizedPhone = getRowValueByHeader(header, row, "正規化電話番号") || normalizePhoneNumberForAnalysis(getRowValueByHeader(header, row, "電話番号"));
    const normalizedGenre = getRowValueByHeader(header, row, "正規化ジャンル") || normalizeSystemGenre(getRowValueByHeader(header, row, "ジャンル"), getRowValueByHeader(header, row, "検索ジャンル"), getRowValueByHeader(header, row, "取得元ジャンル"), getRowValueByHeader(header, row, "店名"));
    const workflowGroup = getRowValueByHeader(header, row, "ワークフローグループ") || judgeWorkflowGroup(header, row).group;

    if (buckets[workflowGroup]) buckets[workflowGroup].push(row);
    if (facilityStatus === "確認対象") buckets["確認_小規模ビル"].push(row);
    if (!textValue(normalizedPhone)) buckets["確認_電話番号なし"].push(row);
    if (!isValidHdGenre(normalizedGenre) || reasonText.indexOf("ジャンル確認") !== -1) buckets["確認_ジャンル確認"].push(row);
    if (addressStatus !== "住所あり" || reasonText.indexOf("住所確認") !== -1 || reasonText.indexOf("簡易住所") !== -1) buckets["確認_住所確認"].push(row);
    if (dupStatus !== "ユニーク") buckets["除外_重複"].push(row);
    if (chainStatus === "チェーン店") buckets["除外_チェーン店"].push(row);
    if (facilityStatus === "除外") buckets["除外_ビル管理"].push(row);
    if (areaStatus === "エリア外") buckets["除外_エリア外"].push(row);
    if (addressStatus === "住所未取得" || reasonText.indexOf("住所未取得") !== -1) buckets["除外_住所未取得"].push(row);
  });

  Object.keys(buckets).forEach(name => {
    writeRowsToSheetByName(ss, name, header, buckets[name]);
  });

  writeFetchSummarySheet(ss, header, rows);
  writeExclusionReasonSheets(ss, header, rows);
}

function executeGenerateComdeskCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("01_営業対象");
  const targetSheet = getOrCreateSheet(ss, "05_コムデスク投入用");
  const finalHeader = getComdeskHeader();

  if (!sourceSheet) {
    writeRowsToExistingSheet(targetSheet, finalHeader, []);
    return;
  }

  const values = sourceSheet.getDataRange().getValues();
  if (values.length <= 1) {
    writeRowsToExistingSheet(targetSheet, finalHeader, []);
    return;
  }

  const header = normalizeHeaderRow(values[0]);
  const rows = values.slice(1).filter(row => {
    return getRowValueByHeader(header, row, "ワークフローグループ") === "01_営業対象" && isComdeskTargetRow(header, row);
  });
  const outputRows = rows.map(row => buildComdeskRow(header, row));

  writeRowsToExistingSheet(targetSheet, finalHeader, outputRows);
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

function executeProcessLogSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ["01_営業対象", "02_確認対象", "03_除外対象", "04_取得失敗", "05_コムデスク投入用", "除外_重複", "除外_チェーン店", "除外_ビル管理", "除外_エリア外", "除外_住所未取得", "確認_電話番号なし", "確認_ジャンル確認"];
  const counts = {};
  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    counts[name] = sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;
  });
  const sourceSheet = ss.getSheetByName("01_NORMALIZED");
  appendProcessLog({
    importedCount: sourceSheet ? Math.max(sourceSheet.getLastRow() - 1, 0) : 0,
    targetCount: counts["01_営業対象"],
    reviewCount: counts["02_確認対象"],
    duplicateCount: counts["除外_重複"],
    chainCount: counts["除外_チェーン店"],
    facilityCount: counts["除外_ビル管理"],
    areaOutCount: counts["除外_エリア外"],
    missingAddressCount: counts["除外_住所未取得"],
    missingPhoneCount: counts["確認_電話番号なし"],
    genreReviewCount: counts["確認_ジャンル確認"],
    fetchFailureCount: counts["04_取得失敗"],
    csvCount: counts["05_コムデスク投入用"],
    message: "ワークフローグループ判定完了"
  });
}

function appendProcessLog(log) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "PROCESS_LOG");
  const header = ["実行日時", "処理名", "取込件数", "営業対象件数", "確認対象件数", "重複件数", "チェーン除外件数", "ビル管理除外件数", "エリア外件数", "住所未取得件数", "電話番号なし件数", "ジャンル確認件数", "取得失敗件数", "CSV出力件数", "メッセージ"];
  if (sheet.getLastRow() === 0) sheet.appendRow(header);
  sheet.appendRow([
    new Date(), "executeAllProcesses", log.importedCount || 0, log.targetCount || 0, log.reviewCount || 0,
    log.duplicateCount || 0, log.chainCount || 0, log.facilityCount || 0, log.areaOutCount || 0,
    log.missingAddressCount || 0, log.missingPhoneCount || 0, log.genreReviewCount || 0,
    log.fetchFailureCount || 0, log.csvCount || 0, log.message || ""
  ]);
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
  const cleanPhone = normalizePhoneNumberForAnalysis(phone).replace(/[^\d]/g, "");
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
  const address1 = city ? city + fullAddr.replace(pref, "").replace(city, "") : addrDetails.addr1;
  const areaText = pref + city;

  const salesRow = Array(31).fill("");
  salesRow[2] = storeName;
  salesRow[4] = addrDetails.pcode;
  salesRow[5] = pref;
  salesRow[6] = address1;
  salesRow[9] = phone;
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

function normalizeSystemGenre(genre, searchGenre, sourceGenre, storeName) {
  const rawGenre = textValue(genre);
  let mappedGenre = HD_GENRE_MAP[rawGenre] || rawGenre;
  const haystack = [sourceGenre, storeName, rawGenre].map(textValue).join(" ");
  if (textValue(searchGenre) === "カフェ" && CAFE_KEYWORDS.some(keyword => haystack.indexOf(keyword) !== -1)) return "カフェ";
  if (!isValidHdGenre(mappedGenre)) {
    const matchedOldGenre = Object.keys(HD_GENRE_MAP).find(oldGenre => rawGenre.indexOf(oldGenre) !== -1 || textValue(sourceGenre).indexOf(oldGenre) !== -1);
    if (matchedOldGenre) mappedGenre = HD_GENRE_MAP[matchedOldGenre];
  }
  return mappedGenre;
}

function judgeFacilityStatus(storeName, address) {
  const haystack = [storeName, address].map(textValue).join(" ");
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

function writeFetchSummarySheet(ss, header, rows) {
  const summaryHeader = ["媒体", "都道府県", "市区町村", "検索ジャンル", "ジャンル", "取得ステータス", "除外理由", "件数"];
  const counter = {};
  rows.forEach(row => {
    const keyValues = [
      getRowValueByHeader(header, row, "媒体"),
      getRowValueByHeader(header, row, "都道府県"),
      getRowValueByHeader(header, row, "市区町村"),
      getRowValueByHeader(header, row, "検索ジャンル"),
      getRowValueByHeader(header, row, "正規化ジャンル") || getRowValueByHeader(header, row, "ジャンル"),
      getRowValueByHeader(header, row, "取得ステータス") || "未設定",
      [getRowValueByHeader(header, row, "除外理由"), getRowValueByHeader(header, row, "営業対象除外理由")].filter(Boolean).join(" / ")
    ];
    const key = keyValues.join("\u0001");
    counter[key] = (counter[key] || 0) + 1;
  });
  const outputRows = Object.keys(counter).map(key => key.split("\u0001").concat([counter[key]]));
  writeRowsToSheetByName(ss, "分析_取得状況サマリー", summaryHeader, outputRows);
}

function writeExclusionReasonSheets(ss, header, rows) {
  const outputHeader = ["店名", "ジャンル", "検索ジャンル", "取得元ジャンル", "都道府県", "市区町村", "住所", "電話番号", "URL", "媒体", "取得ステータス", "除外理由", "営業対象判定", "営業対象除外理由", "ワークフローグループ", "ワークフロー項目", "対応ステータス", "次アクション"];
  const toOutput = sourceRows => sourceRows.map(row => outputHeader.map(name => getRowValueByHeader(header, row, name)));
  const exclusionRows = rows.filter(row => getRowValueByHeader(header, row, "除外理由") || getRowValueByHeader(header, row, "営業対象判定") === "除外" || getRowValueByHeader(header, row, "取得ステータス") === "失敗");
  const genreRows = rows.filter(row => {
    const searchGenre = getRowValueByHeader(header, row, "検索ジャンル");
    const normalizedGenre = getRowValueByHeader(header, row, "正規化ジャンル");
    return !isValidHdGenre(normalizedGenre) || (searchGenre && normalizedGenre && searchGenre !== normalizedGenre);
  });
  const missingAddressRows = rows.filter(row => getRowValueByHeader(header, row, "住所判定") !== "住所あり");
  const detailFailureRows = rows.filter(row => {
    const reasons = [getRowValueByHeader(header, row, "除外理由"), getRowValueByHeader(header, row, "営業対象除外理由")].join(" / ");
    return getRowValueByHeader(header, row, "取得ステータス") === "失敗" || reasons.indexOf("詳細取得失敗") !== -1 || (!getRowValueByHeader(header, row, "店名") && getRowValueByHeader(header, row, "URL"));
  });

  writeRowsToSheetByName(ss, "分析_除外理由別一覧", outputHeader, toOutput(exclusionRows));
  writeRowsToSheetByName(ss, "分析_ジャンル不一致一覧", outputHeader, toOutput(genreRows));
  writeRowsToSheetByName(ss, "確認_住所確認", header, missingAddressRows);
  writeRowsToSheetByName(ss, "分析_詳細取得失敗一覧", outputHeader, toOutput(detailFailureRows));
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
