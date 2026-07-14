/**
 * 花屋 店舗管理システム Ver1.0.0
 * 生花店・フラワーショップ専用。ホームセンター・スーパー・JA直売所等の
 * 花コーナー、および全国チェーンの生花店を除外して、独立系の生花店・
 * フラワーショップだけを営業リスト化するGAS。
 * `gas/apparel-store/code-apparel.js`（アパレル側）と同じ構成・思想・
 * 命名規則で作られている（チェーン判定の3段階化ロジックも同様に採用）。
 *
 * ★このバージョンについて（重要）
 * アパレル版のチェーンマスタ（APPAREL_KNOWN_CHAIN_GAPS）は過去の実データ
 * 検証を重ねて拡充してきたが、花屋業態は今回が初めてのため、
 * FLOWER_KNOWN_CHAIN_GAPSは「確度の高い全国規模のチェーンのみ」を
 * 最小限登録した一次リストになっている（日比谷花壇・青山フラワーマーケット・
 * 花キューピット等）。ホームセンター・スーパー・JA直売所等の「花コーナー」は
 * FLOWER_FACILITY_EXCLUDE_KEYWORDSで除外するので、チェーンマスタが薄くても
 * 大きな取りこぼしにはなりにくい設計だが、実際のCSVで独立店に見えて実は
 * チェーンだった店名等があれば教えてほしい。マスタに追記して精度を上げていく。
 *
 * 運用：
 * 1. 「花屋_CSV投入フォルダ」にCSVを入れる
 * 2. スプレッドシートを再読み込み
 * 3. メニュー「💐 花屋」→「🚀 すべての一括処理を実行」
 */

// =====================================================================
// メニュー
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💐 花屋")
    .addItem("🚀 すべての一括処理を実行", "flowerExecuteAllProcesses")
    .addSeparator()
    .addItem("0. 初期フォルダ・タブ作成", "flowerSetupAll")
    .addItem("📁 1. CSVを一括取り込み", "flowerImportCSVFiles")
    .addItem("2. 正規化・基本判定", "flowerExecuteNormalizeAndValidate")
    .addItem("3. 重複判定", "flowerExecuteDuplicateCheck")
    .addItem("4. チェーン判定", "flowerExecuteChainCheck")
    .addItem("5. 施設判定", "flowerExecuteFacilityCheck")
    .addItem("6. ワークフロー分類", "flowerExecuteWorkflowGrouping")
    .addItem("7. タブ分け", "flowerExecuteSplitSheets")
    .addItem("8. 04_SALES_地域別タブ生成", "flowerExecuteGenerateSalesAreaSheets")
    .addItem("9. 04_SALES_CSVをDrive出力", "flowerExecuteExportSalesAreaCsvFiles")
    .addSeparator()
    .addItem("📊 件数サマリーを更新", "flowerExecuteCountSummary")
    .addItem("🔧 チェーンマスタの不足キーワードを追加", "flowerFixKnownChainMasterGaps")
    .addToUi();
}

// =====================================================================
// 設定
// =====================================================================
const FLOWER_FOLDER_NAMES = {
  input: "花屋_CSV投入フォルダ",
  processed: "花屋_処理済みフォルダ",
  export: "完成版CSVエクスポート"
};

const FLOWER_SHEETS = {
  normalized: "花屋_01_NORMALIZED",
  duplicate: "花屋_02_DUPLICATE_CHECK",
  chain: "花屋_03_CHAIN_CHECK",
  facility: "花屋_04_FACILITY_CHECK",
  target: "花屋_01_営業対象",
  confirm: "花屋_02_確認対象",
  exclude: "花屋_03_除外対象",
  failed: "花屋_04_取得失敗",
  summary: "花屋_00_件数サマリー",
  masterChain: "花屋_MASTER_CHAIN"
};

const FLOWER_NORMALIZED_HEADER = [
  "店名", "ジャンル", "検索ジャンル", "取得元ジャンル", "都道府県", "市区町村", "郵便番号", "住所", "電話番号", "URL", "媒体",
  "HP有無", "営業日", "定休日", "営業開始A", "営業終了A", "営業開始B", "営業終了B", "取得ステータス", "除外理由",
  "正規化電話番号", "正規化店名", "正規化ジャンル", "住所判定", "エリア判定", "エリア判定理由",
  "基本データ判定", "基本データ除外理由"
];

const FLOWER_COMDESK_HEADER = [
  "UUID", "種別", "名前", "カナ", "郵便番号", "都道府県", "住所１", "住所２", "住所カナ",
  "Tel1", "Tel2", "Tel3", "Tel4", "FAX", "URL", "備考", "旧社名", "リードソース",
  "旧進捗", "履歴", "オーナー名", "HPある？", "BP検索", "アポ済商材", "最新履歴", "営業曜日", "休業曜日",
  "午前始", "午前終", "午後始", "午後終"
];

// 検索ジャンル・取得元ジャンルの表記ゆれ → システム標準ジャンルへのマッピング
const FLOWER_GENRE_MAP = {
  "花屋": "生花店・フラワーショップ",
  "生花店": "生花店・フラワーショップ",
  "生花": "生花店・フラワーショップ",
  "フラワーショップ": "生花店・フラワーショップ",
  "フラワー": "生花店・フラワーショップ",
  "花き小売店": "生花店・フラワーショップ",
  "花き店": "生花店・フラワーショップ",
  "花卉店": "生花店・フラワーショップ",
  "花き": "生花店・フラワーショップ",
  "花卉": "生花店・フラワーショップ",

  "園芸店": "園芸・ガーデニング",
  "園芸用品店": "園芸・ガーデニング",
  "ガーデニングショップ": "園芸・ガーデニング",
  "ガーデニング": "園芸・ガーデニング",
  "植木屋": "園芸・ガーデニング",
  "植木店": "園芸・ガーデニング",
  "苗屋": "園芸・ガーデニング",
  "苗木店": "園芸・ガーデニング",
  "グリーンショップ": "園芸・ガーデニング",
  "観葉植物店": "園芸・ガーデニング",

  "プリザーブドフラワー": "プリザーブド・ドライフラワー",
  "プリザーブドフラワーショップ": "プリザーブド・ドライフラワー",
  "ドライフラワー": "プリザーブド・ドライフラワー",
  "ドライフラワーショップ": "プリザーブド・ドライフラワー",

  "造花": "造花・アートフラワー",
  "造花店": "造花・アートフラワー",
  "アートフラワー": "造花・アートフラワー",
  "アートフラワーショップ": "造花・アートフラワー",

  "花贈答": "花贈答・ギフト",
  "フラワーギフト": "花贈答・ギフト",
  "花ギフトショップ": "花贈答・ギフト",
  "花贈答店": "花贈答・ギフト",

  "ブライダルフラワー": "ブライダルフラワー",
  "ウェディングブーケ": "ブライダルフラワー",
  "ウェディングフラワー": "ブライダルフラワー",
  "ブーケショップ": "ブライダルフラワー",

  "フラワーアレンジメント教室": "その他花き",
  "花き卸売": "その他花き",
  "花卉卸売": "その他花き"
};

const FLOWER_TARGET_GENRES = [
  "生花店・フラワーショップ",
  "園芸・ガーデニング",
  "プリザーブド・ドライフラワー",
  "造花・アートフラワー",
  "花贈答・ギフト",
  "ブライダルフラワー",
  "その他花き"
];

// 施設・大型小売として除外するキーワード（ホームセンター・スーパー・
// JA直売所・道の駅等の「花コーナー」を、独立系フラワーショップと
// 区別して除外するためのもの）
const FLOWER_FACILITY_EXCLUDE_KEYWORDS = [
  // ホームセンター（園芸コーナーは大手チェーンの一部門であり独立店ではない）
  "カインズ", "CAINZ", "コメリ", "DCM", "ケーヨーデイツー", "ケーヨーD2", "コーナン", "ビバホーム",
  "スーパービバホーム", "ジョイフル本田", "島忠", "ホームズ", "ナフコ", "ロイヤルホームセンター", "ホームセンター",
  "グリーンランド", "アグリガーデン",

  // スーパー・GMS・ディスカウントストア（インストア花コーナー）
  "イオンモール", "イオン", "AEON", "イトーヨーカドー", "ヨーカドー", "西友", "ライフ", "マックスバリュ",
  "ダイエー", "ベイシア", "トライアル", "業務スーパー", "ヤオコー", "カスミ", "ロピア", "ドンキホーテ", "ドンキ",
  "生協", "コープ", "パルシステム",

  // JA・直売所・道の駅（生産者直売の花き即売所。独立系フラワーショップとは業態が異なる）
  "JA直売所", "JAcom", "農協", "農業協同組合", "道の駅", "産直市場", "産直店", "ファーマーズマーケット",
  "とれたて市場", "直売所",

  // 百貨店・ショッピングモール・駅ビル（テナント出店。単独判断が難しいため確認対象止まりではなく除外）
  "百貨店", "高島屋", "伊勢丹", "三越", "そごう", "大丸", "松坂屋", "阪急", "近鉄",
  "ショッピングセンター", "ショッピングモール", "アウトレット", "駅ビル", "ルミネ", "LUMINE", "アトレ",

  // その他生花以外の量販店（誤ヒット防止）
  "コンビニ", "ドラッグストア", "薬局", "書店", "本屋", "家電", "家具", "雑貨店", "生活雑貨"
];

const FLOWER_FACILITY_REVIEW_KEYWORDS = [
  "ビル", "プラザ", "タワー", "センター", "B1F", "1F", "2F", "3F", "4F", "5F", "階", "地下", "市場", "青果"
];

// [チェーン名, キーワード, 業種, 有効, メモ]
// ★確度の高い全国規模のチェーンのみを一次登録。実データで見つかった
// 未収録チェーンがあれば随時ここに追記する（アパレル版の運用と同じ）。
const FLOWER_KNOWN_CHAIN_GAPS = [
  ["日比谷花壇", "日比谷花壇", "花き", true, "業界最大手。駅・百貨店・単独店舗など全国展開"],
  ["日比谷花壇", "HIBIYA-KADAN", "花き", true, "英語表記"],
  ["青山フラワーマーケット", "青山フラワーマーケット", "花き", true, "パーク・コーポレーション系。全国の駅・商業施設に多数出店"],
  ["青山フラワーマーケット", "Aoyama Flower Market", "花き", true, "英語表記"],
  ["花キューピット", "花キューピット", "花き", true, "全国の独立系花屋が加盟するフラワーギフト配送ネットワーク（加盟店の看板・店名に併記されることが多い）"],
  ["花キューピット", "Hana Cupid", "花き", true, "英語表記"],
  ["イーフローラ", "イーフローラ", "花き", true, "フラワーギフト配送ネットワーク（花キューピットと同様、加盟店表記に注意）"],
  ["イーフローラ", "e-flora", "花き", true, "英語表記"],
  ["Hitohana", "Hitohana", "花き", true, "オンライン系フラワーギフトブランド。実店舗を持つ場合がある"],
  ["ローズ&ペア", "ローズ&ペア", "花き", true, "全国チェーン展開の実店舗型フラワーショップ"]
];

// =====================================================================
// 一括処理
// =====================================================================
function flowerExecuteAllProcesses() {
  flowerSetupAll();
  flowerImportCSVFiles();
  flowerExecuteNormalizeAndValidate();
  flowerExecuteDuplicateCheck();
  flowerAddKnownChainMasterGapsSilently();
  flowerExecuteChainCheck();
  flowerExecuteFacilityCheck();
  flowerExecuteWorkflowGrouping();
  flowerExecuteSplitSheets();
  flowerExecuteGenerateSalesAreaSheets();
  flowerExecuteExportSalesAreaCsvFiles();
  const summary = flowerExecuteCountSummary();

  SpreadsheetApp.getUi().alert(
    "花屋リスト処理が完了しました。\n\n" +
    "営業対象: " + summary.totalTarget + "件\n" +
    "確認対象: " + summary.totalConfirm + "件\n" +
    "除外対象: " + summary.totalExclude + "件\n" +
    "取得失敗: " + summary.totalFailed + "件\n\n" +
    "花屋_04_SALES_地域別タブと完成版CSVエクスポートを確認してください。\n\n" +
    "※「チェーン店疑い」（地名付きの支店名パターンで自動検出したが、マスタには" +
    "未登録の店舗）は確認対象に入っています。本当にチェーン店か目視確認してください。"
  );
}

function flowerSetupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(FLOWER_SHEETS).forEach(key => flowerGetOrCreateSheet_(ss, FLOWER_SHEETS[key]));
  flowerCreateFolders_();
  flowerAddKnownChainMasterGapsSilently();
  ss.toast("花屋用フォルダ・タブを確認しました。", "💐 花屋");
}

function flowerCreateFolders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  flowerGetOrCreateFolder_(parentFolder, FLOWER_FOLDER_NAMES.input);
  flowerGetOrCreateFolder_(parentFolder, FLOWER_FOLDER_NAMES.processed);
  flowerGetOrCreateFolder_(parentFolder, FLOWER_FOLDER_NAMES.export);
}

// =====================================================================
// 1. CSV取り込み
// =====================================================================
function flowerImportCSVFiles() {
  flowerCreateFolders_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = flowerGetOrCreateSheet_(ss, FLOWER_SHEETS.normalized);
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const importFolder = flowerGetOrCreateFolder_(parentFolder, FLOWER_FOLDER_NAMES.input);
  const processedFolder = flowerGetOrCreateFolder_(parentFolder, FLOWER_FOLDER_NAMES.processed);
  const files = importFolder.getFilesByType(MimeType.CSV);

  const parsedFiles = [];
  const unifiedHeader = [];
  const unifiedHeaderSet = new Set();

  while (files.hasNext()) {
    const file = files.next();
    try {
      ss.toast("ファイル「" + file.getName() + "」を解析中...", "📁 CSV一括取り込み");
      const blob = file.getBlob();
      let csvText = blob.getDataAsString("UTF-8").replace(/^﻿/, "");
      if (!csvText.includes("店名") && !csvText.includes("名前")) {
        csvText = blob.getDataAsString("MS932").replace(/^﻿/, "");
      }

      const parsedCsv = Utilities.parseCsv(csvText);
      if (!parsedCsv || parsedCsv.length === 0) {
        file.moveTo(processedFolder);
        continue;
      }

      const fileHeader = parsedCsv[0].map(h => String(h || "").replace(/^﻿/, "").trim());
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
      Logger.log("[花屋CSV取込エラー] " + file.getName() + ": " + e.message);
      ss.toast("「" + file.getName() + "」の読み込みに失敗しました: " + e.message, "⚠️ CSV取込エラー");
    }
  }

  if (parsedFiles.length === 0) {
    ss.toast("新規CSVはありません。既存データがあれば後続処理は可能です。", "📁 CSV一括取り込み");
    return { files: 0, rows: 0 };
  }

  const combinedData = [unifiedHeader];
  parsedFiles.forEach(pf => {
    const colIndexInFile = pf.header.map(col => unifiedHeader.indexOf(col));
    pf.rows.forEach(row => {
      const alignedRow = new Array(unifiedHeader.length).fill("");
      row.forEach((value, i) => {
        const targetIdx = colIndexInFile[i];
        if (targetIdx !== undefined && targetIdx !== -1) alignedRow[targetIdx] = value;
      });
      combinedData.push(alignedRow);
    });
  });

  flowerWriteRowsToExistingSheet_(targetSheet, combinedData[0], combinedData.slice(1));
  return { files: parsedFiles.length, rows: combinedData.length - 1 };
}

// =====================================================================
// 2. 正規化・基本判定
// =====================================================================
function flowerExecuteNormalizeAndValidate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FLOWER_SHEETS.normalized);
  if (!sheet || sheet.getLastRow() <= 1) return { count: 0 };

  const values = sheet.getDataRange().getValues();
  const rawHeader = flowerNormalizeHeaderRow_(values[0]);
  const rawRows = values.slice(1);
  const outputRows = [];

  rawRows.forEach(row => {
    const rawStoreName = flowerGetValue_(rawHeader, row, ["店名", "店舗名", "名前", "施設名", "タイトル", "name", "Name"]);
    const storeName = flowerStripGenreSuffixFromName_(rawStoreName);
    const genre = flowerGetValue_(rawHeader, row, ["ジャンル", "カテゴリ", "カテゴリー", "業種", "種別", "category", "types"]);
    const searchGenre = flowerGetValue_(rawHeader, row, ["検索ジャンル", "検索キーワード", "取得元ジャンル", "keyword"]);
    const sourceGenre = flowerGetValue_(rawHeader, row, ["取得元ジャンル", "元ジャンル", "sourceGenre"]);
    const address = flowerGetValue_(rawHeader, row, ["住所", "所在地", "住所１", "住所1", "address", "Address", "formatted_address"]);
    const parsedAddress = flowerParsePrefCityFromAddress_(address);
    const pref = flowerGetValue_(rawHeader, row, ["都道府県", "pref", "Prefecture"]) || parsedAddress.pref;
    const city = flowerGetValue_(rawHeader, row, ["市区町村", "市町村", "city", "City"]) || parsedAddress.city;
    const zip = flowerParseAddressDetails_(address).pcode;
    const phone = flowerNormalizePhoneDisplay_(flowerGetValue_(rawHeader, row, ["電話番号", "TEL", "Tel", "tel", "電話", "phone", "Phone", "Tel1"]));
    const url = flowerGetValue_(rawHeader, row, ["URL", "Webサイト", "ホームページ", "サイト", "website", "Website", "リンク"]);
    const media = flowerGetValue_(rawHeader, row, ["媒体", "取得元", "source", "Source"]) || "Googleマップ";
    const hpHave = flowerGetValue_(rawHeader, row, ["HP有無", "HPある？", "ホームページ有無"]);
    const businessDays = flowerGetValue_(rawHeader, row, ["営業日", "営業曜日"]);
    const holiday = flowerGetValue_(rawHeader, row, ["定休日", "休業曜日"]);
    const openA = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetValue_(rawHeader, row, ["営業開始A", "営業開始", "午前始"])));
    const closeA = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetValue_(rawHeader, row, ["営業終了A", "営業終了", "午前終"])));
    const openB = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetValue_(rawHeader, row, ["営業開始B", "午後始"])));
    const closeB = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetValue_(rawHeader, row, ["営業終了B", "午後終"])));
    const fetchStatus = flowerGetValue_(rawHeader, row, ["取得ステータス", "status"]);
    const externalReason = flowerGetValue_(rawHeader, row, ["除外理由", "理由"]);

    const normalizedGenre = flowerNormalizeSystemGenre_(genre, searchGenre, sourceGenre, storeName);
    const normalizedPhone = flowerNormalizePhoneNumberForAnalysis_(phone);
    const normalizedName = flowerSimplifyStoreName_(storeName);
    const addressStatus = flowerJudgeAddressStatus_(address, pref, city);
    const areaStatus = flowerJudgeAreaStatus_(address, pref, city);

    const basicReasons = [];
    if (!storeName) basicReasons.push("店名なし");
    if (!normalizedPhone) basicReasons.push("電話番号なし");
    if (addressStatus.status !== "住所あり") basicReasons.push(addressStatus.reason);
    if (!normalizedGenre || !flowerIsValidTargetGenre_(normalizedGenre)) basicReasons.push("ジャンル確認");
    if (fetchStatus === "失敗") basicReasons.push("取得失敗");

    outputRows.push([
      storeName, normalizedGenre || genre, searchGenre, sourceGenre, pref, city, zip, address, phone, url, media,
      flowerNormalizeHpStatus_(hpHave, url), businessDays, holiday, openA, closeA, openB, closeB, fetchStatus, externalReason,
      normalizedPhone, normalizedName, normalizedGenre, addressStatus.status, areaStatus.status, areaStatus.reason,
      basicReasons.length === 0 ? "対象" : "確認対象", flowerUniqueTextList_(basicReasons).join(" / ")
    ]);
  });

  flowerWriteRowsToExistingSheet_(sheet, FLOWER_NORMALIZED_HEADER, outputRows);
  return { count: outputRows.length };
}

// =====================================================================
// 3. 重複判定
// =====================================================================
function flowerExecuteDuplicateCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(FLOWER_SHEETS.normalized);
  const targetSheet = flowerGetOrCreateSheet_(ss, FLOWER_SHEETS.duplicate);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = flowerNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const seenPhones = new Set();
  const seenNames = new Set();
  const outputRows = [];

  rows.forEach(row => {
    const rawName = flowerGetRowValueByHeader_(header, row, "店名");
    const rawPhone = flowerNormalizePhoneNumberForAnalysis_(flowerGetRowValueByHeader_(header, row, "電話番号")).replace(/[^\d]/g, "");
    const cleanName = flowerSimplifyStoreName_(rawName);
    let duplicateReason = "ユニーク";

    if (rawPhone && rawPhone.length >= 9) {
      if (seenPhones.has(rawPhone)) duplicateReason = "重複（電話番号一致）";
      else seenPhones.add(rawPhone);
    }

    if (duplicateReason === "ユニーク" && cleanName) {
      if (seenNames.has(cleanName)) duplicateReason = "重複（店舗名一致）";
      else seenNames.add(cleanName);
    }

    outputRows.push(row.concat([duplicateReason]));
  });

  flowerWriteRowsToExistingSheet_(targetSheet, header.concat(["重複判定"]), outputRows);
  return { count: outputRows.length };
}

// =====================================================================
// 4. チェーン判定
// =====================================================================
// アパレル版と同じ3段階判定：
//   "チェーン店"      … MASTER_CHAINへの明示登録、または号店・駅前店・
//                       モール系テナント名などの強いシグナルに一致（確実）
//   "チェーン店疑い"  … 地名＋「店」、スペース/括弧区切りの「〇〇店」など、
//                       チェーンらしい命名パターンではあるが確証がない
//   "単独店"          … いずれにも一致しない
function flowerExecuteChainCheck() {
  flowerAddKnownChainMasterGapsSilently();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(FLOWER_SHEETS.masterChain);
  const sourceSheet = ss.getSheetByName(FLOWER_SHEETS.duplicate);
  const targetSheet = flowerGetOrCreateSheet_(ss, FLOWER_SHEETS.chain);
  if (!masterSheet || !sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const masterValues = masterSheet.getDataRange().getValues();
  const chainMaster = [];
  for (let i = 1; i < masterValues.length; i++) {
    const chainName = flowerTextValue_(masterValues[i][0]);
    const keyword = flowerTextValue_(masterValues[i][1]);
    const isValid = masterValues[i][3];
    if (isValid && keyword) {
      chainMaster.push({
        chainName,
        keyword,
        normalizedKeyword: flowerNormalizeForKeywordMatch_(keyword)
      });
    }
  }

  const sourceValues = sourceSheet.getDataRange().getValues();
  const header = flowerNormalizeHeaderRow_(sourceValues[0]);
  const rows = sourceValues.slice(1);
  const outputRows = [];

  rows.forEach(row => {
    const storeName = flowerGetRowValueByHeader_(header, row, "店名");
    const genre = flowerGetRowValueByHeader_(header, row, "ジャンル");
    const searchGenre = flowerGetRowValueByHeader_(header, row, "検索ジャンル");
    const sourceGenre = flowerGetRowValueByHeader_(header, row, "取得元ジャンル");
    const haystackRaw = [storeName, genre, searchGenre, sourceGenre].join(" ");
    const haystack = flowerNormalizeForKeywordMatch_(haystackRaw);

    let chainStatus = "単独店";
    let matchedChainName = "";
    let chainReason = "単独店確認";

    for (const master of chainMaster) {
      if (haystack.indexOf(master.normalizedKeyword) !== -1) {
        chainStatus = "チェーン店";
        matchedChainName = master.chainName;
        chainReason = "マスタ合致: キーワード[" + master.keyword + "]";
        break;
      }
    }

    if (chainStatus === "単独店") {
      const branchLevel = flowerIsLikelyBranchStoreName_(storeName);

      if (branchLevel === "high") {
        chainStatus = "チェーン店";
        matchedChainName = "支店名付き店舗";
        chainReason = "自動検出: 支店名形式（駅・号店・モール等の強いシグナル）";
      } else if (branchLevel === "heuristic") {
        chainStatus = "チェーン店疑い";
        matchedChainName = "支店名付き店舗（推定）";
        chainReason = "自動検出: 地名や区切り記号付きの支店名パターン（単独店の可能性もあるため要確認）";
      }
    }

    outputRows.push(row.concat([
      chainStatus,
      matchedChainName,
      chainReason
    ]));
  });

  flowerWriteRowsToExistingSheet_(targetSheet, header.concat(["チェーン判定", "チェーン名", "チェーン理由"]), outputRows);
  return { count: outputRows.length };
}

function flowerAddKnownChainMasterGapsSilently() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = ss.getSheetByName(FLOWER_SHEETS.masterChain);
  if (!masterSheet) {
    masterSheet = flowerGetOrCreateSheet_(ss, FLOWER_SHEETS.masterChain);
    masterSheet.getRange(1, 1, 1, 5).setValues([["チェーン名", "キーワード", "業種", "有効", "メモ"]]);
    masterSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }

  const existingValues = masterSheet.getDataRange().getValues();
  const existingKeywords = new Set(existingValues.slice(1).map(r => flowerTextValue_(r[1]).normalize("NFC")));
  const newRows = FLOWER_KNOWN_CHAIN_GAPS.filter(row => {
    const keyword = flowerTextValue_(row[1]).normalize("NFC");
    return keyword && !existingKeywords.has(keyword);
  });

  if (newRows.length === 0) return { added: [] };

  const startRow = masterSheet.getLastRow() + 1;
  masterSheet.getRange(startRow, 1, newRows.length, 5).setNumberFormat("@").setValues(newRows);
  masterSheet.getRange(startRow, 4, newRows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  return { added: newRows };
}

function flowerFixKnownChainMasterGaps() {
  const result = flowerAddKnownChainMasterGapsSilently();
  SpreadsheetApp.getUi().alert(
    result.added.length === 0
      ? "追加対象がありません（すべて登録済みです）。"
      : "【完了】" + result.added.length + "件のキーワードを花屋_MASTER_CHAINへ追加しました。\n\n反映するには「4. チェーン判定」以降を再実行してください。"
  );
}

// =====================================================================
// 5. 施設判定
// =====================================================================
function flowerExecuteFacilityCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(FLOWER_SHEETS.chain) || ss.getSheetByName(FLOWER_SHEETS.duplicate) || ss.getSheetByName(FLOWER_SHEETS.normalized);
  const targetSheet = flowerGetOrCreateSheet_(ss, FLOWER_SHEETS.facility);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = flowerNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const outputRows = [];

  rows.forEach(row => {
    const facility = flowerJudgeFacilityStatus_(
      flowerGetRowValueByHeader_(header, row, "店名"),
      flowerGetRowValueByHeader_(header, row, "住所"),
      flowerGetRowValueByHeader_(header, row, "ジャンル"),
      flowerGetRowValueByHeader_(header, row, "検索ジャンル"),
      flowerGetRowValueByHeader_(header, row, "取得元ジャンル")
    );
    const sales = flowerJudgeSalesTargetStatus_(header, row, facility);
    outputRows.push(row.concat([
      facility.status,
      facility.reason,
      sales.status,
      sales.reason
    ]));
  });

  flowerWriteRowsToExistingSheet_(targetSheet, header.concat(["施設判定", "施設判定理由", "営業対象判定", "営業対象除外理由"]), outputRows);
  return { count: outputRows.length };
}

function flowerJudgeFacilityStatus_(storeName, address, genre, searchGenre, sourceGenre) {
  const haystack = flowerNormalizeForKeywordMatch_([storeName, address, genre, searchGenre, sourceGenre].join(" "));
  const excludeKeyword = FLOWER_FACILITY_EXCLUDE_KEYWORDS.find(keyword => haystack.indexOf(flowerNormalizeForKeywordMatch_(keyword)) !== -1);
  if (excludeKeyword) return { status: "除外", reason: "完全除外キーワード一致: " + excludeKeyword };

  const reviewKeyword = FLOWER_FACILITY_REVIEW_KEYWORDS.find(keyword => haystack.indexOf(flowerNormalizeForKeywordMatch_(keyword)) !== -1);
  if (reviewKeyword) return { status: "確認対象", reason: "確認対象キーワード一致: " + reviewKeyword };

  return { status: "対象", reason: "" };
}

function flowerJudgeSalesTargetStatus_(header, row, facility) {
  const reasons = [];
  let hasReview = false;
  let hasExclude = false;

  const fetchStatus = flowerGetRowValueByHeader_(header, row, "取得ステータス");
  const externalReason = flowerGetRowValueByHeader_(header, row, "除外理由");
  const basicStatus = flowerGetRowValueByHeader_(header, row, "基本データ判定");
  const basicReason = flowerGetRowValueByHeader_(header, row, "基本データ除外理由");
  const areaStatus = flowerGetRowValueByHeader_(header, row, "エリア判定");
  const areaReason = flowerGetRowValueByHeader_(header, row, "エリア判定理由");
  const dupStatus = flowerGetRowValueByHeader_(header, row, "重複判定");
  const chainStatus = flowerGetRowValueByHeader_(header, row, "チェーン判定");

  if (fetchStatus === "失敗") { hasExclude = true; reasons.push("取得失敗"); }
  if (externalReason) { hasExclude = true; reasons.push(externalReason); }
  if (basicStatus !== "対象") { hasReview = true; reasons.push(basicReason || "基本データ確認"); }
  if (areaStatus === "判定不可") { hasReview = true; reasons.push(areaReason || "エリア判定不可"); }
  if (areaStatus === "エリア外") { hasExclude = true; reasons.push(areaReason || "エリア外"); }
  if (dupStatus && dupStatus !== "ユニーク") { hasExclude = true; reasons.push(dupStatus || "重複"); }
  if (chainStatus === "チェーン店") { hasExclude = true; reasons.push("チェーン店"); }
  if (chainStatus === "チェーン店疑い") { hasReview = true; reasons.push("チェーン店疑い（支店名パターン検出。要確認）"); }
  if (facility.status === "確認対象") { hasReview = true; reasons.push(facility.reason); }
  if (facility.status === "除外") { hasExclude = true; reasons.push(facility.reason); }

  const joined = flowerUniqueTextList_(reasons).join(" / ");
  if (!joined) return { status: "対象", reason: "" };
  if (hasReview && !hasExclude) return { status: "確認対象", reason: joined };
  return { status: "除外", reason: joined };
}

// =====================================================================
// 6. ワークフロー分類
// =====================================================================
function flowerExecuteWorkflowGrouping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FLOWER_SHEETS.facility);
  if (!sheet || sheet.getLastRow() <= 1) return { count: 0 };

  const values = sheet.getDataRange().getValues();
  const header = flowerNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const outputRows = [];

  rows.forEach(row => {
    const workflow = flowerJudgeWorkflowGroup_(header, row);
    outputRows.push(row.concat([
      workflow.group,
      workflow.item,
      workflow.status,
      workflow.nextAction
    ]));
  });

  flowerWriteRowsToExistingSheet_(sheet, header.concat(["ワークフローグループ", "ワークフロー項目", "対応ステータス", "次アクション"]), outputRows);
  return { count: outputRows.length };
}

function flowerJudgeWorkflowGroup_(header, row) {
  const storeName = flowerGetRowValueByHeader_(header, row, "店名");
  const address = flowerGetRowValueByHeader_(header, row, "住所");
  const url = flowerGetRowValueByHeader_(header, row, "URL");
  const fetchStatus = flowerGetRowValueByHeader_(header, row, "取得ステータス");
  const dupStatus = flowerGetRowValueByHeader_(header, row, "重複判定");
  const chainStatus = flowerGetRowValueByHeader_(header, row, "チェーン判定");
  const facilityStatus = flowerGetRowValueByHeader_(header, row, "施設判定");
  const areaStatus = flowerGetRowValueByHeader_(header, row, "エリア判定");
  const addressStatus = flowerGetRowValueByHeader_(header, row, "住所判定");
  const normalizedPhone = flowerGetRowValueByHeader_(header, row, "正規化電話番号");
  const normalizedGenre = flowerGetRowValueByHeader_(header, row, "正規化ジャンル");

  if (fetchStatus === "失敗" || (!storeName && url)) {
    return { group: FLOWER_SHEETS.failed, item: "詳細取得失敗", status: "未対応", nextAction: "再取得" };
  }
  if (dupStatus && dupStatus !== "ユニーク") return { group: FLOWER_SHEETS.exclude, item: "重複除外", status: "除外確定", nextAction: "投入しない" };
  if (chainStatus === "チェーン店") return { group: FLOWER_SHEETS.exclude, item: "チェーン店除外", status: "除外確定", nextAction: "投入しない" };
  if (facilityStatus === "除外") return { group: FLOWER_SHEETS.exclude, item: "大型小売・施設除外", status: "除外確定", nextAction: "投入しない" };
  if (areaStatus === "エリア外") return { group: FLOWER_SHEETS.exclude, item: "エリア外除外", status: "除外確定", nextAction: "投入しない" };
  if (!storeName || (!address && !url) || (addressStatus === "住所未取得" && !url)) {
    return { group: FLOWER_SHEETS.exclude, item: "住所未取得除外", status: "除外確定", nextAction: "投入しない" };
  }
  // チェーン店疑い（マスタ未登録だが地名付き支店名パターンで検出）は
  // 即除外にせず確認対象へ回し、人の目で最終判断してもらう。
  if (chainStatus === "チェーン店疑い") {
    return { group: FLOWER_SHEETS.confirm, item: "チェーン店疑い確認", status: "未対応", nextAction: "支店名パターンで検出。本当にチェーン店か目視確認してください" };
  }
  if (facilityStatus === "確認対象") return { group: FLOWER_SHEETS.confirm, item: "小規模ビル・市場確認", status: "未対応", nextAction: "テナントか路面店か確認" };
  if (!normalizedPhone && storeName && address) return { group: FLOWER_SHEETS.confirm, item: "電話番号なし確認", status: "未対応", nextAction: "電話番号補完" };
  if (!normalizedGenre || !flowerIsValidTargetGenre_(normalizedGenre)) return { group: FLOWER_SHEETS.confirm, item: "ジャンル確認", status: "未対応", nextAction: "ジャンルを目視確認" };
  if (addressStatus !== "住所あり") return { group: FLOWER_SHEETS.confirm, item: "住所確認", status: "未対応", nextAction: "住所確認" };
  if (areaStatus === "判定不可") return { group: FLOWER_SHEETS.confirm, item: "住所確認", status: "未対応", nextAction: "住所確認" };

  if (flowerIsComdeskTargetRow_(header, row)) return { group: FLOWER_SHEETS.target, item: "営業対象", status: "未対応", nextAction: "コムデスク投入" };
  return { group: FLOWER_SHEETS.confirm, item: "住所確認", status: "未対応", nextAction: "目視確認" };
}

// =====================================================================
// 7. タブ分け
// =====================================================================
function flowerExecuteSplitSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(FLOWER_SHEETS.facility);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = flowerNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const buckets = {};
  buckets[FLOWER_SHEETS.target] = [];
  buckets[FLOWER_SHEETS.confirm] = [];
  buckets[FLOWER_SHEETS.exclude] = [];
  buckets[FLOWER_SHEETS.failed] = [];

  rows.forEach(row => {
    const group = flowerGetRowValueByHeader_(header, row, "ワークフローグループ") || FLOWER_SHEETS.confirm;
    if (buckets[group]) buckets[group].push(row);
    else buckets[FLOWER_SHEETS.confirm].push(row);
  });

  Object.keys(buckets).forEach(sheetName => {
    flowerWriteRowsToSheetByName_(ss, sheetName, header, buckets[sheetName]);
  });

  return { count: rows.length };
}

// =====================================================================
// 8. 04_SALES地域別タブ生成
// =====================================================================
function flowerExecuteGenerateSalesAreaSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(FLOWER_SHEETS.target);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = flowerNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const areaContainers = {};

  rows.forEach(row => {
    if (!flowerIsComdeskTargetRow_(header, row)) return;
    const area = flowerGetAreaKeyForRow_(header, row);
    if (!areaContainers[area]) areaContainers[area] = [];
    areaContainers[area].push(flowerBuildComdeskRow_(header, row));
  });

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("花屋_04_SALES_") && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });

  Object.keys(areaContainers).forEach(area => {
    const sheetName = "花屋_04_SALES_" + flowerSafeSheetNamePart_(area);
    flowerWriteRowsToExistingSheet_(flowerGetOrCreateSheet_(ss, sheetName), FLOWER_COMDESK_HEADER, areaContainers[area]);
  });

  return { count: Object.keys(areaContainers).length };
}

function flowerGetAreaKeyForRow_(header, row) {
  const city = flowerGetRowValueByHeader_(header, row, "市区町村");
  if (city) return city;
  const pref = flowerGetRowValueByHeader_(header, row, "都道府県");
  if (pref) return pref;
  return "エリア不明";
}

// =====================================================================
// 9. CSV出力
// =====================================================================
function flowerExecuteExportSalesAreaCsvFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const exportFolder = flowerGetOrCreateFolder_(parentFolder, FLOWER_FOLDER_NAMES.export);
  const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
  let exported = 0;

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (!sheetName.startsWith("花屋_04_SALES_")) return;
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return;

    const areaName = sheetName.replace("花屋_04_SALES_", "") || "ダウンロードリスト";
    const fileName = "【花屋営業リスト】" + areaName + "_" + formattedDate + ".csv";
    const existingFiles = exportFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

    const blob = Utilities.newBlob("﻿" + flowerConvertArrayToCsvText_(values), "text/csv", fileName);
    exportFolder.createFile(blob);
    exported++;
  });

  ss.toast("CSV出力完了: " + exported + "ファイル", "📂 CSV出力");
  return { exported };
}

// =====================================================================
// 10. 件数サマリー
// =====================================================================
function flowerExecuteCountSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const countRows = sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    return sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;
  };

  const totalNormalized = countRows(FLOWER_SHEETS.normalized);
  const totalTarget = countRows(FLOWER_SHEETS.target);
  const totalConfirm = countRows(FLOWER_SHEETS.confirm);
  const totalExclude = countRows(FLOWER_SHEETS.exclude);
  const totalFailed = countRows(FLOWER_SHEETS.failed);

  const areaRows = [];
  let totalSales = 0;
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith("花屋_04_SALES_")) return;
    const count = Math.max(sheet.getLastRow() - 1, 0);
    if (count <= 0) return;
    areaRows.push([name, count]);
    totalSales += count;
  });
  areaRows.sort((a, b) => b[1] - a[1]);

  const output = [
    ["区分", "件数"],
    ["── 全体 ──", ""],
    ["花屋_01_NORMALIZED（取込総数）", totalNormalized],
    [FLOWER_SHEETS.target, totalTarget],
    [FLOWER_SHEETS.confirm, totalConfirm],
    [FLOWER_SHEETS.exclude, totalExclude],
    [FLOWER_SHEETS.failed, totalFailed],
    ["", ""],
    ["── 04_SALES_地域別 ──", ""]
  ].concat(areaRows.length ? areaRows : [["（地域別タブなし）", 0]])
    .concat([
      ["", ""],
      ["地域別合計（01_営業対象と一致するはず）", totalSales],
      ["更新日時", Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss")]
    ]);

  const summarySheet = flowerGetOrCreateSheet_(ss, FLOWER_SHEETS.summary);
  flowerWriteRowsToExistingSheet_(summarySheet, output[0], output.slice(1));
  summarySheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  summarySheet.autoResizeColumns(1, 2);
  ss.setActiveSheet(summarySheet);
  ss.moveActiveSheet(1);

  return { totalNormalized, totalTarget, totalConfirm, totalExclude, totalFailed, totalSales };
}

// =====================================================================
// 判定・正規化ヘルパー
// =====================================================================
function flowerNormalizeSystemGenre_(genre, searchGenre, sourceGenre, storeName) {
  const rawGenre = flowerTextValue_(genre);
  const searchGenreText = flowerTextValue_(searchGenre);
  const sourceGenreText = flowerTextValue_(sourceGenre);
  const nameText = flowerTextValue_(storeName);

  if (flowerIsValidTargetGenre_(rawGenre)) return rawGenre;

  const candidates = [searchGenreText, rawGenre, sourceGenreText];

  for (const candidate of candidates) {
    if (candidate && FLOWER_GENRE_MAP[candidate]) return FLOWER_GENRE_MAP[candidate];
  }

  for (const candidate of candidates) {
    if (!candidate) continue;
    const matched = Object.keys(FLOWER_GENRE_MAP).find(oldGenre => candidate.indexOf(oldGenre) !== -1);
    if (matched) return FLOWER_GENRE_MAP[matched];
  }

  if (nameText) {
    const nameKana = flowerHiraganaToKatakana_(nameText);
    const matchedTarget = FLOWER_TARGET_GENRES.find(g => nameKana.indexOf(flowerHiraganaToKatakana_(g)) !== -1);
    if (matchedTarget) return matchedTarget;

    const matchedOld = Object.keys(FLOWER_GENRE_MAP).find(oldGenre => nameKana.indexOf(flowerHiraganaToKatakana_(oldGenre)) !== -1);
    if (matchedOld) return FLOWER_GENRE_MAP[matchedOld];
  }

  return rawGenre;
}

function flowerJudgeAreaStatus_(address, pref, city) {
  const cleanAddress = flowerNormalizeAddressText_(address);
  const prefText = flowerTextValue_(pref);
  const cityText = flowerTextValue_(city);
  if (!cleanAddress) return { status: "判定不可", reason: "住所未取得" };
  if (!prefText || !cityText) return { status: "判定不可", reason: "都道府県または市区町村が空" };
  if (cleanAddress.indexOf(prefText) === -1 && cleanAddress.indexOf(cityText) === -1) {
    return { status: "判定不可", reason: "住所に都道府県/市区町村が含まれない" };
  }
  return { status: "エリア内", reason: "" };
}

function flowerJudgeAddressStatus_(address, pref, city) {
  const cleanAddress = flowerNormalizeAddressText_(address);
  if (!cleanAddress) return { status: "住所未取得", reason: "住所未取得" };
  if (!flowerTextValue_(pref) || !flowerTextValue_(city)) return { status: "住所未取得", reason: "都道府県または市区町村が空" };
  return { status: "住所あり", reason: "" };
}

function flowerIsComdeskTargetRow_(header, row) {
  const phone = flowerGetRowValueByHeader_(header, row, "正規化電話番号") || flowerNormalizePhoneNumberForAnalysis_(flowerGetRowValueByHeader_(header, row, "電話番号"));
  const genre = flowerGetRowValueByHeader_(header, row, "正規化ジャンル") || flowerGetRowValueByHeader_(header, row, "ジャンル");
  return flowerGetRowValueByHeader_(header, row, "営業対象判定") === "対象" &&
    flowerGetRowValueByHeader_(header, row, "重複判定") === "ユニーク" &&
    // "チェーン店"も"チェーン店疑い"も除外し、"単独店"のみ通す
    flowerGetRowValueByHeader_(header, row, "チェーン判定") === "単独店" &&
    flowerGetRowValueByHeader_(header, row, "施設判定") === "対象" &&
    flowerGetRowValueByHeader_(header, row, "エリア判定") === "エリア内" &&
    !!phone &&
    !!flowerGetRowValueByHeader_(header, row, "住所") &&
    flowerIsValidTargetGenre_(genre);
}

function flowerIsValidTargetGenre_(genre) {
  return FLOWER_TARGET_GENRES.indexOf(flowerTextValue_(genre)) !== -1;
}

// 支店名らしさを判定する（アパレル版と同じ3段階ロジック）。
// 返り値: "" (支店名パターンではない) / "high" (号店・駅前店・モール系
// テナント名など、ほぼ確実にチェーンの支店と分かる強いシグナル) /
// "heuristic" (地名＋店等、チェーンの可能性が高いが確証がないシグナル)。
function flowerIsLikelyBranchStoreName_(storeName) {
  const name = flowerTextValue_(storeName).normalize("NFKC").trim();
  if (!name) return "";

  // ①単独店・個人店にありがちな「〜店」は保護（誤除外防止）
  const personalLikePatterns = [
    /花店$/, /生花店$/, /園芸店$/, /植木店$/, /苗木店$/, /専門店$/, /フラワーショップ$/, /ブティック$/,
    /商店$/, /本店$/, /販売店$/
  ];
  if (personalLikePatterns.some(re => re.test(name))) return "";

  // ②明らかな商業施設・モールのテナントを示す接尾辞（強いシグナル）
  if (/(モール店|イオン店|アウトレット店|SC店|ショッピングセンター店|スクエア店|パーク店)$/.test(name)) return "high";

  // ③号店・駅・インター等、チェーン店の支店名として非常に一般的な接尾辞（強いシグナル）
  if (/[0-9０-９]+号店$/.test(name)) return "high";
  if (/(駅前|駅|北口|南口|東口|西口|インター|バイパス).{0,10}店$/.test(name)) return "high";

  // ④スペースや括弧で区切られた「〇〇店」（例:"店舗名 春日部店","店舗名(古河店)"）
  if (/[\s\(（][^\s\(（]+店[\)）]?$/.test(name)) return "heuristic";

  // ⑤地名（都道府県・市区町村）＋「店」（例:"〇〇下妻市店","〇〇日立市店"）
  if (/.+(都|道|府|県|市|区|町|村).{0,10}店$/.test(name)) return "heuristic";

  // ⑥行政区画文字を含まない短い地名・固有名詞＋「店」（例:"下妻店","古河店"）
  if (/^.{1,6}店$/.test(name)) return "heuristic";

  return "";
}

function flowerStripGenreSuffixFromName_(name) {
  const raw = flowerTextValue_(name);
  const match = raw.match(/^(.*?)\s*[（(]([^）)]{1,20})[）)]\s*$/);
  if (!match) return raw;
  const inner = flowerTextValue_(match[2]);
  const knownGenreWords = FLOWER_TARGET_GENRES.concat(Object.keys(FLOWER_GENRE_MAP));
  if (knownGenreWords.indexOf(inner) !== -1) return flowerTextValue_(match[1]);
  return raw;
}

function flowerNormalizeHpStatus_(hp, url) {
  const value = flowerTextValue_(hp).toLowerCase();
  if (value === "1" || value === "true" || value.indexOf("有") !== -1) return "1";
  if (url) return "1";
  return "0";
}

// =====================================================================
// コムデスク行生成
// =====================================================================
function flowerBuildComdeskRow_(header, row) {
  const storeName = flowerGetRowValueByHeader_(header, row, "店名");
  const fullAddr = flowerGetRowValueByHeader_(header, row, "住所");
  const addrDetails = flowerParseAddressDetails_(fullAddr);
  const pref = flowerGetRowValueByHeader_(header, row, "都道府県") || addrDetails.pref;
  const city = flowerGetRowValueByHeader_(header, row, "市区町村");
  const phone = flowerGetRowValueByHeader_(header, row, "電話番号");
  const cleanPhone = flowerNormalizePhoneNumberForAnalysis_(phone).replace(/[^\d]/g, "");
  const media = flowerGetRowValueByHeader_(header, row, "媒体") || "Googleマップ";
  const url = flowerGetRowValueByHeader_(header, row, "URL");
  const hpHave = flowerGetRowValueByHeader_(header, row, "HP有無");
  const hpStatus = (hpHave.indexOf("有") !== -1 || hpHave === "1" || hpHave.toLowerCase() === "true") ? "1" : "0";
  const bizDaysVal = flowerRemoveHolidayFromBizDays_(flowerGetRowValueByHeader_(header, row, "営業日"), flowerGetRowValueByHeader_(header, row, "定休日"));
  const holidayVal = flowerGetRowValueByHeader_(header, row, "定休日");
  const rawOpenA = flowerGetRowValueByHeader_(header, row, "営業開始A") || flowerGetRowValueByHeader_(header, row, "営業開始");
  const openAVal = flowerFormatToPureTime_(flowerToHalfWidthForTime_(rawOpenA));
  const closeAVal = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetRowValueByHeader_(header, row, "営業終了A") || flowerGetRowValueByHeader_(header, row, "営業終了")));
  const openBVal = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetRowValueByHeader_(header, row, "営業開始B")));
  const closeBVal = flowerFormatToPureTime_(flowerToHalfWidthForTime_(flowerGetRowValueByHeader_(header, row, "営業終了B")));
  const timeValues = flowerNormalizeBusinessTimeValues_(rawOpenA, openAVal, closeAVal, openBVal, closeBVal);

  const cleanAddr1 = city ? addrDetails.addr1.replace(city, "") : addrDetails.addr1;
  const address1 = city ? city + cleanAddr1 : addrDetails.addr1;
  const areaText = pref + city;

  const salesRow = Array(31).fill("");
  salesRow[2] = storeName;
  salesRow[4] = addrDetails.pcode;
  salesRow[5] = pref;
  salesRow[6] = address1;
  salesRow[9] = phone;
  salesRow[14] = url;
  salesRow[17] = media;
  salesRow[21] = hpStatus;
  salesRow[22] = areaText + "tel" + cleanPhone;
  salesRow[25] = bizDaysVal;
  salesRow[26] = holidayVal;
  salesRow[27] = timeValues.openA;
  salesRow[28] = timeValues.closeA;
  salesRow[29] = timeValues.openB;
  salesRow[30] = timeValues.closeB;
  return salesRow;
}

// =====================================================================
// 汎用ヘルパー
// =====================================================================
function flowerGetValue_(header, row, names) {
  for (const name of names) {
    const idx = header.indexOf(name);
    if (idx !== -1 && row[idx] !== undefined && row[idx] !== null && flowerTextValue_(row[idx]) !== "") {
      return flowerTextValue_(row[idx]);
    }
  }
  return "";
}

function flowerConvertArrayToCsvText_(array) {
  return array.map(row => row.map(cell => {
    const str = String(cell === null || cell === undefined ? "" : cell).replace(/"/g, '""');
    if (str.includes(",") || str.includes("\n") || str.includes("\r") || str.includes('"')) return '"' + str + '"';
    return str;
  }).join(",")).join("\r\n");
}

function flowerSimplifyStoreName_(name) {
  if (!name) return "";
  let n = String(name).normalize("NFKC").toLowerCase();
  n = n.replace(/[ぁ-ゖ]/g, m => String.fromCharCode(m.charCodeAt(0) + 0x60));
  n = n.replace(/[\s ・、。，．・！？!?()（）【】\[\]「」『』_－\-〜~]/g, "");
  n = n.replace(/(店|駅前店|北口店|南口店|東口店|西口店|インター店|本店|支店|営業所)$/, "");
  return n.trim();
}

function flowerNormalizeForKeywordMatch_(text) {
  return flowerTextValue_(text)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s　・．.、,＆&()（）【】\[\]「」『』_－\-ー]/g, "");
}

function flowerHiraganaToKatakana_(text) {
  return flowerTextValue_(text).replace(/[ぁ-ゖ]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
}

function flowerGetOrCreateFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function flowerGetOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function flowerNormalizeHeaderRow_(header) {
  return header.map(h => String(h || "").replace(/^﻿/, "").trim());
}

function flowerGetRowValueByHeader_(header, row, name) {
  const idx = header.indexOf(name);
  return idx === -1 ? "" : flowerTextValue_(row[idx]);
}

function flowerWriteRowsToExistingSheet_(sheet, header, rows) {
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clearContents();
  sheet.clearFormats();

  const output = [header].concat(rows || []);
  const maxCols = Math.max.apply(null, output.map(r => r.length));
  const normalized = output.map(r => {
    const row = r.slice();
    while (row.length < maxCols) row.push("");
    return row;
  });

  const range = sheet.getRange(1, 1, normalized.length, maxCols);
  range.setNumberFormat("@");
  range.setValues(normalized);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, maxCols);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();
  }
}

function flowerWriteRowsToSheetByName_(ss, sheetName, header, rows) {
  const sheet = flowerGetOrCreateSheet_(ss, sheetName);
  flowerWriteRowsToExistingSheet_(sheet, header, rows);
}

function flowerNormalizePhoneDisplay_(phone) {
  return flowerTextValue_(phone)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[ー−―]/g, "-")
    .replace(/[^0-9\-]/g, "")
    .trim();
}

function flowerNormalizePhoneNumberForAnalysis_(phone) {
  const digits = flowerTextValue_(phone)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[^\d]/g, "");
  if (!digits) return "";
  if (digits.length === 11) return digits.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3");
  if (digits.startsWith("03") || digits.startsWith("06")) return digits.replace(/(\d{2})(\d{4})(\d{4})/, "$1-$2-$3");
  if (digits.length === 10) return digits.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
  return digits;
}

function flowerNormalizeAddressText_(address) {
  return flowerTextValue_(address)
    .replace(/(?:〒\d{3}-?\d{4}\s*|日本、\s*|日本\s*)/g, "")
    .replace(/\s+/g, "");
}

function flowerParsePrefCityFromAddress_(address) {
  const cleanAddress = flowerNormalizeAddressText_(address);
  const match = cleanAddress.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))?((?:.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村]))?/);
  return {
    pref: match && match[1] ? match[1] : "",
    city: match && match[2] ? match[2] : ""
  };
}

function flowerParseAddressDetails_(fullAddress) {
  let pcode = "";
  let pref = "";
  let addr1 = flowerTextValue_(fullAddress);
  const pcodeMatch = addr1.match(/〒?(\d{3})-?(\d{4})/);
  if (pcodeMatch) {
    pcode = pcodeMatch[1] + "-" + pcodeMatch[2];
    addr1 = addr1.replace(pcodeMatch[0], "");
  }
  addr1 = addr1.replace(/^日本、\s*/, "").replace(/^日本\s*/, "");
  const prefMatch = addr1.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))/);
  if (prefMatch) {
    pref = prefMatch[1];
    addr1 = addr1.replace(pref, "");
  }
  return { pcode, pref, addr1: addr1.replace(/\s+/g, "") };
}

function flowerToHalfWidthForTime_(str) {
  return flowerTextValue_(str).replace(/[０-９：]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
}

function flowerNormalizeBusinessTimeValues_(rawOpenA, openAVal, closeAVal, openBVal, closeBVal) {
  const result = { openA: openAVal, closeA: closeAVal, openB: openBVal, closeB: closeBVal };
  if (!openBVal && !closeBVal) {
    const timeMatches = flowerToHalfWidthForTime_(rawOpenA).match(/(\d{1,2}:\d{2})/g);
    if (timeMatches && timeMatches.length >= 2) {
      result.openA = flowerFormatToPureTime_(timeMatches[0]);
      result.closeA = "";
      result.openB = "";
      result.closeB = flowerFormatToPureTime_(timeMatches[timeMatches.length - 1]);
    } else if (openAVal) {
      result.closeA = "";
      result.openB = "";
      result.closeB = closeAVal;
    }
  }
  return result;
}

function flowerRemoveHolidayFromBizDays_(bizDays, holiday) {
  let bizDaysVal = flowerTextValue_(bizDays);
  const holidayVal = flowerTextValue_(holiday);
  if (!bizDaysVal || !holidayVal) return bizDaysVal;
  ["月", "火", "水", "木", "金", "土", "日", "祝"].forEach(day => {
    if (holidayVal.indexOf(day) !== -1) {
      const regex = new RegExp(day + "[・、/]?|[・、/]?" + day, "g");
      bizDaysVal = bizDaysVal.replace(regex, "");
    }
  });
  return bizDaysVal.replace(/^[・、/]+|[・、/]+$/g, "").replace(/[・、/]{2,}/g, "・");
}

function flowerFormatToPureTime_(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    const hh = Utilities.formatDate(val, "JST", "HH");
    const mm = Utilities.formatDate(val, "JST", "mm");
    const hourInt = parseInt(hh, 10);
    return (mm === "00" || mm === "") ? String(hourInt) : hourInt + ":" + mm;
  }
  const str = String(val).trim();
  if (str === "") return "";
  const timeMatch = str.match(/(\d{1,2}):(\d{2})/);
  if (timeMatch) {
    const hourInt = parseInt(timeMatch[1], 10);
    const minStr = timeMatch[2];
    return minStr === "00" ? String(hourInt) : hourInt + ":" + minStr;
  }
  const numMatch = str.match(/^(\d{1,2})$/);
  return numMatch ? String(parseInt(numMatch[1], 10)) : str;
}

function flowerSafeSheetNamePart_(name) {
  return flowerTextValue_(name)
    .replace(/[\\\/\?\*\[\]\:]/g, "")
    .replace(/_/g, "")
    .substring(0, 30) || "エリア不明";
}

function flowerUniqueTextList_(values) {
  const seen = {};
  return values.map(flowerTextValue_).filter(value => {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function flowerTextValue_(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}
