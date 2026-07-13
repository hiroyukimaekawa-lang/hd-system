/**
 * 楽天トラベル宿泊施設専用GAS Ver1.3.0
 *
 * 目的：
 * Octoparseで取得した楽天トラベル宿泊施設リストを、
 * CSV投入フォルダ方式で自動取り込みし、
 * 詳細URL作成・詳細データ結合・個人店寄り分類・エリア分け・CSV出力まで行う。
 *
 * ★Ver1.3.0の変更点（Ver1.2.0からの差分）：
 * 1. 部屋数（客室数）による営業対象判定を追加。
 *    「電気切り替え条件で宿泊施設に無料HP作成」という営業は、自前で立派なHPや
 *    マーケティング体制をすでに持っている大型ホテルよりも、そうした体制を
 *    持っていない可能性が高い小規模施設（民宿・ペンション・ゲストハウス等）に
 *    向いている。部屋数が一定数（既定15室）以上の施設は、既存のチェーン系
 *    ホテル・大型旅館である可能性が高いため、チェーンキーワードに一致しなくても
 *    部屋数だけで除外できるようにした。
 *    Octoparse側の部屋数列の実際のヘッダー名がまだ確定していないため、
 *    RAKUTEN_ROOM_COUNT_COLUMNS によく使われそうな候補名を複数登録し、
 *    実際の列名が分かり次第そこに追加するだけで反映されるようにしてある
 *    （店名・住所等、他の項目と同じ「候補名リストから探す」方式に統一）。
 * 2. 宿泊系チェーンの除外キーワードを大幅に拡充。
 *    従来はアパホテル・東横INN・ルートイン等、主要チェーンの一部しか
 *    登録されていなかったため、他の全国チェーン（ホテルマイステイズ、
 *    ヴィアイン、チサンホテル、JRホテルグループ、ワシントンホテル等）が
 *    未除外のまま営業対象に混ざるリスクがあった。飲食店側のGAS
 *    （MASTER_CHAIN・fixKnownChainMasterGaps）で行っているのと同じ考え方で、
 *    既知の主要チェーンをまとめて追加した。
 * 3. 04_SALES系タブの備考欄に部屋数を表示するようにし、架電担当が
 *    リストを見ただけで規模感を把握できるようにした。
 *
 * Ver1.2.0の変更点（Ver1.1.0以前からの差分）：
 * Octoparseのタスクを1本にまとめて、一覧の時点で「住所・TEL・FAX」まで
 * 同じ行に取得できるようになったため、従来の
 *   1. 一覧CSV取り込み → 2. 詳細URL一覧作成 → 3. Octoparseタスク2で詳細CSV取得 → 4. 結合
 * という二段階方式に加えて、
 *   1. 一覧CSV（住所・TEL・FAX込み）を取り込むだけで正規化まで進む
 * という「一体型CSV」方式に自動対応した。
 * 「楽天_raw」シートに フィールド3(住所)・フィールド5(電話番号) 列があり、
 * かつ実際にデータが入っていれば、詳細CSVがなくても自動的に一体型として処理する。
 * 「楽天_詳細抽出」シートにデータが入っている場合は、従来通り二段階結合を行う
 * （両対応・自動判定なので、Octoparseタスクの組み方はどちらでもよい）。
 *
 * 通常運用（一体型CSV / Octoparseタスク1本のみの場合）：
 * 1. 「0. 初期タブ・投入フォルダを作成」を実行
 * 2. Octoparseで取得した「住所・TEL・FAX込み」のCSVを「楽天_raw_CSV投入フォルダ」に入れる
 * 3. 「🚀 全自動処理：投入済みCSVからCSV出力まで実行」を実行
 *    → 詳細CSVが無くても、一覧CSVに住所・電話番号が入っていればそのまま最後まで処理される
 *
 * 通常運用（二段階方式 / Octoparseタスク2本に分ける場合。従来通り）：
 * 1. 「0. 初期タブ・投入フォルダを作成」を実行
 * 2. Octoparseタスク1の一覧CSVを「楽天_raw_CSV投入フォルダ」に入れる
 * 3. 「1. 一覧CSV取り込み・詳細URL一覧作成」を実行
 * 4. 「楽天_詳細URL一覧」の「基本情報URL_std」をOctoparseタスク2に入れる
 * 5. Octoparseタスク2の詳細CSVを「楽天_詳細_CSV投入フォルダ」に入れる
 * 6. 「🚀 全自動処理：投入済みCSVからCSV出力まで実行」を実行
 */

// =====================================================================
// 0. メニュー
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏕 楽天トラベル宿泊施設")
    .addItem("0. 初期タブ・投入フォルダを作成", "rakutenSetupAll")
    .addSeparator()
    .addItem("1. 一覧CSV取り込み・詳細URL一覧作成", "rakutenImportRawAndCreateDetailUrls")
    .addItem("2. 詳細CSVだけ取り込み", "rakutenImportDetailOnly")
    .addSeparator()
    .addItem("3. 楽天 詳細URL一覧を作成", "rakutenCreateDetailUrlList")
    .addItem("4. 楽天 一覧＋詳細を結合・正規化", "rakutenNormalizeAndMerge")
    .addItem("5. 楽天 宿泊施設を分類・地域別タブ作成", "rakutenSplitClassifySheets")
    .addItem("6. 楽天 システム投入CSVタブ作成", "rakutenCreateSalesSheets")
    .addItem("7. 楽天 CSVをDrive出力", "rakutenExportSalesCsvFiles")
    .addSeparator()
    .addItem("🚀 全自動処理：投入済みCSVからCSV出力まで実行", "rakutenRunAllFromFolders")
    .addToUi();
}

// =====================================================================
// 1. 設定
// =====================================================================
const RAKUTEN_SHEETS = {
  raw: "楽天_raw",
  detailUrl: "楽天_詳細URL一覧",
  detail: "楽天_詳細抽出",
  normalized: "楽天_正規化",
  target: "楽天_営業対象",
  confirm: "楽天_要確認",
  exclude: "楽天_除外",
  duplicate: "楽天_重複"
};

const RAKUTEN_FOLDER_NAMES = {
  rawInput: "楽天_raw_CSV投入フォルダ",
  rawProcessed: "楽天_raw_処理済みフォルダ",
  detailInput: "楽天_詳細_CSV投入フォルダ",
  detailProcessed: "楽天_詳細_処理済みフォルダ",
  export: "完成版CSVエクスポート"
};

// Octoparseタスク1（一覧）の実際の出力列名。
// 「フィールド1_テキスト_テキスト」＝施設名、「フィールド1_リンク_リンク」＝詳細ページURL。
// ※一体型CSV（住所・TEL・FAX込み）の場合は「フィールド1_テキスト」「フィールド1_リンク」（末尾の重複なし）になることもあるため、
//   両方のパターンを候補に入れている。
const RAKUTEN_RAW_NAME_COLUMNS = ["フィールド1_テキスト_テキスト", "フィールド1_テキスト", "jsraleventscroll", "タイトル", "施設名", "店名", "名前"];
const RAKUTEN_RAW_URL_COLUMNS = ["フィールド1_リンク_リンク", "フィールド1_リンク", "jsraleventscroll_URL", "タイトルURL", "フィールド", "URL", "楽天URL"];

// ★Ver1.3.0 NEW: 部屋数（客室数）列の候補名。
// Octoparse側の実際のヘッダー名がまだ確定していないため、店名・住所等と同じ
// 「候補名リストの中から最初に見つかったものを使う」方式にしてある。
// 実際の列名が分かったら、この配列の先頭に追加するだけで反映される。
const RAKUTEN_ROOM_COUNT_COLUMNS = [
  "部屋数",
  "客室数",
  "総部屋数",
  "総客室数",
  "室数",
  "部屋数(室)",
  "客室数(室)",
  "Room",
  "Rooms",
  "room",
  "rooms"
];

// ★Ver1.3.0 NEW: 部屋数がこの数値以上なら「規模が大きく、自前でHP・マーケティング体制を
// すでに持っている可能性が高い」とみなして営業対象から除外する。
// 「電気切り替え条件でHP無料作成」の営業接点は、そうした体制をまだ持っていない
// 小規模施設（民宿・ペンション・ゲストハウス等）ほど刺さりやすいという想定。
const RAKUTEN_ROOM_COUNT_EXCLUDE_THRESHOLD = 15;

const RAKUTEN_NORMALIZED_HEADER = [
  "店名",
  "宿泊ジャンル",
  "都道府県",
  "市区町村",
  "大エリア",
  "郵便番号",
  "住所",
  "電話番号",
  "FAX",
  "URL",
  "媒体",
  "HP有無",
  "営業日",
  "定休日",
  "営業開始A",
  "営業終了A",
  "営業開始B",
  "営業終了B",
  "施設説明",
  "アクセス",
  "口コミ",
  "料金",
  "設備タグ",
  "画像URL",
  "駐車場",
  "宿泊施設判定",
  "判定理由",
  "重複判定",
  "元一覧URL",
  "基本情報URL_std",
  "詳細取得方式",
  "部屋数" // ★Ver1.3.0 NEW: 既存列との位置ズレ事故を避けるため末尾に追加
];

const RAKUTEN_COMDESK_HEADER = [
  "UUID", "種別", "名前", "カナ", "郵便番号", "都道府県", "住所１", "住所２", "住所カナ",
  "Tel1", "Tel2", "Tel3", "Tel4", "FAX", "URL", "備考", "旧社名", "リードソース",
  "旧進捗", "履歴", "オーナー名", "HPある？", "BP検索", "アポ済商材", "最新履歴",
  "営業曜日", "休業曜日", "午前始", "午前終", "午後始", "午後終"
];

const RAKUTEN_TARGET_KEYWORDS = [
  "民宿",
  "ペンション",
  "ゲストハウス",
  "グランピング",
  "キャンプ",
  "コテージ",
  "ロッジ",
  "山荘",
  "古民家",
  "農家民宿",
  "小さな宿",
  "海辺の宿",
  "料理宿",
  "ログハウス",
  "宿坊"
];

const RAKUTEN_CONFIRM_KEYWORDS = [
  "貸別荘",
  "一棟貸し",
  "一棟",
  "ヴィラ",
  "villa",
  "旅館",
  "温泉",
  "inn",
  "別荘",
  "コンドミニアム",
  "ホテル"
];

// ★Ver1.3.0で拡充: 全国チェーン系ホテル・旅館ブランドを追加。
// 「電気切り替え条件でHP無料作成」の営業対象としては、すでに自社サイト・
// 予約システム・マーケティング体制を持っているチェーン系は不向きなため、
// 部屋数フィルタと二重で弾けるようにしておく（部屋数が未取得の場合の保険にもなる）。
const RAKUTEN_EXCLUDE_KEYWORDS = [
  "rakuten stay",
  "楽天stay",
  "ｒａｋｕｔｅｎ stay",
  "ｒａｋｕｔｅｎ　ｓｔａｙ",
  "亀の井ホテル",
  "アパホテル",
  "東横inn",
  "東横イン",
  "ルートイン",
  "route inn",
  "スーパーホテル",
  "ドーミーイン",
  "dormy inn",
  "リブマックス",
  "ヒルトン",
  "マリオット",
  "プリンスホテル",
  "星野リゾート",
  "大江戸温泉物語",
  "グランドホテル",
  "グランドタワー",
  "ホテル＆レジデンス",
  "ホテル&レジデンス",
  "リゾートホテル",
  "ビジネスホテル",
  "カプセルホテル",
  "コンフォートホテル",
  "コンフォートイン",
  "ダイワロイネット",
  // ↓ここからVer1.3.0で追加した全国チェーン
  "ホテルマイステイズ",
  "mystays",
  "変なホテル",
  "ヴィアイン",
  "via inn",
  "チサンホテル",
  "センチュリオンホテル",
  "ネストホテル",
  "ホテルウィングインターナショナル",
  "ホテルグレイスリー",
  "ホテルメッツ",
  "ホテルグランヴィア",
  "相鉄フレッサイン",
  "相鉄グランドフレッサ",
  "r&bホテル",
  "ホテル法華クラブ",
  "法華クラブ",
  "ホテルサンルート",
  "サンルートホテル",
  "ソラリア西鉄ホテル",
  "西鉄ホテル",
  "ニューオータニ",
  "ホテルオークラ",
  "リーガロイヤルホテル",
  "全日空ホテル",
  "anaクラウンプラザ",
  "ana crowne plaza",
  "ワシントンホテル",
  "藤田観光",
  "スマイルホテル",
  "smile hotel",
  "hotel az",
  "ホテルaz",
  "グリーンリッチホテル",
  "リッチモンドホテル",
  "richmond hotel",
  "jrイン",
  "jr inn",
  "東急ステイ",
  "東急ホテルズ",
  "東急ホテル",
  "相鉄ホテルズ",
  "ホテル日航",
  "ベッセルホテル",
  "ホテルwbf",
  "ヴィラフォンテーヌ",
  "アークホテル",
  "ロワジールホテル",
  "京王プレッソイン",
  "京王プラザホテル",
  "第一ホテル",
  "ホテルレオパレス",
  "アルモニーアンブラッセ",
  // ↓茨城県宿泊.csv（実データ）で確認した「〇〇ホテルグループ」表記のチェーン。
  // 個別ブランド名を覚えるより、「ホテルグループ」という自己申告的な表記自体を
  // 汎用の除外条件にした方が同種の中小チェーンを広く拾えるため追加。
  "ホテルグループ",
  "bbhホテルグループ"
];

// =====================================================================
// 2. 初期設定
// =====================================================================
function rakutenSetupAll() {
  rakutenSetupSheets();
  rakutenCreateInputFolders_();

  SpreadsheetApp.getUi().alert(
    "初期設定が完了しました。\n\n" +
    "スプレッドシートと同じDriveフォルダ内に、以下のフォルダを作成しました。\n\n" +
    "・楽天_raw_CSV投入フォルダ\n" +
    "・楽天_raw_処理済みフォルダ\n" +
    "・楽天_詳細_CSV投入フォルダ\n" +
    "・楽天_詳細_処理済みフォルダ\n" +
    "・完成版CSVエクスポート"
  );
}

function rakutenSetupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const initialSheets = [
    RAKUTEN_SHEETS.raw,
    RAKUTEN_SHEETS.detailUrl,
    RAKUTEN_SHEETS.detail,
    RAKUTEN_SHEETS.normalized,
    RAKUTEN_SHEETS.target,
    RAKUTEN_SHEETS.confirm,
    RAKUTEN_SHEETS.exclude,
    RAKUTEN_SHEETS.duplicate
  ];

  initialSheets.forEach(name => rktGetOrCreateSheet_(ss, name));

  const detailSheet = ss.getSheetByName(RAKUTEN_SHEETS.detail);
  if (detailSheet.getLastRow() === 0) {
    detailSheet.getRange(1, 1, 1, 11).setValues([[
      "施設名",
      "住所ラベル",
      "住所",
      "TELラベル",
      "電話番号",
      "FAXラベル",
      "FAX",
      "交通アクセスラベル",
      "交通アクセス",
      "駐車場ラベル",
      "駐車場"
    ]]);
  }
}

function rakutenCreateInputFolders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.rawInput);
  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.rawProcessed);
  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.detailInput);
  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.detailProcessed);
  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.export);
}

// =====================================================================
// 3. フォルダ投入型：一覧CSV取り込み・詳細URL一覧作成
// =====================================================================
function rakutenImportRawAndCreateDetailUrls() {
  rakutenCreateInputFolders_();

  const imported = rakutenImportRawCsvFiles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAKUTEN_SHEETS.raw);
  const rawValues = rawSheet ? rawSheet.getDataRange().getValues() : [];
  const rawCount = Math.max(rawValues.length - 1, 0);

  if (rawCount <= 0) {
    SpreadsheetApp.getUi().alert(
      "一覧CSVの取り込みができませんでした。\n\n" +
      "楽天_raw_CSV投入フォルダにOctoparseの一覧CSVを入れてから再実行してください。"
    );
    return;
  }

  const rawHeader = rktNormalizeHeader_(rawValues[0]);
  const rawRows = rawValues.slice(1);

  // ★一体型CSV（住所・TEL・FAXが一覧の時点で取得済み）の場合は、
  //   詳細URL一覧の作成・Octoparseタスク2は不要なのでスキップする。
  if (rktRawHasEmbeddedDetail_(rawHeader, rawRows)) {
    SpreadsheetApp.getUi().alert(
      "一覧CSVの取り込みが完了しました。\n\n" +
      `取り込みファイル数: ${imported.files}件\n` +
      `一覧データ件数: ${rawCount}件\n\n` +
      "このCSVには住所・電話番号（フィールド3・フィールド5）がすでに含まれている「一体型CSV」と判定しました。\n" +
      "詳細URL一覧の作成やOctoparseタスク2は不要です。\n\n" +
      "続けて「🚀 全自動処理」または「4. 楽天 一覧＋詳細を結合・正規化」を実行してください。"
    );
    return;
  }

  rakutenCreateDetailUrlList(true);

  SpreadsheetApp.getUi().alert(
    "一覧CSVの取り込みと詳細URL一覧の作成が完了しました。\n\n" +
    `取り込みファイル数: ${imported.files}件\n` +
    `一覧データ件数: ${rawCount}件\n\n` +
    "次に「楽天_詳細URL一覧」タブの「基本情報URL_std」列をOctoparseタスク2に入れてください。"
  );
}

function rakutenImportDetailOnly() {
  rakutenCreateInputFolders_();

  const imported = rakutenImportDetailCsvFiles();

  if (!imported || imported.rows === 0) {
    SpreadsheetApp.getUi().alert(
      "詳細CSVの取り込みができませんでした。\n\n" +
      "楽天_詳細_CSV投入フォルダにOctoparseタスク2のCSVを入れてから再実行してください。"
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    "詳細CSVの取り込みが完了しました。\n\n" +
    `取り込みファイル数: ${imported.files}件\n` +
    `詳細データ件数: ${imported.rows}件\n\n` +
    "次に「🚀 全自動処理」を実行してください。"
  );
}

function rakutenImportRawCsvFiles() {
  return rakutenImportCsvFilesToSheet_(
    RAKUTEN_FOLDER_NAMES.rawInput,
    RAKUTEN_FOLDER_NAMES.rawProcessed,
    RAKUTEN_SHEETS.raw
  );
}

function rakutenImportDetailCsvFiles() {
  return rakutenImportCsvFilesToSheet_(
    RAKUTEN_FOLDER_NAMES.detailInput,
    RAKUTEN_FOLDER_NAMES.detailProcessed,
    RAKUTEN_SHEETS.detail
  );
}

function rakutenImportCsvFilesToSheet_(inputFolderName, processedFolderName, targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  const inputFolder = rktGetOrCreateFolder_(parentFolder, inputFolderName);
  const processedFolder = rktGetOrCreateFolder_(parentFolder, processedFolderName);
  const targetSheet = rktGetOrCreateSheet_(ss, targetSheetName);

  const files = inputFolder.getFiles();
  const parsedFiles = [];
  const unifiedHeader = [];
  const headerSet = new Set();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (!fileName.toLowerCase().endsWith(".csv")) {
      continue;
    }

    try {
      const blob = file.getBlob();
      let csvText = blob.getDataAsString("UTF-8").replace(/^﻿/, "");

      if (csvText.indexOf("�") !== -1) {
        csvText = blob.getDataAsString("MS932").replace(/^﻿/, "");
      }

      const parsed = Utilities.parseCsv(csvText);

      if (!parsed || parsed.length === 0) {
        file.moveTo(processedFolder);
        continue;
      }

      const header = parsed[0].map(h => String(h || "").replace(/^﻿/, "").trim());
      const rows = parsed.slice(1).filter(row => row.join("").trim() !== "");

      header.forEach(col => {
        if (col && !headerSet.has(col)) {
          headerSet.add(col);
          unifiedHeader.push(col);
        }
      });

      parsedFiles.push({
        fileName,
        header,
        rows
      });

      file.moveTo(processedFolder);

    } catch (e) {
      Logger.log(`[楽天CSV取込エラー] ${fileName}: ${e.message}`);
      SpreadsheetApp.getUi().alert(`CSV取込エラー: ${fileName}\n${e.message}`);
    }
  }

  if (parsedFiles.length === 0) {
    return {
      files: 0,
      rows: 0
    };
  }

  const output = [unifiedHeader];

  parsedFiles.forEach(pf => {
    const colMap = pf.header.map(col => unifiedHeader.indexOf(col));

    pf.rows.forEach(row => {
      const aligned = new Array(unifiedHeader.length).fill("");

      row.forEach((value, i) => {
        const targetIdx = colMap[i];
        if (targetIdx !== -1 && targetIdx !== undefined) {
          aligned[targetIdx] = value;
        }
      });

      output.push(aligned);
    });
  });

  rktWrite_(targetSheet, output);

  return {
    files: parsedFiles.length,
    rows: output.length - 1
  };
}

// =====================================================================
// 4. 楽天_raw → 楽天_詳細URL一覧（二段階方式用。一体型CSVでは不要）
// =====================================================================
function rakutenCreateDetailUrlList(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAKUTEN_SHEETS.raw);

  if (!rawSheet) {
    if (!silent) SpreadsheetApp.getUi().alert(`「${RAKUTEN_SHEETS.raw}」シートがありません。`);
    return { count: 0 };
  }

  const values = rawSheet.getDataRange().getValues();

  if (values.length <= 1) {
    if (!silent) SpreadsheetApp.getUi().alert(`「${RAKUTEN_SHEETS.raw}」にデータがありません。`);
    return { count: 0 };
  }

  const header = rktNormalizeHeader_(values[0]);
  const rows = values.slice(1);

  const output = [["一覧行番号", "施設名", "一覧URL", "基本情報URL_std"]];
  const seen = new Set();

  rows.forEach((row, i) => {
    const name = rktGetValue_(header, row, RAKUTEN_RAW_NAME_COLUMNS);
    const listUrl = rktGetValue_(header, row, RAKUTEN_RAW_URL_COLUMNS);
    const stdUrl = rktConvertToStdUrl_(listUrl);

    if (!stdUrl || seen.has(stdUrl)) return;

    seen.add(stdUrl);
    output.push([i + 2, name, listUrl, stdUrl]);
  });

  const sheet = rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.detailUrl);
  rktWrite_(sheet, output);

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      `楽天詳細URL一覧を作成しました。\n作成件数: ${output.length - 1}件\n\n` +
      `「${RAKUTEN_SHEETS.detailUrl}」の「基本情報URL_std」列をOctoparseタスク2に貼り付けてください。`
    );
  }

  return { count: output.length - 1 };
}

// =====================================================================
// 5. 全自動処理
// =====================================================================
function rakutenRunAllFromFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  rakutenCreateInputFolders_();

  rakutenImportRawCsvFiles();
  rakutenImportDetailCsvFiles();

  const rawSheet = ss.getSheetByName(RAKUTEN_SHEETS.raw);
  const detailSheet = ss.getSheetByName(RAKUTEN_SHEETS.detail);

  const rawValues = rawSheet ? rawSheet.getDataRange().getValues() : [];
  const rawCount = Math.max(rawValues.length - 1, 0);
  const detailCount = detailSheet ? Math.max(detailSheet.getLastRow() - 1, 0) : 0;

  if (rawCount <= 0) {
    SpreadsheetApp.getUi().alert(
      "一覧データがありません。\n\n" +
      "楽天_raw_CSV投入フォルダにOctoparseの一覧CSVを入れてから再実行してください。"
    );
    return;
  }

  const rawHeader = rktNormalizeHeader_(rawValues[0]);
  const rawRows = rawValues.slice(1);
  const hasEmbeddedDetail = rktRawHasEmbeddedDetail_(rawHeader, rawRows);

  // 詳細CSVもなく、一覧側にも住所・電話番号が埋め込まれていない場合のみ、
  // 「詳細CSVがまだ足りない」ケースとして案内して止める。
  if (detailCount <= 0 && !hasEmbeddedDetail) {
    rakutenCreateDetailUrlList(true);

    SpreadsheetApp.getUi().alert(
      "一覧CSVの取り込みは完了しましたが、詳細データ（住所・電話番号）がまだありません。\n\n" +
      "・二段階方式の場合：「楽天_詳細URL一覧」タブの「基本情報URL_std」列をOctoparseタスク2に入れて、\n" +
      "  詳細CSVを「楽天_詳細_CSV投入フォルダ」に入れてから、もう一度「🚀 全自動処理」を実行してください。\n" +
      "・一体型CSV（Octoparse1本で住所・電話番号まで取得する）場合は、\n" +
      "  一覧CSVに「フィールド3」（住所）「フィールド5」（電話番号）の列と値が含まれているか確認してください。"
    );
    return;
  }

  if (detailCount > 0) {
    // 従来の二段階方式のときだけ詳細URL一覧を作っておく（任意・確認用）
    rakutenCreateDetailUrlList(true);
  }

  rakutenNormalizeAndMerge(true);
  rakutenSplitClassifySheets(true);
  rakutenCreateSalesSheets(true);
  rakutenExportSalesCsvFiles(true);
  rakutenCreateSalesRegionSummary(true);

  const targetSheet = ss.getSheetByName(RAKUTEN_SHEETS.target);
  const confirmSheet = ss.getSheetByName(RAKUTEN_SHEETS.confirm);
  const excludeSheet = ss.getSheetByName(RAKUTEN_SHEETS.exclude);
  const duplicateSheet = ss.getSheetByName(RAKUTEN_SHEETS.duplicate);

  const targetCount = targetSheet ? Math.max(targetSheet.getLastRow() - 1, 0) : 0;
  const confirmCount = confirmSheet ? Math.max(confirmSheet.getLastRow() - 1, 0) : 0;
  const excludeCount = excludeSheet ? Math.max(excludeSheet.getLastRow() - 1, 0) : 0;
  const duplicateCount = duplicateSheet ? Math.max(duplicateSheet.getLastRow() - 1, 0) : 0;

  SpreadsheetApp.getUi().alert(
    "楽天宿泊施設リストの全自動処理が完了しました。\n\n" +
    `処理方式: ${hasEmbeddedDetail && detailCount <= 0 ? "一体型CSV（一覧に住所・電話番号込み）" : "二段階結合（一覧＋詳細CSV）"}\n` +
    `一覧データ: ${rawCount}件\n` +
    `詳細データ: ${detailCount}件\n\n` +
    `営業対象: ${targetCount}件\n` +
    `要確認: ${confirmCount}件\n` +
    `除外: ${excludeCount}件\n` +
    `重複: ${duplicateCount}件\n\n` +
    "04_SALES_宿泊施設 / 市区町村別タブ / 完成版CSVエクスポートを確認してください。"
  );
}

function rakutenRunAfterDetailExtraction() {
  rakutenNormalizeAndMerge();
  rakutenSplitClassifySheets();
  rakutenCreateSalesSheets();
  rakutenExportSalesCsvFiles();
}

// =====================================================================
// 6. 楽天_raw (+ 楽天_詳細抽出) → 楽天_正規化
//    ・楽天_詳細抽出にデータがあれば、従来通りURL/施設名で突き合わせる「二段階結合」。
//    ・楽天_詳細抽出が無くても、楽天_rawに住所・電話番号列（フィールド3・フィールド5）が
//      埋まっていれば、突き合わせ不要の「一体型結合」を自動で行う。
// =====================================================================
function rakutenNormalizeAndMerge(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAKUTEN_SHEETS.raw);

  if (!rawSheet) {
    if (!silent) SpreadsheetApp.getUi().alert(`「${RAKUTEN_SHEETS.raw}」シートがありません。`);
    return { count: 0 };
  }

  const rawValues = rawSheet.getDataRange().getValues();

  if (rawValues.length <= 1) {
    if (!silent) SpreadsheetApp.getUi().alert(`「${RAKUTEN_SHEETS.raw}」にデータがありません。`);
    return { count: 0 };
  }

  const rawHeader = rktNormalizeHeader_(rawValues[0]);
  const rawRows = rawValues.slice(1);

  const detailSheet = ss.getSheetByName(RAKUTEN_SHEETS.detail);
  const detailValues = detailSheet ? detailSheet.getDataRange().getValues() : [];

  let outputRows;
  let modeLabel;

  if (detailValues.length > 1) {
    // ---- 二段階結合（従来方式）----
    const detailHeader = rktNormalizeHeader_(detailValues[0]);
    const detailRows = detailValues.slice(1);
    outputRows = rktBuildNormalizedRowsTwoStage_(rawHeader, rawRows, detailHeader, detailRows);
    modeLabel = "二段階結合";
  } else if (rktRawHasEmbeddedDetail_(rawHeader, rawRows)) {
    // ---- 一体型結合（新方式：一覧CSVに住所・電話番号込み）----
    outputRows = rktBuildNormalizedRowsCombined_(rawHeader, rawRows);
    modeLabel = "一体型CSV結合";
  } else {
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        "詳細データ（住所・電話番号）が見つかりません。\n\n" +
        "・二段階方式の場合は「楽天_詳細_CSV投入フォルダ」に詳細CSVを入れてください。\n" +
        "・一体型CSV（Octoparse1本で住所・電話番号まで取得したCSV）の場合は、\n" +
        "  「フィールド3」（住所）・「フィールド5」（電話番号）列が含まれているか確認してください。"
      );
    }
    return { count: 0 };
  }

  const sheet = rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.normalized);
  rktWrite_(sheet, [RAKUTEN_NORMALIZED_HEADER].concat(outputRows));

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      `楽天一覧＋詳細データの結合・正規化が完了しました（${modeLabel}）。\n出力件数: ${outputRows.length}件`
    );
  }

  return { count: outputRows.length };
}

// ---- 二段階結合（従来方式のロジックをそのまま関数化）----
function rktBuildNormalizedRowsTwoStage_(rawHeader, rawRows, detailHeader, detailRows) {
  const detailIndex = rktBuildDetailIndex_(detailHeader, detailRows);
  const outputRows = [];

  rawRows.forEach((rawRow, i) => {
    const listName = rktGetValue_(rawHeader, rawRow, RAKUTEN_RAW_NAME_COLUMNS);
    const listUrl = rktGetValue_(rawHeader, rawRow, RAKUTEN_RAW_URL_COLUMNS);
    const stdUrl = rktConvertToStdUrl_(listUrl);

    const detail = rktFindDetail_(detailIndex, stdUrl, listName, i);
    const detailData = detail ? rktNormalizeDetailRow_(detailHeader, detail.row) : {};

    const name = detailData.name || listName;
    const address = rktCleanAddress_(detailData.address);
    const parsed = rktParseAddress_(address);

    const description = rktGetValue_(rawHeader, rawRow, ["htlspecial", "施設説明", "説明"]);
    const listAccess = rktGetValue_(rawHeader, rawRow, ["htlaccess", "アクセス"]);
    const access = detailData.access || listAccess;

    // 都道府県＋大エリア列（例:「千葉県 銚子・旭・九十九里・東金・茂原」）があれば優先的に使う
    const prefAreaRaw = rktGetValue_(rawHeader, rawRow, ["都道府県"]);
    const prefArea = rktParsePrefArea_(prefAreaRaw);
    const area = prefArea.area || rktJudgeLargeArea_(name, address, access, description);

    const review = rktGetValue_(rawHeader, rawRow, ["cstmrevl", "口コミ", "レビュー"]);
    const price = rktGetValue_(rawHeader, rawRow, ["価格20", "価格", "料金"]);
    const imageUrl = rktGetValue_(rawHeader, rawRow, ["学歴", "画像URL", "画像URL1"]);

    const tags = [
      rktGetValue_(rawHeader, rawRow, ["内容16"]),
      rktGetValue_(rawHeader, rawRow, ["内容17"]),
      rktGetValue_(rawHeader, rawRow, ["内容18"]),
      rktGetValue_(rawHeader, rawRow, ["内容19"]),
      rktGetValue_(rawHeader, rawRow, ["タグ15"])
    ].filter(Boolean).join(" / ");

    const genre = rktJudgeLodgingGenre_(name, description, tags);
    const phoneDisplay = rktNormalizePhoneDisplay_(detailData.phone);
    const faxDisplay = rktIsDummyFax_(detailData.fax) ? "" : rktNormalizePhoneDisplay_(detailData.fax);

    // ★Ver1.3.0 NEW: 部屋数は一覧側・詳細側どちらに来ても拾えるよう両方見る
    // （一覧側の値を優先。無ければ詳細側）。
    const roomCountRaw = rktGetValue_(rawHeader, rawRow, RAKUTEN_ROOM_COUNT_COLUMNS)
      || rktGetValue_(detailHeader, detail ? detail.row : [], RAKUTEN_ROOM_COUNT_COLUMNS);
    const roomCount = rktParseRoomCount_(roomCountRaw);

    let salesJudge = rktJudgeRakutenSalesTarget_(name, genre, description, tags, address, roomCount);

    const missingReasons = [];
    if (!phoneDisplay) missingReasons.push("電話番号なし");
    if (!address) missingReasons.push("住所なし");

    if (salesJudge.status !== "除外" && missingReasons.length > 0) {
      salesJudge = {
        status: "要確認",
        reason: salesJudge.reason + " / " + missingReasons.join(" / ")
      };
    }

    outputRows.push([
      name,
      genre,
      parsed.pref || prefArea.pref,
      parsed.city,
      area,
      parsed.zip,
      address,
      phoneDisplay,
      faxDisplay,
      stdUrl || listUrl,
      "楽天トラベル",
      "0",
      "",
      "",
      "",
      "",
      "",
      "",
      description,
      access,
      review,
      price,
      tags,
      imageUrl,
      detailData.parking || "",
      salesJudge.status,
      salesJudge.reason,
      "",
      listUrl,
      stdUrl,
      detail ? detail.method : "未結合",
      roomCount === null ? "" : String(roomCount) // ★Ver1.3.0 NEW
    ]);
  });

  return outputRows;
}

// ---- 一体型結合（新方式：突き合わせ不要。1行の中に住所・電話番号が全部入っている）----
function rktBuildNormalizedRowsCombined_(rawHeader, rawRows) {
  const outputRows = [];

  rawRows.forEach(rawRow => {
    const name = rktGetValue_(rawHeader, rawRow, RAKUTEN_RAW_NAME_COLUMNS);
    const listUrl = rktGetValue_(rawHeader, rawRow, RAKUTEN_RAW_URL_COLUMNS);
    const stdUrl = rktConvertToStdUrl_(listUrl);

    // 都道府県＋大エリア（例:「千葉県 銚子・旭・九十九里・東金・茂原」）を分割
    const prefAreaRaw = rktGetValue_(rawHeader, rawRow, ["都道府県"]);
    const prefArea = rktParsePrefArea_(prefAreaRaw);

    // 住所：Octoparseの「フィールドN」はラベル/値が交互に並ぶ形式で、実データ確認の結果
    // フィールド2="住所"(ラベル)→フィールド3=値、という並びだった（茨城県宿泊.csvで確認）。
    // ただし列番号はOctoparseのタスク設定によりズレる可能性があるため、
    // まずラベル文字列で探し（rktFindFieldByLabel_）、見つからない場合のみ
    // 従来の固定列番号候補にフォールバックする。
    const addressRaw = rktFindFieldByLabel_(rawHeader, rawRow, ["住所"])
      || rktGetValue_(rawHeader, rawRow, ["フィールド3", "住所"]);
    const address = rktCleanAddress_(addressRaw);
    const parsedAddr = rktParseAddress_(address);

    const pref = parsedAddr.pref || prefArea.pref;
    const city = parsedAddr.city;
    const area = prefArea.area || rktJudgeLargeArea_(name, address, "", "");

    // 電話番号：フィールド4="TEL"(ラベル)→フィールド5=値、という並び（実データで確認）
    const phoneRaw = rktFindFieldByLabel_(rawHeader, rawRow, ["TEL", "電話番号"])
      || rktGetValue_(rawHeader, rawRow, ["フィールド5", "電話番号", "TEL"]);
    const phone = rktNormalizePhoneDisplay_(phoneRaw);

    // FAX：実データ（茨城県宿泊.csv）にはFAXのラベル/値ペアが存在せず、
    // フィールド6/7は「総部屋数」（後述の部屋数）に使われていることが判明したため、
    // フィールド7を決め打ちでFAXとして読むのをやめ、ラベル文字列一致でのみ拾うようにした。
    // FAXラベルが存在しないCSVでは空欄になる（＝実態通り）。
    const faxRaw = rktFindFieldByLabel_(rawHeader, rawRow, ["FAX"]);
    const fax = rktIsDummyFax_(faxRaw) ? "" : rktNormalizePhoneDisplay_(faxRaw);

    const description = rktGetValue_(rawHeader, rawRow, ["hotelcharacter", "施設説明", "説明"]);
    // 価格：実データでは「価格」列自体は "[最安料金（目安）]" という見出し文字列が入っており、
    // 実際の金額は価格1（例:"3,273円～"）・価格2（税込表記）に入っていたため、
    // 価格1・価格2を優先するよう修正（茨城県宿泊.csvで確認）。
    const price = rktGetValue_(rawHeader, rawRow, ["価格1", "価格2", "incldtax", "plnprc", "料金"]);
    const review = rktGetValue_(rawHeader, rawRow, ["hotelrating", "cstmrevl", "口コミ"]);
    const imageUrl = rktGetValue_(rawHeader, rawRow, ["画像URL", "画像URL1"]);
    // タグ：実データにはmoreplan3やタグ15に相当する列がなく、代わりに"planoutline"
    // （プラン説明文）に「グランピング」等のジャンルを示す語が入ることがあったため、
    // ジャンル・営業対象判定用のテキストにのみ含める（表示用タグ列としては汚いので出さない）
    const planOutlineRaw = rktGetValue_(rawHeader, rawRow, ["planoutline"]);
    const tags = rktGetValue_(rawHeader, rawRow, ["moreplan3", "タグ15"]);
    const judgeText = tags + " " + rktNormalizeText_(planOutlineRaw).slice(0, 200);

    const genre = rktJudgeLodgingGenre_(name, description, judgeText);

    // ★Ver1.3.0: 部屋数（客室数）。実データではフィールド6="総部屋数"(ラベル)→
    // フィールド7=値（例:"65室"）という並びだったため、まずラベル一致で探し、
    // 見つからなければ従来の候補名リスト（RAKUTEN_ROOM_COUNT_COLUMNS）にフォールバックする。
    const roomCountRaw = rktFindFieldByLabel_(rawHeader, rawRow, ["総部屋数", "部屋数", "客室数", "室数"])
      || rktGetValue_(rawHeader, rawRow, RAKUTEN_ROOM_COUNT_COLUMNS);
    const roomCount = rktParseRoomCount_(roomCountRaw);

    let salesJudge = rktJudgeRakutenSalesTarget_(name, genre, description, judgeText, address, roomCount);

    const missingReasons = [];
    if (!phone) missingReasons.push("電話番号なし");
    if (!address) missingReasons.push("住所なし");

    if (salesJudge.status !== "除外" && missingReasons.length > 0) {
      salesJudge = {
        status: "要確認",
        reason: salesJudge.reason + " / " + missingReasons.join(" / ")
      };
    }

    outputRows.push([
      name,
      genre,
      pref,
      city,
      area,
      parsedAddr.zip,
      address,
      phone,
      fax,
      stdUrl || listUrl,
      "楽天トラベル",
      "0",
      "",
      "",
      "",
      "",
      "",
      "",
      description,
      "",
      review,
      price,
      tags,
      imageUrl,
      "",
      salesJudge.status,
      salesJudge.reason,
      "",
      listUrl,
      stdUrl,
      "一体型CSV結合",
      roomCount === null ? "" : String(roomCount) // ★Ver1.3.0 NEW
    ]);
  });

  return outputRows;
}

// =====================================================================
// 7. 楽天_正規化 → 営業対象/要確認/除外/重複/エリア別
// =====================================================================
function rakutenSplitClassifySheets(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RAKUTEN_SHEETS.normalized);

  if (!sheet) {
    if (!silent) SpreadsheetApp.getUi().alert(`「${RAKUTEN_SHEETS.normalized}」シートがありません。`);
    return { target: 0, confirm: 0, exclude: 0, duplicate: 0 };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { target: 0, confirm: 0, exclude: 0, duplicate: 0 };
  }

  const header = rktNormalizeHeader_(values[0]);
  const rows = values.slice(1);

  const targetRows = [];
  const confirmRows = [];
  const excludeRows = [];
  const duplicateRows = [];
  const areaMap = {};

  const seenPhones = new Set();
  const seenNameAddress = new Set();

  rows.forEach(row => {
    const name = rktGetValue_(header, row, ["店名"]);
    const address = rktGetValue_(header, row, ["住所"]);
    const phone = rktNormalizePhoneDigits_(rktGetValue_(header, row, ["電話番号"]));
    const city = rktGetValue_(header, row, ["市区町村"]) || rktGetValue_(header, row, ["大エリア"]) || "エリア不明";
    const currentStatus = rktGetValue_(header, row, ["宿泊施設判定"]);

    let duplicateReason = "";

    if (phone && seenPhones.has(phone)) {
      duplicateReason = "重複（電話番号一致）";
    } else if (phone) {
      seenPhones.add(phone);
    }

    const nameAddressKey = rktSimplifyName_(name) + "|" + rktNormalizeText_(address);

    if (!duplicateReason && nameAddressKey !== "|" && seenNameAddress.has(nameAddressKey)) {
      duplicateReason = "重複（施設名＋住所一致）";
    } else if (!duplicateReason && nameAddressKey !== "|") {
      seenNameAddress.add(nameAddressKey);
    }

    const newRow = row.slice();
    const duplicateIdx = header.indexOf("重複判定");

    if (duplicateIdx !== -1) {
      newRow[duplicateIdx] = duplicateReason || "ユニーク";
    }

    if (!areaMap[city]) areaMap[city] = [];
    areaMap[city].push(newRow);

    if (duplicateReason) {
      duplicateRows.push(newRow);
      return;
    }

    if (currentStatus === "営業対象") {
      targetRows.push(newRow);
    } else if (currentStatus === "除外") {
      excludeRows.push(newRow);
    } else {
      confirmRows.push(newRow);
    }
  });

  rktWrite_(rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.target), [header].concat(targetRows));
  rktWrite_(rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.confirm), [header].concat(confirmRows));
  rktWrite_(rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.exclude), [header].concat(excludeRows));
  rktWrite_(rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.duplicate), [header].concat(duplicateRows));

  // 前回実行分の楽天_エリア_〇〇タブを削除してから作り直す（使わなくなった地域タブが残らないように）
  ss.getSheets().forEach(s => {
    if (s.getName().indexOf("楽天_エリア_") === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(s);
    }
  });

  Object.keys(areaMap).forEach(city => {
    const safeCity = rktSafeSheetName_(city);
    rktWrite_(rktGetOrCreateSheet_(ss, "楽天_エリア_" + safeCity), [header].concat(areaMap[city]));
  });

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "楽天宿泊施設の分類が完了しました。\n\n" +
      `営業対象: ${targetRows.length}件\n` +
      `要確認: ${confirmRows.length}件\n` +
      `除外: ${excludeRows.length}件\n` +
      `重複: ${duplicateRows.length}件`
    );
  }

  return {
    target: targetRows.length,
    confirm: confirmRows.length,
    exclude: excludeRows.length,
    duplicate: duplicateRows.length
  };
}

// =====================================================================
// 8. 楽天_営業対象 → 04_SALES_宿泊施設 / 市区町村別
// =====================================================================
function rakutenCreateSalesSheets(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(RAKUTEN_SHEETS.target);

  if (!targetSheet) {
    if (!silent) SpreadsheetApp.getUi().alert(`「${RAKUTEN_SHEETS.target}」シートがありません。`);
    return { count: 0 };
  }

  const values = targetSheet.getDataRange().getValues();

  if (values.length <= 1) {
    if (!silent) SpreadsheetApp.getUi().alert("楽天_営業対象にデータがありません。");
    return { count: 0 };
  }

  const header = rktNormalizeHeader_(values[0]);
  const rows = values.slice(1);

  const allRows = [];
  const cityMap = {};

  rows.forEach(row => {
    const salesRow = rktBuildComdeskRow_(header, row);
    allRows.push(salesRow);

    const city = rktGetValue_(header, row, ["市区町村"]) || rktGetValue_(header, row, ["大エリア"]) || "エリア不明";

    if (!cityMap[city]) cityMap[city] = [];
    cityMap[city].push(salesRow);
  });

  ss.getSheets().forEach(s => {
    const name = s.getName();
    if (name === "04_SALES_宿泊施設" || name.indexOf("04_SALES_宿泊_") === 0) {
      if (ss.getSheets().length > 1) ss.deleteSheet(s);
    }
  });

  rktWrite_(rktGetOrCreateSheet_(ss, "04_SALES_宿泊施設"), [RAKUTEN_COMDESK_HEADER].concat(allRows));

  Object.keys(cityMap).forEach(city => {
    const safeCity = rktSafeSheetName_(city);
    rktWrite_(
      rktGetOrCreateSheet_(ss, "04_SALES_宿泊_" + safeCity),
      [RAKUTEN_COMDESK_HEADER].concat(cityMap[city])
    );
  });

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      `システム投入用CSVタブを作成しました。\n04_SALES_宿泊施設: ${allRows.length}件`
    );
  }

  return { count: allRows.length };
}

// =====================================================================
// 9. 04_SALES_宿泊系 → Drive CSV出力
// =====================================================================
function rakutenExportSalesCsvFiles(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();
  const exportFolder = rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.export);
  const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");

  let exported = 0;

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();

    if (!(sheetName === "04_SALES_宿泊施設" || sheetName.indexOf("04_SALES_宿泊_") === 0)) return;

    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) return;

    const areaName = sheetName.replace("04_SALES_", "");
    const fileName = `【営業リスト】${areaName}_${formattedDate}.csv`;

    const bom = "﻿";
    const csv = rktArrayToCsv_(values);
    const blob = Utilities.newBlob(bom + csv, "text/csv", fileName);

    const existing = exportFolder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    exportFolder.createFile(blob);
    exported++;
  });

  if (!silent) {
    SpreadsheetApp.getUi().alert(`楽天宿泊施設CSVをDrive出力しました。\n出力ファイル数: ${exported}件`);
  }

  return { exported };
}

// =====================================================================
// 10. 詳細データ結合系ヘルパー（二段階結合方式で使用）
// =====================================================================
function rktBuildDetailIndex_(detailHeader, detailRows) {
  const byUrl = {};
  const byName = {};
  const byOrder = [];

  detailRows.forEach((row, i) => {
    const detail = rktNormalizeDetailRow_(detailHeader, row);
    const stdUrl = rktConvertToStdUrl_(detail.url);
    const cleanName = rktSimplifyName_(detail.name);

    if (stdUrl) {
      byUrl[stdUrl] = { row, index: i, method: "URL結合" };
    }

    if (cleanName) {
      byName[cleanName] = { row, index: i, method: "施設名結合" };
    }

    byOrder.push({ row, index: i, method: "行順結合" });
  });

  return { byUrl, byName, byOrder };
}

function rktFindDetail_(detailIndex, stdUrl, name, rawIndex) {
  if (stdUrl && detailIndex.byUrl[stdUrl]) {
    return detailIndex.byUrl[stdUrl];
  }

  const cleanName = rktSimplifyName_(name);

  if (cleanName && detailIndex.byName[cleanName]) {
    return detailIndex.byName[cleanName];
  }

  return detailIndex.byOrder[rawIndex] || null;
}

function rktNormalizeDetailRow_(header, row) {
  const block = rktGetValue_(header, row, ["基本情報ブロック", "テキスト", "Text", "text"]);
  const parsedFromBlock = rktParseBasicInfoBlock_(block);

  const name = rktGetValue_(header, row, ["施設名", "店名", "ホテル名", "フィールド1", "フィールド 1", "タイトル"]);
  const address = rktGetValue_(header, row, ["住所", "フィールド3", "フィールド 3"]) || parsedFromBlock.address;
  const phone = rktGetValue_(header, row, ["電話番号", "TEL", "Tel", "tel", "フィールド5", "フィールド 5"]) || parsedFromBlock.phone;
  const fax = rktGetValue_(header, row, ["FAX", "Fax", "fax", "フィールド7", "フィールド 7"]) || parsedFromBlock.fax;
  const access = rktGetValue_(header, row, ["交通アクセス", "アクセス", "フィールド9", "フィールド 9"]) || parsedFromBlock.access;
  const parking = rktGetValue_(header, row, ["駐車場", "フィールド11", "フィールド 11"]) || parsedFromBlock.parking;
  const url = rktGetValue_(header, row, ["楽天URL", "URL", "WebページURL", "ページURL", "Page URL", "ページのURL"]);

  return { name, address, phone, fax, access, parking, url };
}

function rktParseBasicInfoBlock_(block) {
  const text = rktText_(block).replace(/\r?\n/g, " ").replace(/\s+/g, " ");
  const result = { address: "", phone: "", fax: "", access: "", parking: "" };

  if (!text) return result;

  const addressMatch = text.match(/住所\s*(.*?)(?:\s*(?:TEL|電話番号|FAX|交通アクセス|駐車場)\s*)/i);
  if (addressMatch) result.address = addressMatch[1];

  const phoneMatch = text.match(/(?:TEL|電話番号)\s*([0-9０-９\-ー−―]{8,})/i);
  if (phoneMatch) result.phone = phoneMatch[1];

  const faxMatch = text.match(/FAX\s*([0-9０-９\-ー−―]{8,})/i);
  if (faxMatch) result.fax = faxMatch[1];

  const accessMatch = text.match(/交通アクセス\s*(.*?)(?:\s*駐車場\s*|$)/);
  if (accessMatch) result.access = accessMatch[1];

  const parkingMatch = text.match(/駐車場\s*(.*)$/);
  if (parkingMatch) result.parking = parkingMatch[1];

  return result;
}

// =====================================================================
// 11. 判定系ヘルパー
// =====================================================================
function rktJudgeLodgingGenre_(name, description, tags) {
  const text = rktNormalizeText_([name, description, tags].join(" "));

  if (text.indexOf("グランピング") !== -1) return "グランピング";
  if (text.indexOf("民宿") !== -1) return "民宿";
  if (text.indexOf("ペンション") !== -1) return "ペンション";
  if (text.indexOf("ゲストハウス") !== -1) return "ゲストハウス";
  if (text.indexOf("コテージ") !== -1) return "コテージ";
  if (text.indexOf("ロッジ") !== -1) return "ロッジ";
  if (text.indexOf("貸別荘") !== -1) return "貸別荘";
  if (text.indexOf("一棟貸し") !== -1 || text.indexOf("一棟") !== -1) return "一棟貸し";
  if (text.indexOf("ヴィラ") !== -1 || text.indexOf("villa") !== -1) return "ヴィラ";
  if (text.indexOf("旅館") !== -1) return "旅館";
  if (text.indexOf("ホテル") !== -1) return "ホテル";

  return "その他宿泊施設";
}

// ★Ver1.3.0: roomCount引数を追加。
// 判定の優先順位: ①チェーン除外キーワード → ②部屋数しきい値 → ③営業対象キーワード
// → ④要確認キーワード → ⑤判定不能（要確認）。
// 部屋数を②に置いたのは、「民宿」等の営業対象キーワードに一致していても、
// 実際には部屋数が多い＝すでに規模の大きい施設（チェーンの一部・大型旅館等）
// である可能性が高いため、キーワード一致より部屋数を優先させたいという考え方。
function rktJudgeRakutenSalesTarget_(name, genre, description, tags, address, roomCount) {
  const text = rktNormalizeText_([name, genre, description, tags, address].join(" "));

  const matchedExclude = RAKUTEN_EXCLUDE_KEYWORDS.find(k => text.indexOf(rktNormalizeText_(k)) !== -1);

  if (matchedExclude) {
    return { status: "除外", reason: "除外キーワード一致: " + matchedExclude };
  }

  if (typeof roomCount === "number" && roomCount !== null && !isNaN(roomCount) && roomCount >= RAKUTEN_ROOM_COUNT_EXCLUDE_THRESHOLD) {
    return {
      status: "除外",
      reason: `部屋数${roomCount}室のため除外（${RAKUTEN_ROOM_COUNT_EXCLUDE_THRESHOLD}室以上は規模大と判定）`
    };
  }

  const matchedTarget = RAKUTEN_TARGET_KEYWORDS.find(k => text.indexOf(rktNormalizeText_(k)) !== -1);

  if (matchedTarget) {
    return { status: "営業対象", reason: "営業対象キーワード一致: " + matchedTarget };
  }

  const matchedConfirm = RAKUTEN_CONFIRM_KEYWORDS.find(k => text.indexOf(rktNormalizeText_(k)) !== -1);

  if (matchedConfirm) {
    return { status: "要確認", reason: "要確認キーワード一致: " + matchedConfirm };
  }

  return { status: "要確認", reason: "判定キーワードなし" };
}

function rktJudgeLargeArea_(name, address, access, description) {
  const text = [name, address, access, description].join(" ");
  const parsed = rktParseAddress_(address);

  if (parsed.city) return parsed.city;

  if (text.indexOf("鴨川") !== -1) return "鴨川市";
  if (text.indexOf("勝浦") !== -1) return "勝浦市";
  if (text.indexOf("御宿") !== -1) return "御宿町";
  if (text.indexOf("大多喜") !== -1 || text.indexOf("養老渓谷") !== -1) return "大多喜町";
  if (text.indexOf("いすみ") !== -1) return "いすみ市";
  if (text.indexOf("南房総") !== -1) return "南房総市";
  if (text.indexOf("館山") !== -1) return "館山市";

  return "その他・要確認";
}

// 「都道府県」列に入っている「県名＋大エリア名」（例：「千葉県 銚子・旭・九十九里・東金・茂原 」）を
// 都道府県と大エリア（楽天公式のエリアくくり）に分割する。
function rktParsePrefArea_(text) {
  const t = rktText_(text).replace(/[\s　]+/g, " ").trim();

  if (!t) return { pref: "", area: "" };

  const match = t.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))\s*(.*)$/);

  if (match) {
    return { pref: match[1], area: match[2].trim() };
  }

  return { pref: "", area: t };
}

// ★Ver1.3.0 NEW: 「15室」「全15室」「客室数：20」「20部屋」等の文字列から
// 数字部分だけを取り出す。数字が取れなければnullを返す（＝部屋数不明。
// 部屋数フィルタは適用せず、他の判定基準に委ねる）。
function rktParseRoomCount_(value) {
  const text = rktText_(value);
  if (!text) return null;

  const halfWidth = text.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  const match = halfWidth.match(/(\d+)/);

  if (!match) return null;

  const n = parseInt(match[1], 10);
  return Number.isFinite(n) ? n : null;
}

// =====================================================================
// 12. システム投入CSV生成ヘルパー
// =====================================================================
function rktBuildComdeskRow_(header, row) {
  const name = rktGetValue_(header, row, ["店名"]);
  const zip = rktGetValue_(header, row, ["郵便番号"]);
  const pref = rktGetValue_(header, row, ["都道府県"]);
  const city = rktGetValue_(header, row, ["市区町村"]);
  const address = rktGetValue_(header, row, ["住所"]);
  const phone = rktGetValue_(header, row, ["電話番号"]);
  const fax = rktGetValue_(header, row, ["FAX"]);
  const url = rktGetValue_(header, row, ["URL"]);
  const description = rktGetValue_(header, row, ["施設説明"]);
  const access = rktGetValue_(header, row, ["アクセス"]);
  const parking = rktGetValue_(header, row, ["駐車場"]);
  const genre = rktGetValue_(header, row, ["宿泊ジャンル"]);
  const price = rktGetValue_(header, row, ["料金"]);
  const review = rktGetValue_(header, row, ["口コミ"]);
  const roomCount = rktGetValue_(header, row, ["部屋数"]); // ★Ver1.3.0 NEW
  const cleanPhone = rktNormalizePhoneDigits_(phone);

  const addr1 = rktBuildAddress1_(pref, city, address);
  const areaText = pref + city;
  const bp = areaText + "tel" + cleanPhone;

  const memoParts = [];

  if (genre) memoParts.push("宿泊ジャンル: " + genre);
  if (roomCount) memoParts.push("部屋数: " + roomCount + "室"); // ★Ver1.3.0 NEW
  if (description) memoParts.push("施設説明: " + description);
  if (access) memoParts.push("アクセス: " + access);
  if (parking) memoParts.push("駐車場: " + parking);
  if (price) memoParts.push("料金: " + price);
  if (review) memoParts.push("口コミ: " + review);

  const salesRow = Array(31).fill("");

  salesRow[0] = "";
  salesRow[1] = "";
  salesRow[2] = name;
  salesRow[3] = "";
  salesRow[4] = zip;
  salesRow[5] = pref;
  salesRow[6] = addr1;
  salesRow[7] = "";
  salesRow[8] = "";
  salesRow[9] = phone;
  salesRow[10] = "";
  salesRow[11] = "";
  salesRow[12] = "";
  salesRow[13] = fax;
  salesRow[14] = url;
  salesRow[15] = memoParts.join(" / ");
  salesRow[16] = "";
  salesRow[17] = "楽天トラベル";
  salesRow[18] = "";
  salesRow[19] = "";
  salesRow[20] = "";
  salesRow[21] = "0";
  salesRow[22] = bp;
  salesRow[23] = "";
  salesRow[24] = "";
  salesRow[25] = "";
  salesRow[26] = "";
  salesRow[27] = "";
  salesRow[28] = "";
  salesRow[29] = "";
  salesRow[30] = "";

  return salesRow;
}

function rktBuildAddress1_(pref, city, address) {
  let addr = rktCleanAddress_(address);

  if (pref) {
    addr = addr.replace(pref, "");
  }

  if (city && addr.indexOf(city) === 0) {
    return addr;
  }

  return city ? city + addr.replace(city, "") : addr;
}

// =====================================================================
// 13. 汎用ヘルパー
// =====================================================================
function rktConvertToStdUrl_(url) {
  let cleanUrl = rktText_(url).split("?")[0].replace(/\/$/, "");

  if (!cleanUrl) return "";

  const hotelMatch = cleanUrl.match(/\/HOTEL\/(\d+)\/(\d+)(?:_std)?\.html$/);

  if (hotelMatch) {
    return "https://travel.rakuten.co.jp/HOTEL/" + hotelMatch[1] + "/" + hotelMatch[2] + "_std.html";
  }

  const planMatch = cleanUrl.match(/hotelinfo\/(?:plan|hotel)\/(\d+)/);

  if (planMatch) {
    return "https://travel.rakuten.co.jp/HOTEL/" + planMatch[1] + "/" + planMatch[1] + "_std.html";
  }

  return "";
}

function rktParseAddress_(address) {
  let addr = rktCleanAddress_(address);
  let zip = "";
  let pref = "";
  let city = "";

  const zipMatch = addr.match(/〒?(\d{3}-\d{4})/);

  if (zipMatch) {
    zip = zipMatch[1];
    addr = addr.replace(/〒?\d{3}-\d{4}\s*/, "");
  }

  const prefMatch = addr.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))/);

  if (prefMatch) {
    pref = prefMatch[1];
    addr = addr.replace(pref, "");
  }

  const cityMatch = addr.match(/^(.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村])/);

  if (cityMatch) {
    city = cityMatch[1];
  }

  return { zip, pref, city, rest: addr };
}

function rktCleanAddress_(address) {
  return rktText_(address)
    .replace(/地図を見る/g, "")
    .replace(/^住所\s*/, "")
    .replace(/\s+/g, "")
    .trim();
}

function rktNormalizePhoneDisplay_(phone) {
  return rktText_(phone)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[ー−―]/g, "-")
    .replace(/[^0-9\-]/g, "")
    .trim();
}

function rktNormalizePhoneDigits_(phone) {
  return rktNormalizePhoneDisplay_(phone).replace(/[^\d]/g, "");
}

// FAXが「000-000-0000」「00-0000-0000」「-」「--」等のダミー値・空値かどうかを判定する。
// 楽天の基本情報欄はFAX未登録の施設でもこの手のダミー値が入っていることがあるため、
// 数字部分が空、または全て0のときは「FAXなし」として扱う。
function rktIsDummyFax_(fax) {
  const digits = rktNormalizePhoneDigits_(fax);

  if (!digits) return true;
  if (/^0+$/.test(digits)) return true;

  return false;
}

// 楽天_raw シートに、住所・電話番号がすでに埋め込まれた「一体型CSV」かどうかを判定する。
// フィールド3（住所）・フィールド5（電話番号）に相当する列があり、
// かつ実際に1件でも値が入っていれば「一体型」と判定する。
function rktRawHasEmbeddedDetail_(header, rows) {
  const hasAddressCol = header.indexOf("フィールド3") !== -1 || header.indexOf("住所") !== -1;
  const hasPhoneCol = header.indexOf("フィールド5") !== -1 ||
    header.indexOf("電話番号") !== -1 ||
    header.indexOf("TEL") !== -1;

  if (!hasAddressCol || !hasPhoneCol) return false;

  return rows.some(row => {
    const phone = rktFindFieldByLabel_(header, row, ["TEL", "電話番号"])
      || rktGetValue_(header, row, ["フィールド5", "電話番号", "TEL"]);
    const address = rktFindFieldByLabel_(header, row, ["住所"])
      || rktGetValue_(header, row, ["フィールド3", "住所"]);
    return !!(phone || address);
  });
}

// ★Ver1.3.0 NEW（茨城県宿泊.csvの実データで判明した構造への対応）：
// Octoparseの「フィールドN」列は、ラベルと値が交互に並ぶ形式で出力されることがある
// （例: フィールド2="住所"(ラベル文字列そのもの), フィールド3="〒300-2706茨城県..."(実際の値)、
//       フィールド6="総部屋数"(ラベル), フィールド7="65室"(実際の値)）。
// この並びの列番号はOctoparseのタスク設定やスクレイピング対象ページの構造によって
// ズレる可能性があるため、決め打ちの列番号（フィールド3・フィールド5・フィールド7など）
// に頼るのではなく、「ラベル列の中身が指定したキーワードと一致するか」を全列から探し、
// 一致したらその直後の列（＝値列）を返す方式にした。
// これにより、例えば将来のエクスポートでフィールド番号がずれても、ラベルさえ同じなら
// 正しく値を拾える。
function rktFindFieldByLabel_(header, row, labelKeywords) {
  for (let i = 0; i < header.length - 1; i++) {
    if (!/^フィールド\d+$/.test(header[i])) continue;

    const labelValue = rktText_(row[i]);
    if (!labelValue) continue;

    const matched = labelKeywords.some(k => labelValue.indexOf(k) !== -1);

    if (matched) {
      return rktText_(row[i + 1]);
    }
  }

  return "";
}

function rktSimplifyName_(name) {
  return rktNormalizeText_(name)
    .replace(/[\s ・、。，．・！？!?()（）【】\[\]「」『』_－\-〜~]/g, "")
    .replace(/(本店|支店|営業所|店)$/g, "")
    .trim();
}

function rktNormalizeText_(value) {
  return rktText_(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/　/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function rktNormalizeHeader_(header) {
  return header.map(h => rktText_(h).replace(/^﻿/, "").trim());
}

function rktGetValue_(header, row, names) {
  for (const name of names) {
    const idx = header.indexOf(name);

    if (idx !== -1 && row[idx] !== undefined && row[idx] !== null && rktText_(row[idx]) !== "") {
      return rktText_(row[idx]);
    }
  }

  return "";
}

function rktText_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}

function rktGetOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

function rktWrite_(sheet, values) {
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.clearContents();
  sheet.clearFormats();

  if (!values || values.length === 0) return;

  const maxCols = Math.max.apply(null, values.map(r => r.length));

  const normalized = values.map(r => {
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

function rktSafeSheetName_(name) {
  return rktText_(name).replace(/[\\\/\?\*\[\]\:]/g, "").substring(0, 80) || "エリア不明";
}

function rktGetOrCreateFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function rktArrayToCsv_(array) {
  return array.map(row => row.map(cell => {
    const str = String(cell === null || cell === undefined ? "" : cell).replace(/"/g, '""');

    if (
      str.indexOf(",") !== -1 ||
      str.indexOf("\n") !== -1 ||
      str.indexOf("\r") !== -1 ||
      str.indexOf('"') !== -1
    ) {
      return '"' + str + '"';
    }

    return str;
  }).join(",")).join("\r\n");
}


// =====================================================================
// 14. 地域別営業リスト件数サマリー
// =====================================================================

/**
 * 04_SALES_宿泊_地域名 タブを集計して、
 * 複数地域が混ざった営業リストの地域別件数を出す
 */
function rakutenCreateSalesRegionSummary(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = rktGetOrCreateSheet_(ss, "00_地域別_営業リスト件数");

  let regionCounts = rktBuildSalesRegionCountsFromAreaSheets_();

  // 地域別タブがまだない場合は、04_SALES_宿泊施設の住所から地域を推定して集計
  if (regionCounts.length === 0) {
    regionCounts = rktBuildSalesRegionCountsFromMixedSheet_();
  }

  regionCounts.sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.region.localeCompare(b.region, "ja");
  });

  const totalCount = regionCounts.reduce((sum, item) => sum + item.count, 0);
  const allSheet = ss.getSheetByName("04_SALES_宿泊施設");
  const allCount = allSheet ? Math.max(allSheet.getLastRow() - 1, 0) : 0;

  const now = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss");

  const rows = [];

  rows.push(["地域別営業リスト件数サマリー", "", "", ""]);
  rows.push(["更新日時", now, "", ""]);
  rows.push(["", "", "", ""]);
  rows.push(["全体営業リスト件数", allCount, "", ""]);
  rows.push(["地域別合計", totalCount, "", ""]);
  rows.push(["差分", allCount - totalCount, "", ""]);
  rows.push(["", "", "", ""]);
  rows.push(["地域", "営業リスト件数", "元タブ", "備考"]);

  if (regionCounts.length === 0) {
    rows.push(["地域データなし", 0, "", "04_SALES_宿泊施設または地域別タブを確認してください"]);
  } else {
    regionCounts.forEach(item => {
      rows.push([
        item.region,
        item.count,
        item.sheetName || "",
        item.note || ""
      ]);
    });
  }

  rktWrite_(summarySheet, rows);

  summarySheet.getRange(1, 1, 1, 4).setFontWeight("bold").setFontSize(13);
  summarySheet.getRange(4, 1, 3, 2).setFontWeight("bold");
  summarySheet.getRange(8, 1, 1, 4).setFontWeight("bold");
  summarySheet.autoResizeColumns(1, 4);

  ss.setActiveSheet(summarySheet);
  ss.moveActiveSheet(1);

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "地域別営業リスト件数を作成しました。\n\n" +
      `全体営業リスト件数: ${allCount}件\n` +
      `地域別合計: ${totalCount}件\n` +
      `地域数: ${regionCounts.length}件\n\n` +
      "「00_地域別_営業リスト件数」タブを確認してください。"
    );
  }

  return {
    allCount,
    totalCount,
    regionCounts
  };
}

/**
 * 04_SALES_宿泊_地域名 タブから件数を集計
 */
function rktBuildSalesRegionCountsFromAreaSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();

    if (sheetName.indexOf("04_SALES_宿泊_") !== 0) return;

    const region = sheetName.replace("04_SALES_宿泊_", "");
    const count = Math.max(sheet.getLastRow() - 1, 0);

    result.push({
      region,
      count,
      sheetName,
      note: "地域別タブから集計"
    });
  });

  return result;
}

/**
 * 地域別タブがない場合の予備処理
 * 04_SALES_宿泊施設 の 住所１ から市区町村を推定して集計
 */
function rktBuildSalesRegionCountsFromMixedSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("04_SALES_宿泊施設");

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h || "").trim());
  const rows = values.slice(1);

  const prefIdx = header.indexOf("都道府県");
  const addr1Idx = header.indexOf("住所１");

  if (addr1Idx === -1) {
    return [];
  }

  const countMap = {};

  rows.forEach(row => {
    const pref = prefIdx !== -1 ? String(row[prefIdx] || "").trim() : "";
    const addr1 = String(row[addr1Idx] || "").trim();

    const region = rktExtractRegionFromSalesAddress_(pref, addr1);

    if (!countMap[region]) {
      countMap[region] = 0;
    }

    countMap[region]++;
  });

  return Object.keys(countMap).map(region => ({
    region,
    count: countMap[region],
    sheetName: "04_SALES_宿泊施設",
    note: "住所１から地域を推定"
  }));
}

/**
 * 住所１から市区町村を抜き出す
 */
function rktExtractRegionFromSalesAddress_(pref, addr1) {
  const address = String(addr1 || "").trim();

  if (!address) return "エリア不明";

  const cityMatch = address.match(/^(.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村])/);

  if (cityMatch) {
    return cityMatch[1];
  }

  if (pref) {
    return pref + "_市区町村不明";
  }

  return "エリア不明";
}
