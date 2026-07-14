/**
 * 楽天トラベル宿泊施設専用GAS Ver2.0.0
 *
 * 目的：
 * Octoparseで取得した楽天トラベル宿泊施設リストを、
 * CSV投入フォルダ方式で自動取り込みし、
 * 個人店寄り分類・エリア分け・CSV出力まで行う。
 *
 * ★Ver2.0.0の変更点（Ver1.3.0からの差分）：
 * Octoparse側のタスクが1本にまとまり、投入フォルダに入れるCSVの時点で
 * 住所・電話番号（・部屋数）まで全て取得できている状態になったため、
 * 飲食店側のGAS（gas/list-normalizer/code.js）と同じ「投入フォルダ1つだけ」の
 * 構成に統一した。旧バージョンにあった「一覧CSV投入フォルダ／詳細CSV投入フォルダ」
 * の2フォルダ体制、詳細URL一覧の作成、Octoparseタスク2本に分けての二段階結合は
 * 廃止した（Ver1.3.0までのソースはgitの履歴を参照）。
 *
 * ★Ver1.3.0の変更点（Ver1.2.0からの差分。現在も有効な仕様）：
 * 1. 部屋数（客室数）による営業対象判定を追加。
 *    「電気切り替え条件で宿泊施設に無料HP作成」という営業は、自前で立派なHPや
 *    マーケティング体制をすでに持っている大型ホテルよりも、そうした体制を
 *    持っていない可能性が高い小規模施設（民宿・ペンション・ゲストハウス等）に
 *    向いている。部屋数が一定数（既定15室）以上の施設は、既存のチェーン系
 *    ホテル・大型旅館である可能性が高いため、チェーンキーワードに一致しなくても
 *    部屋数だけで除外できるようにした。
 *    Octoparseの「フィールドN」列はラベルと値が交互に並ぶ形式で出力されるため
 *    （実データ確認: フィールド6="総部屋数"(ラベル)→フィールド7="65室"(値)）、
 *    列番号を決め打ちせず、ラベル文字列で値列を探す方式にしてある
 *    （rktFindFieldByLabel_。列番号がズレても追従できる）。
 * 2. 宿泊系チェーンの除外キーワードを大幅に拡充（アパホテル・ルートイン・
 *    東横INN等の主要チェーンに加え、ホテルマイステイズ・ヴィアイン・
 *    チサンホテル・ワシントンホテル等）。加えて「〇〇ホテルグループ」という
 *    自己申告的な表記自体も汎用の除外条件にしている（個別ブランド名を
 *    覚えなくても同種の中小チェーンを広く拾える）。
 * 3. 04_SALES系タブの備考欄に部屋数を表示し、架電担当がリストを見ただけで
 *    規模感を把握できるようにした。
 *
 * 通常運用：
 * 1. 「0. 初期タブ・投入フォルダを作成」を実行
 * 2. Octoparseで取得した（住所・電話番号・部屋数まで入った）CSVを
 *    「楽天_CSV投入フォルダ」に入れる
 * 3. 「🚀 全自動処理：投入済みCSVからCSV出力まで実行」を実行
 */

// =====================================================================
// 0. メニュー
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏕 楽天トラベル宿泊施設")
    .addItem("0. 初期タブ・投入フォルダを作成", "rakutenSetupAll")
    .addSeparator()
    .addItem("1. CSV取り込み", "rakutenImportCsv")
    .addItem("2. 楽天 一覧を正規化", "rakutenNormalizeAndMerge")
    .addItem("3. 楽天 宿泊施設を分類・地域別タブ作成", "rakutenSplitClassifySheets")
    .addItem("4. 楽天 システム投入CSVタブ作成", "rakutenCreateSalesSheets")
    .addItem("5. 楽天 CSVをDrive出力", "rakutenExportSalesCsvFiles")
    .addSeparator()
    .addItem("🚀 全自動処理：投入済みCSVからCSV出力まで実行", "rakutenRunAllFromFolder")
    .addSeparator()
    .addItem("⏰ 自動実行トリガーを設定（15分おき）", "rakutenInstallAutoTrigger")
    .addItem("⏰ 自動実行トリガーを解除", "rakutenRemoveAutoTrigger")
    .addToUi();
}

// =====================================================================
// 0.5 自動実行トリガー
// =====================================================================
// 投入フォルダにCSVを置くだけで、スプレッドシートを開かなくても
// 自動で「🚀 全自動処理」が走るようにする時間主導型トリガー。
// スクレイパー（scraper/scrape.js）がCSVを楽天_CSV投入フォルダに置いた後、
// このトリガーが次の周期（既定15分）で拾って処理する。
// 何もCSVが無いときは rakutenRunAllFromFolder() 内の rawCount<=0 チェックで
// 即座に抜けるだけなので、無駄なCSV出力やアラートは発生しない
// （ただし alert() は時間主導トリガーからは呼べないため、トリガー実行時は
// 内部で例外を握りつぶすようにしている。詳しくは rakutenRunAllFromFolderSilent 参照）。
function rakutenInstallAutoTrigger() {
  rakutenRemoveAutoTrigger();

  ScriptApp.newTrigger("rakutenRunAllFromFolderSilent")
    .timeBased()
    .everyMinutes(15)
    .create();

  SpreadsheetApp.getUi().alert(
    "自動実行トリガーを設定しました。\n\n" +
    "以後、楽天_CSV投入フォルダにCSVを置くだけで、15分以内に自動で処理されます。\n" +
    "（スプレッドシートを開いている必要はありません）"
  );
}

function rakutenRemoveAutoTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(t => {
    if (t.getHandlerFunction() === "rakutenRunAllFromFolderSilent") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  if (removed > 0) {
    Logger.log(`[楽天] 既存の自動実行トリガーを${removed}件解除しました。`);
  }
}

// 時間主導トリガーから呼ばれる版。SpreadsheetApp.getUi()はトリガー実行時に
// 使えず例外になるため、rakutenRunAllFromFolder()本体を直接は呼ばずに
// silent=trueで内部処理だけ行うラッパーにしてある。
function rakutenRunAllFromFolderSilent() {
  try {
    rakutenRunAllFromFolder(true);
  } catch (e) {
    Logger.log(`[楽天] 自動実行中にエラー: ${e.message}`);
  }
}

// =====================================================================
// 1. 設定
// =====================================================================
const RAKUTEN_SHEETS = {
  raw: "楽天_raw",
  normalized: "楽天_正規化",
  target: "楽天_営業対象",
  confirm: "楽天_要確認",
  exclude: "楽天_除外",
  duplicate: "楽天_重複"
};

const RAKUTEN_FOLDER_NAMES = {
  input: "楽天_CSV投入フォルダ",
  processed: "楽天_処理済みフォルダ",
  export: "完成版CSVエクスポート"
};

// Octoparseの実際の出力列名（茨城県宿泊.csvで確認済み）。
// 「フィールド1_テキスト」＝施設名、「フィールド1_リンク」＝詳細ページURL。
const RAKUTEN_RAW_NAME_COLUMNS = ["フィールド1_テキスト_テキスト", "フィールド1_テキスト", "jsraleventscroll", "タイトル", "施設名", "店名", "名前"];
const RAKUTEN_RAW_URL_COLUMNS = ["フィールド1_リンク_リンク", "フィールド1_リンク", "jsraleventscroll_URL", "タイトルURL", "フィールド", "URL", "楽天URL"];

// 部屋数（客室数）列の候補名。
// ラベル文字列一致（rktFindFieldByLabel_）で見つからなかった場合の
// フォールバック用。実際の列名が別途分かれば、この配列の先頭に追加するだけで反映される。
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

// 部屋数がこの数値以上なら「規模が大きく、自前でHP・マーケティング体制を
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
  "部屋数"
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

// 全国チェーン系ホテル・旅館ブランド。
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
  // 「〇〇ホテルグループ」という自己申告的な表記自体を汎用の除外条件にした
  // （実データ「BBHホテルグループ」で確認。個別ブランド名を覚えなくても
  // 同種の中小チェーンを広く拾える）。
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
    "・楽天_CSV投入フォルダ\n" +
    "・楽天_処理済みフォルダ\n" +
    "・完成版CSVエクスポート"
  );
}

function rakutenSetupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const initialSheets = [
    RAKUTEN_SHEETS.raw,
    RAKUTEN_SHEETS.normalized,
    RAKUTEN_SHEETS.target,
    RAKUTEN_SHEETS.confirm,
    RAKUTEN_SHEETS.exclude,
    RAKUTEN_SHEETS.duplicate
  ];

  initialSheets.forEach(name => rktGetOrCreateSheet_(ss, name));
}

function rakutenCreateInputFolders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.input);
  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.processed);
  rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.export);
}

// =====================================================================
// 3. フォルダ投入型：CSV取り込み（投入フォルダは1つだけ）
// =====================================================================
function rakutenImportCsv() {
  rakutenCreateInputFolders_();

  const imported = rakutenImportCsvFiles();

  if (!imported || imported.rows === 0) {
    SpreadsheetApp.getUi().alert(
      "CSVの取り込みができませんでした。\n\n" +
      "楽天_CSV投入フォルダにOctoparseのCSVを入れてから再実行してください。"
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    "CSVの取り込みが完了しました。\n\n" +
    `取り込みファイル数: ${imported.files}件\n` +
    `データ件数: ${imported.rows}件\n\n` +
    "次に「🚀 全自動処理」を実行してください。"
  );
}

function rakutenImportCsvFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  const inputFolder = rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.input);
  const processedFolder = rktGetOrCreateFolder_(parentFolder, RAKUTEN_FOLDER_NAMES.processed);
  const targetSheet = rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.raw);

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
// 4. 全自動処理
// =====================================================================
// silent=true の場合はUIアラートを一切呼ばない。
// 時間主導トリガー（rakutenRunAllFromFolderSilent経由）から呼ばれたときは
// UIコンテキストが存在せず SpreadsheetApp.getUi() が例外になるため、
// メニューから手動実行する場合(silent省略=false)とトリガーからの自動実行
// (silent=true)を明示的に分けている。
function rakutenRunAllFromFolder(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  rakutenCreateInputFolders_();
  rakutenImportCsvFiles();

  const rawSheet = ss.getSheetByName(RAKUTEN_SHEETS.raw);
  const rawValues = rawSheet ? rawSheet.getDataRange().getValues() : [];
  const rawCount = Math.max(rawValues.length - 1, 0);

  if (rawCount <= 0) {
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        "データがありません。\n\n" +
        "楽天_CSV投入フォルダにOctoparseのCSVを入れてから再実行してください。"
      );
    }
    return;
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

  Logger.log(
    `[楽天] 全自動処理完了: 取り込み${rawCount}件 / 営業対象${targetCount}件 / ` +
    `要確認${confirmCount}件 / 除外${excludeCount}件 / 重複${duplicateCount}件`
  );

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "楽天宿泊施設リストの全自動処理が完了しました。\n\n" +
      `取り込みデータ: ${rawCount}件\n\n` +
      `営業対象: ${targetCount}件\n` +
      `要確認: ${confirmCount}件\n` +
      `除外: ${excludeCount}件\n` +
      `重複: ${duplicateCount}件\n\n` +
      "04_SALES_宿泊施設 / 市区町村別タブ / 完成版CSVエクスポートを確認してください。"
    );
  }
}

// 後方互換用エイリアス：Ver1.3.0までは関数名が複数形「rakutenRunAllFromFolders」だった。
// スプレッドシート上のボタン（図形描画）やトリガーが古い関数名のまま紐付いている場合、
// Ver2.0.0で単数形「rakutenRunAllFromFolder」に変更したことで
// 「スクリプト関数が見つかりません」エラーになるため、旧名でも動くようにしておく。
// ボタンを付け直す場合は新しい方（rakutenRunAllFromFolder）を割り当ててよい。
function rakutenRunAllFromFolders() {
  return rakutenRunAllFromFolder();
}

// =====================================================================
// 5. 楽天_raw → 楽天_正規化
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

  const outputRows = rktBuildNormalizedRows_(rawHeader, rawRows);

  const sheet = rktGetOrCreateSheet_(ss, RAKUTEN_SHEETS.normalized);
  rktWrite_(sheet, [RAKUTEN_NORMALIZED_HEADER].concat(outputRows));

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      `楽天データの正規化が完了しました。\n出力件数: ${outputRows.length}件`
    );
  }

  return { count: outputRows.length };
}

// 楽天_raw（1行1施設。住所・電話番号・部屋数まで取得済みのCSV）→ 正規化済み行の配列。
function rktBuildNormalizedRows_(rawHeader, rawRows) {
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
    // 列番号はOctoparseのタスク設定によりズレる可能性があるため、
    // まずラベル文字列で探し（rktFindFieldByLabel_）、見つからない場合のみ
    // 固定列番号候補にフォールバックする。
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

    // FAX：Octoparseのフィールドラベル形式（実データでは存在しないケースが多い）に加えて、
    // 「FAX」という素の列名も候補にした（スクレイパー等、自前でCSVを組み立てる側が
    // フィールドN形式を使わずそのまま"FAX"列を出力してくるケースに対応するため）。
    const faxRaw = rktFindFieldByLabel_(rawHeader, rawRow, ["FAX"])
      || rktGetValue_(rawHeader, rawRow, ["FAX"]);
    const fax = rktIsDummyFax_(faxRaw) ? "" : rktNormalizePhoneDisplay_(faxRaw);

    const description = rktGetValue_(rawHeader, rawRow, ["hotelcharacter", "施設説明", "説明"]);
    // 価格：実データでは「価格」列自体は "[最安料金（目安）]" という見出し文字列が入っており、
    // 実際の金額は価格1（例:"3,273円～"）・価格2（税込表記）に入っていたため、
    // 価格1・価格2を優先する（茨城県宿泊.csvで確認）。
    const price = rktGetValue_(rawHeader, rawRow, ["価格1", "価格2", "incldtax", "plnprc", "料金"]);
    const review = rktGetValue_(rawHeader, rawRow, ["hotelrating", "cstmrevl", "口コミ"]);
    const imageUrl = rktGetValue_(rawHeader, rawRow, ["画像URL", "画像URL1"]);
    // タグ：実データにはmoreplan3やタグ15に相当する列がなく、代わりに"planoutline"
    // （プラン説明文）に「グランピング」等のジャンルを示す語が入ることがあったため、
    // ジャンル・営業対象判定用のテキストにのみ含める（表示用タグ列としては汚いので出さない）。
    // "設備タグ"は自前スクレイパー側が直接出してくる場合の素の列名候補として追加。
    const planOutlineRaw = rktGetValue_(rawHeader, rawRow, ["planoutline"]);
    const tags = rktGetValue_(rawHeader, rawRow, ["moreplan3", "タグ15", "設備タグ"]);
    const judgeText = tags + " " + rktNormalizeText_(planOutlineRaw).slice(0, 200);

    const genre = rktJudgeLodgingGenre_(name, description, judgeText);

    // 部屋数（客室数）。実データではフィールド6="総部屋数"(ラベル)→
    // フィールド7=値（例:"65室"）という並びだったため、まずラベル一致で探し、
    // 見つからなければ候補名リスト（RAKUTEN_ROOM_COUNT_COLUMNS）にフォールバックする。
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
      roomCount === null ? "" : String(roomCount)
    ]);
  });

  return outputRows;
}

// =====================================================================
// 6. 楽天_正規化 → 営業対象/要確認/除外/重複/エリア別
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
// 7. 楽天_営業対象 → 04_SALES_宿泊施設 / 市区町村別
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
// 8. 04_SALES_宿泊系 → Drive CSV出力
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
// 9. 判定系ヘルパー
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

// 「15室」「全15室」「客室数：20」「20部屋」等の文字列から数字部分だけを取り出す。
// 数字が取れなければnullを返す（＝部屋数不明。部屋数フィルタは適用せず、他の判定基準に委ねる）。
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
// 10. システム投入CSV生成ヘルパー
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
  const roomCount = rktGetValue_(header, row, ["部屋数"]);
  const cleanPhone = rktNormalizePhoneDigits_(phone);

  const addr1 = rktBuildAddress1_(pref, city, address);
  const areaText = pref + city;
  const bp = areaText + "tel" + cleanPhone;

  const memoParts = [];

  if (genre) memoParts.push("宿泊ジャンル: " + genre);
  if (roomCount) memoParts.push("部屋数: " + roomCount + "室");
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
// 11. 汎用ヘルパー
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
// 12. 地域別営業リスト件数サマリー
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
