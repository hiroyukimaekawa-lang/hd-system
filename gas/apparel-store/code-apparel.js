/**
 * アパレル 店舗管理システム Ver2.5.0
 * アパレルチェーン・ドラッグストア・カインズ・ホームセンター・大型小売除外強化版
 *
 * ★Ver2.5.0の変更点（Ver2.4.0からの差分）：
 * 1. `apparelIsLikelyBranchStoreName_` を強化し、マスタに登録していない
 *    未知のチェーン店（「〇〇下妻店」「〇〇 古河店」のような地名・区切り記号
 *    付きの支店名）も自動検知できるようにした。
 * 2. ただし、この自動検知には「本当にチェーン店か、単独店がたまたま
 *    似た命名をしているだけか」を機械的には区別できないという弱点がある
 *    （地名付きの「〇〇店」という命名自体は、チェーン店でなくても普通に
 *    使われることがあるため）。誤って独立系の小規模店舗を営業リストから
 *    永久に落としてしまうと、その店舗には二度と営業をかけられなくなり、
 *    機会損失が大きい。そこで検知結果を確度で2段階に分けた：
 *      - "high"（号店・駅前店・モール系テナント名など、ほぼ確実にチェーンと
 *        分かる強いシグナル）→ 従来通り「チェーン店」として即除外
 *      - "heuristic"（地名＋店、スペース/括弧区切りの「〇〇店」など、
 *        チェーンの可能性は高いが確証はないシグナル）→ 「チェーン店疑い」
 *        として一旦「02_確認対象」に回し、人の目で最終確認してもらう
 *    チェーンマスタ（APPAREL_KNOWN_CHAIN_GAPS）に明示的に登録されている
 *    既知チェーンの判定は、当然ながらこれまで通り確実な除外のまま。
 * 3. チェーンマスタに Workman（英語表記）・買取大吉・おたからや・
 *    ジュエルカフェを追加。
 *
 * 運用：
 * 1. 「アパレル_CSV投入フォルダ」にCSVを入れる
 * 2. スプレッドシートを再読み込み
 * 3. メニュー「👕 アパレル」→「🚀 すべての一括処理を実行」
 */

// =====================================================================
// メニュー
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("👕 アパレル")
    .addItem("🚀 すべての一括処理を実行", "apparelExecuteAllProcesses")
    .addSeparator()
    .addItem("0. 初期フォルダ・タブ作成", "apparelSetupAll")
    .addItem("📁 1. CSVを一括取り込み", "apparelImportCSVFiles")
    .addItem("2. 正規化・基本判定", "apparelExecuteNormalizeAndValidate")
    .addItem("3. 重複判定", "apparelExecuteDuplicateCheck")
    .addItem("4. チェーン判定", "apparelExecuteChainCheck")
    .addItem("5. 施設判定", "apparelExecuteFacilityCheck")
    .addItem("6. ワークフロー分類", "apparelExecuteWorkflowGrouping")
    .addItem("7. タブ分け", "apparelExecuteSplitSheets")
    .addItem("8. 04_SALES_地域別タブ生成", "apparelExecuteGenerateSalesAreaSheets")
    .addItem("9. 04_SALES_CSVをDrive出力", "apparelExecuteExportSalesAreaCsvFiles")
    .addSeparator()
    .addItem("📊 件数サマリーを更新", "apparelExecuteCountSummary")
    .addItem("🔧 チェーンマスタの不足キーワードを追加", "apparelFixKnownChainMasterGaps")
    .addToUi();
}

// =====================================================================
// 設定
// =====================================================================
const APPAREL_FOLDER_NAMES = {
  input: "アパレル_CSV投入フォルダ",
  processed: "アパレル_処理済みフォルダ",
  export: "完成版CSVエクスポート"
};

const APPAREL_SHEETS = {
  normalized: "アパレル_01_NORMALIZED",
  duplicate: "アパレル_02_DUPLICATE_CHECK",
  chain: "アパレル_03_CHAIN_CHECK",
  facility: "アパレル_04_FACILITY_CHECK",
  target: "アパレル_01_営業対象",
  confirm: "アパレル_02_確認対象",
  exclude: "アパレル_03_除外対象",
  failed: "アパレル_04_取得失敗",
  summary: "アパレル_00_件数サマリー",
  masterChain: "アパレル_MASTER_CHAIN"
};

const APPAREL_NORMALIZED_HEADER = [
  "店名", "ジャンル", "検索ジャンル", "取得元ジャンル", "都道府県", "市区町村", "郵便番号", "住所", "電話番号", "URL", "媒体",
  "HP有無", "営業日", "定休日", "営業開始A", "営業終了A", "営業開始B", "営業終了B", "取得ステータス", "除外理由",
  "正規化電話番号", "正規化店名", "正規化ジャンル", "住所判定", "エリア判定", "エリア判定理由",
  "基本データ判定", "基本データ除外理由"
];

const APPAREL_COMDESK_HEADER = [
  "UUID", "種別", "名前", "カナ", "郵便番号", "都道府県", "住所１", "住所２", "住所カナ",
  "Tel1", "Tel2", "Tel3", "Tel4", "FAX", "URL", "備考", "旧社名", "リードソース",
  "旧進捗", "履歴", "オーナー名", "HPある？", "BP検索", "アポ済商材", "最新履歴", "営業曜日", "休業曜日",
  "午前始", "午前終", "午後始", "午後終"
];

const APPAREL_GENRE_MAP = {
  "婦人服": "レディースファッション",
  "レディース服": "レディースファッション",
  "レディースショップ": "レディースファッション",
  "レディースファッション": "レディースファッション",
  "紳士服": "メンズファッション",
  "メンズ服": "メンズファッション",
  "メンズショップ": "メンズファッション",
  "メンズウェア": "メンズファッション",
  "メンズファッション": "メンズファッション",
  "セレクトショップ": "セレクトショップ",
  "古着屋": "古着・vintage",
  "古着店": "古着・vintage",
  "古着": "古着・vintage",
  "ヴィンテージ": "古着・vintage",
  "ビンテージ": "古着・vintage",
  "リサイクルショップ": "古着・vintage",
  "呉服店": "呉服・和装",
  "着物": "呉服・和装",
  "和装": "呉服・和装",
  "着物レンタル": "呉服・和装",
  "子ども服": "子供服・ベビー服",
  "子供服": "子供服・ベビー服",
  "キッズ服": "子供服・ベビー服",
  "ベビー服": "子供服・ベビー服",
  "ベビー用品": "子供服・ベビー服",
  "フォーマル": "ブライダル・フォーマル",
  "礼服": "ブライダル・フォーマル",
  "ウェディングドレス": "ブライダル・フォーマル",
  "ブライダルショップ": "ブライダル・フォーマル",
  "ランジェリー": "下着・ランジェリー",
  "インナーウェア": "下着・ランジェリー",
  "下着店": "下着・ランジェリー",
  "下着": "下着・ランジェリー",
  "衣料品店": "その他アパレル",
  "洋服店": "その他アパレル",
  "服屋": "その他アパレル",
  "アパレルショップ": "その他アパレル",
  "ファッション雑貨店": "その他アパレル",
  "アパレル": "その他アパレル"
};

const APPAREL_TARGET_GENRES = [
  "レディースファッション",
  "メンズファッション",
  "セレクトショップ",
  "古着・vintage",
  "呉服・和装",
  "子供服・ベビー服",
  "ブライダル・フォーマル",
  "下着・ランジェリー",
  "その他アパレル"
];

// 施設・大型小売として除外するキーワード
const APPAREL_FACILITY_EXCLUDE_KEYWORDS = [
  "イオンモール", "イオン", "AEON", "ららぽーと", "アリオ", "パルコ", "PARCO", "ルミネ", "LUMINE",
  "アトレ", "エキュート", "マルイ", "OIOI", "百貨店", "高島屋", "伊勢丹", "三越", "そごう",
  "大丸", "松坂屋", "阪急", "近鉄", "ショッピングセンター", "ショッピングモール", "アウトレット",
  "駅ビル", "テナント", "フロアガイド",
  "ドラッグストア", "ドラッグ", "薬局", "調剤薬局", "くすり", "クスリ", "マツモトキヨシ", "マツキヨ",
  "ウエルシア", "ウェルシア", "ツルハドラッグ", "ツルハ", "サンドラッグ", "ドラッグストアコスモス", "コスモス薬品",
  "スギ薬局", "スギドラッグ", "ココカラファイン", "クリエイトSD", "クリエイトエス・ディー", "カワチ薬品",
  "クスリのアオキ", "ドラッグセイムス", "セイムス", "キリン堂", "ゲンキー", "ヤックスドラッグ", "薬王堂",
  "カインズ", "CAINZ", "コメリ", "DCM", "ケーヨーデイツー", "ケーヨーD2", "コーナン", "ビバホーム",
  "スーパービバホーム", "ジョイフル本田", "島忠", "ホームズ", "ナフコ", "ロイヤルホームセンター", "ホームセンター",
  "ドンキホーテ", "ドンキ", "ヨーカドー", "イトーヨーカドー", "アピタ", "ピアゴ", "西友", "ライフ",
  "マックスバリュ", "ベイシア", "トライアル", "業務スーパー", "ヤオコー", "カスミ", "ロピア",
  "ニトリ", "ヤマダデンキ", "ヤマダ電機", "ケーズデンキ", "エディオン", "ビックカメラ", "ヨドバシカメラ",
  "ダイソー", "DAISO", "セリア", "Seria", "キャンドゥ", "Can Do", "ワッツ", "Watts",
  "雑貨店", "生活雑貨", "家具", "家電", "スーパー", "コンビニ", "書店", "本屋", "買取", "質屋"
];

const APPAREL_FACILITY_REVIEW_KEYWORDS = [
  "ビル", "プラザ", "タワー", "センター", "B1F", "1F", "2F", "3F", "4F", "5F", "階", "地下"
];

// [チェーン名, キーワード, 業種, 有効, メモ]
const APPAREL_KNOWN_CHAIN_GAPS = [
  // アパレル大手・ファストファッション
  ["ユニクロ", "ユニクロ", "アパレル", true, "大手アパレルチェーン"],
  ["ユニクロ", "UNIQLO", "アパレル", true, "英語表記"],
  ["GU", "ジーユー", "アパレル", true, "大手アパレルチェーン"],
  ["GU", "GU", "アパレル", true, "英語表記"],
  ["しまむら", "しまむら", "アパレル", true, "大手アパレルチェーン"],
  ["しまむら", "ファッションセンターしまむら", "アパレル", true, "正式名称対策"],
  ["アベイル", "アベイル", "アパレル", true, "しまむらグループ"],
  ["バースデイ", "バースデイ", "アパレル", true, "しまむらグループ子供服"],
  ["西松屋", "西松屋", "アパレル", true, "子供服大手チェーン"],
  ["無印良品", "無印良品", "アパレル", true, "大型小売・アパレルあり"],
  ["ワークマン", "ワークマン", "アパレル", true, "大手作業服・アパレルチェーン"],
  ["ワークマン女子", "ワークマン女子", "アパレル", true, "ワークマン系列"],
  ["WORKMAN Plus", "WORKMAN Plus", "アパレル", true, "ワークマン系列"],
  // ↓Ver2.5.0で追加。英語表記「Workman」単体（例:"Workman Colors"）は
  // 既存の「ワークマン」「WORKMAN Plus」いずれとも一致しなかったため追加。
  ["Workman", "Workman", "アパレル", true, "英語表記単体（Workman Colors等の派生店名にも一致）"],
  ["PLST", "PLST", "アパレル", true, "ファーストリテイリング系列"],
  ["プラステ", "プラステ", "アパレル", true, "ファーストリテイリング系列"],
  ["ZARA", "ZARA", "アパレル", true, "海外大手ファッション"],
  ["H&M", "H&M", "アパレル", true, "海外大手ファッション"],
  ["GAP", "GAP", "アパレル", true, "海外大手ファッション"],
  ["GAP", "ギャップ", "アパレル", true, "カタカナ表記"],
  ["FOREVER21", "FOREVER21", "アパレル", true, "海外大手ファッション"],
  ["SHEIN", "SHEIN", "アパレル", true, "大型アパレルブランド"],

  // スーツ・ビジネスウェア
  ["洋服の青山", "洋服の青山", "アパレル", true, "スーツ大手チェーン"],
  ["AOKI", "AOKI", "アパレル", true, "スーツ大手チェーン"],
  ["ORIHICA", "ORIHICA", "アパレル", true, "AOKI系列"],
  ["オリヒカ", "オリヒカ", "アパレル", true, "AOKI系列"],
  ["コナカ", "コナカ", "アパレル", true, "スーツ大手チェーン"],
  ["SUIT SELECT", "SUIT SELECT", "アパレル", true, "コナカ系列"],
  ["スーツセレクト", "スーツセレクト", "アパレル", true, "コナカ系列"],
  ["はるやま", "はるやま", "アパレル", true, "スーツ大手チェーン"],
  ["P.S.FA", "P.S.FA", "アパレル", true, "はるやま系列"],
  ["パーフェクトスーツファクトリー", "パーフェクトスーツファクトリー", "アパレル", true, "はるやま系列"],
  ["タカキュー", "タカキュー", "アパレル", true, "ビジネスウェア大手"],
  ["TAKA-Q", "TAKA-Q", "アパレル", true, "英語表記"],
  ["鎌倉シャツ", "鎌倉シャツ", "アパレル", true, "ビジネスシャツチェーン"],
  ["メーカーズシャツ鎌倉", "メーカーズシャツ鎌倉", "アパレル", true, "正式名称"],

  // カジュアル・ジーンズ・量販系
  ["ライトオン", "ライトオン", "アパレル", true, "大手カジュアルチェーン"],
  ["Right-on", "Right-on", "アパレル", true, "英語表記"],
  ["ジーンズメイト", "ジーンズメイト", "アパレル", true, "大手カジュアルチェーン"],
  ["マックハウス", "マックハウス", "アパレル", true, "大手カジュアルチェーン"],
  ["Mac-House", "Mac-House", "アパレル", true, "英語表記"],
  ["ハニーズ", "ハニーズ", "アパレル", true, "大手レディースチェーン"],
  ["Honeys", "Honeys", "アパレル", true, "英語表記"],
  ["WEGO", "WEGO", "アパレル", true, "大手カジュアルチェーン"],
  ["ウィゴー", "ウィゴー", "アパレル", true, "カタカナ表記"],
  ["SPINNS", "SPINNS", "アパレル", true, "若者向けチェーン"],
  ["スピンズ", "スピンズ", "アパレル", true, "カタカナ表記"],
  ["coca", "coca", "アパレル", true, "低価格アパレルチェーン"],
  ["コカ", "コカ", "アパレル", true, "カタカナ表記"],
  ["ベルーナ", "ベルーナ", "アパレル", true, "ミセス系アパレルチェーン"],
  ["BELLUNA", "BELLUNA", "アパレル", true, "英語表記"],
  ["レトロガール", "レトロガール", "アパレル", true, "大手レディースチェーン"],
  ["RETRO GIRL", "RETRO GIRL", "アパレル", true, "英語表記"],
  ["イング", "イング", "アパレル", true, "レディースチェーン"],
  ["INGNI", "INGNI", "アパレル", true, "英語表記"],

  // アダストリア系
  ["グローバルワーク", "グローバルワーク", "アパレル", true, "アダストリア系"],
  ["GLOBAL WORK", "GLOBAL WORK", "アパレル", true, "英語表記"],
  ["ローリーズファーム", "ローリーズファーム", "アパレル", true, "アダストリア系"],
  ["LOWRYS FARM", "LOWRYS FARM", "アパレル", true, "英語表記"],
  ["ニコアンド", "ニコアンド", "アパレル", true, "アダストリア系"],
  ["niko and", "niko and", "アパレル", true, "英語表記"],
  ["レプシィム", "レプシィム", "アパレル", true, "アダストリア系"],
  ["LEPSIM", "LEPSIM", "アパレル", true, "英語表記"],
  ["スタディオクリップ", "スタディオクリップ", "アパレル", true, "アダストリア系"],
  ["studio CLIP", "studio CLIP", "アパレル", true, "英語表記"],
  ["ベイフロー", "ベイフロー", "アパレル", true, "アダストリア系"],
  ["BAYFLOW", "BAYFLOW", "アパレル", true, "英語表記"],
  ["ジーナシス", "ジーナシス", "アパレル", true, "アダストリア系"],
  ["JEANASIS", "JEANASIS", "アパレル", true, "英語表記"],
  ["ヘザー", "ヘザー", "アパレル", true, "アダストリア系"],
  ["Heather", "Heather", "アパレル", true, "英語表記"],
  ["HARE", "HARE", "アパレル", true, "アダストリア系"],
  ["RAGEBLUE", "RAGEBLUE", "アパレル", true, "アダストリア系"],
  ["レイジブルー", "レイジブルー", "アパレル", true, "カタカナ表記"],

  // ストライプ・キャン系
  ["アースミュージック", "アースミュージック", "アパレル", true, "ストライプ系"],
  ["earth music", "earth music", "アパレル", true, "英語表記"],
  ["アメリカンホリック", "アメリカンホリック", "アパレル", true, "ストライプ系"],
  ["AMERICAN HOLIC", "AMERICAN HOLIC", "アパレル", true, "英語表記"],
  ["グリーンパークス", "グリーンパークス", "アパレル", true, "ストライプ系"],
  ["Green Parks", "Green Parks", "アパレル", true, "英語表記"],
  ["サマンサモスモス", "サマンサモスモス", "アパレル", true, "キャン系"],
  ["Samansa Mos2", "Samansa Mos2", "アパレル", true, "英語表記"],
  ["SM2", "SM2", "アパレル", true, "略称表記"],
  ["テチチ", "テチチ", "アパレル", true, "キャン系"],
  ["Te chichi", "Te chichi", "アパレル", true, "英語表記"],
  ["クラフトスタンダードブティック", "クラフトスタンダード", "アパレル", true, "ストライプ系"],
  ["CRAFT STANDARD BOUTIQUE", "CRAFT STANDARD BOUTIQUE", "アパレル", true, "英語表記"],

  // ワールド系・オンワード系・百貨店ブランド
  ["シューラルー", "シューラルー", "アパレル", true, "ワールド系"],
  ["SHOO-LA-RUE", "SHOO-LA-RUE", "アパレル", true, "英語表記"],
  ["オペークドットクリップ", "オペーク", "アパレル", true, "ワールド系"],
  ["OPAQUE.CLIP", "OPAQUE", "アパレル", true, "英語表記"],
  ["grove", "grove", "アパレル", true, "ワールド系"],
  ["グローブ", "グローブ", "アパレル", true, "カタカナ表記"],
  ["index", "index", "アパレル", true, "ワールド系"],
  ["インデックス", "インデックス", "アパレル", true, "カタカナ表記"],
  ["UNTITLED", "UNTITLED", "アパレル", true, "ワールド系"],
  ["アンタイトル", "アンタイトル", "アパレル", true, "カタカナ表記"],
  ["INDIVI", "INDIVI", "アパレル", true, "ワールド系"],
  ["インディヴィ", "インディヴィ", "アパレル", true, "カタカナ表記"],
  ["TAKEO KIKUCHI", "TAKEO KIKUCHI", "アパレル", true, "ワールド系"],
  ["タケオキクチ", "タケオキクチ", "アパレル", true, "カタカナ表記"],
  ["23区", "23区", "アパレル", true, "オンワード系"],
  ["組曲", "組曲", "アパレル", true, "オンワード系"],
  ["自由区", "自由区", "アパレル", true, "オンワード系"],
  ["ICB", "ICB", "アパレル", true, "オンワード系"],
  ["any SiS", "any SiS", "アパレル", true, "オンワード系"],
  ["any FAM", "any FAM", "アパレル", true, "オンワード系"],

  // セレクトショップ大手
  ["ビームス", "ビームス", "アパレル", true, "大手セレクトショップ"],
  ["BEAMS", "BEAMS", "アパレル", true, "英語表記"],
  ["シップス", "シップス", "アパレル", true, "大手セレクトショップ"],
  ["SHIPS", "SHIPS", "アパレル", true, "英語表記"],
  ["ユナイテッドアローズ", "ユナイテッドアローズ", "アパレル", true, "大手セレクトショップ"],
  ["UNITED ARROWS", "UNITED ARROWS", "アパレル", true, "英語表記"],
  ["グリーンレーベルリラクシング", "グリーンレーベル", "アパレル", true, "UA系列"],
  ["green label relaxing", "green label relaxing", "アパレル", true, "英語表記"],
  ["ビューティーアンドユース", "ビューティーアンドユース", "アパレル", true, "UA系列"],
  ["BEAUTY&YOUTH", "BEAUTY&YOUTH", "アパレル", true, "英語表記"],
  ["ナノユニバース", "ナノユニバース", "アパレル", true, "大手セレクトショップ"],
  ["nano universe", "nano universe", "アパレル", true, "英語表記"],
  ["ジャーナルスタンダード", "ジャーナルスタンダード", "アパレル", true, "ベイクルーズ系"],
  ["JOURNAL STANDARD", "JOURNAL STANDARD", "アパレル", true, "英語表記"],
  ["トゥモローランド", "トゥモローランド", "アパレル", true, "大手セレクトショップ"],
  ["TOMORROWLAND", "TOMORROWLAND", "アパレル", true, "英語表記"],
  ["アーバンリサーチ", "アーバンリサーチ", "アパレル", true, "大手セレクトショップ"],
  ["URBAN RESEARCH", "URBAN RESEARCH", "アパレル", true, "英語表記"],
  ["フリークスストア", "フリークスストア", "アパレル", true, "大手セレクトショップ"],
  ["FREAK'S STORE", "FREAK'S STORE", "アパレル", true, "英語表記"],
  ["アダムエロペ", "アダムエロペ", "アパレル", true, "ジュン系"],
  ["ADAM ET ROPE", "ADAM ET ROPE", "アパレル", true, "英語表記"],
  ["ロペピクニック", "ロペピクニック", "アパレル", true, "ジュン系"],
  ["ROPE' PICNIC", "ROPE' PICNIC", "アパレル", true, "英語表記"],

  // パル系・ルームウェア・レディースブランド
  ["3COINS", "3COINS", "大型小売", true, "パル系大型雑貨・除外"],
  ["スリーコインズ", "スリーコインズ", "大型小売", true, "カタカナ表記"],
  ["Kastane", "Kastane", "アパレル", true, "パル系"],
  ["カスタネ", "カスタネ", "アパレル", true, "カタカナ表記"],
  ["mystic", "mystic", "アパレル", true, "パル系"],
  ["ミスティック", "ミスティック", "アパレル", true, "カタカナ表記"],
  ["CIAOPANIC", "CIAOPANIC", "アパレル", true, "パル系"],
  ["チャオパニック", "チャオパニック", "アパレル", true, "カタカナ表記"],
  ["Discoat", "Discoat", "アパレル", true, "パル系"],
  ["ディスコート", "ディスコート", "アパレル", true, "カタカナ表記"],
  ["COLONY 2139", "COLONY 2139", "アパレル", true, "パル系"],
  ["コロニー", "コロニー", "アパレル", true, "カタカナ表記"],
  ["スナイデル", "スナイデル", "アパレル", true, "マッシュ系"],
  ["SNIDEL", "SNIDEL", "アパレル", true, "英語表記"],
  ["ジェラートピケ", "ジェラートピケ", "アパレル", true, "マッシュ系"],
  ["gelato pique", "gelato pique", "アパレル", true, "英語表記"],
  ["ミラオーウェン", "ミラオーウェン", "アパレル", true, "マッシュ系"],
  ["Mila Owen", "Mila Owen", "アパレル", true, "英語表記"],
  ["フレイアイディー", "フレイアイディー", "アパレル", true, "マッシュ系"],
  ["FRAY I.D", "FRAY I.D", "アパレル", true, "英語表記"],

  // スポーツ・アウトドア
  ["ナイキ", "ナイキ", "アパレル", true, "スポーツ大手"],
  ["NIKE", "NIKE", "アパレル", true, "英語表記"],
  ["アディダス", "アディダス", "アパレル", true, "スポーツ大手"],
  ["adidas", "adidas", "アパレル", true, "英語表記"],
  ["プーマ", "プーマ", "アパレル", true, "スポーツ大手"],
  ["PUMA", "PUMA", "アパレル", true, "英語表記"],
  ["ニューバランス", "ニューバランス", "アパレル", true, "スポーツ大手"],
  ["New Balance", "New Balance", "アパレル", true, "英語表記"],
  ["アシックス", "アシックス", "アパレル", true, "スポーツ大手"],
  ["ASICS", "ASICS", "アパレル", true, "英語表記"],
  ["アンダーアーマー", "アンダーアーマー", "アパレル", true, "スポーツ大手"],
  ["UNDER ARMOUR", "UNDER ARMOUR", "アパレル", true, "英語表記"],
  ["ノースフェイス", "ノースフェイス", "アパレル", true, "アウトドア大手"],
  ["THE NORTH FACE", "THE NORTH FACE", "アパレル", true, "英語表記"],
  ["コロンビア", "コロンビア", "アパレル", true, "アウトドア大手"],
  ["Columbia", "Columbia", "アパレル", true, "英語表記"],
  ["モンベル", "モンベル", "アパレル", true, "アウトドア大手"],
  ["mont-bell", "mont-bell", "アパレル", true, "英語表記"],
  ["パタゴニア", "パタゴニア", "アパレル", true, "アウトドア大手"],
  ["Patagonia", "Patagonia", "アパレル", true, "英語表記"],
  ["スノーピーク", "スノーピーク", "アパレル", true, "アウトドア大手"],
  ["Snow Peak", "Snow Peak", "アパレル", true, "英語表記"],

  // 靴・バッグ・下着・ベビー
  ["ABC-MART", "ABC-MART", "アパレル", true, "靴大手チェーン"],
  ["ABCマート", "ABCマート", "アパレル", true, "カタカナ表記"],
  ["REGAL SHOES", "REGAL SHOES", "アパレル", true, "靴大手チェーン"],
  ["リーガルシューズ", "リーガルシューズ", "アパレル", true, "カタカナ表記"],
  ["ASBee", "ASBee", "アパレル", true, "靴大手チェーン"],
  ["アスビー", "アスビー", "アパレル", true, "カタカナ表記"],
  ["ワコール", "ワコール", "アパレル", true, "下着大手"],
  ["Wacoal", "Wacoal", "アパレル", true, "英語表記"],
  ["トリンプ", "トリンプ", "アパレル", true, "下着大手"],
  ["Triumph", "Triumph", "アパレル", true, "英語表記"],
  ["アモスタイル", "アモスタイル", "アパレル", true, "トリンプ系"],
  ["AMO'S STYLE", "AMO'S STYLE", "アパレル", true, "英語表記"],
  ["チュチュアンナ", "チュチュアンナ", "アパレル", true, "下着・靴下大手"],
  ["tutuanna", "tutuanna", "アパレル", true, "英語表記"],
  ["靴下屋", "靴下屋", "アパレル", true, "タビオ系"],
  ["Tabio", "Tabio", "アパレル", true, "英語表記"],
  ["アカチャンホンポ", "アカチャンホンポ", "アパレル", true, "ベビー用品大手"],
  ["赤ちゃん本舗", "赤ちゃん本舗", "アパレル", true, "漢字表記"],
  ["ミキハウス", "ミキハウス", "アパレル", true, "子供服大手"],
  ["MIKI HOUSE", "MIKI HOUSE", "アパレル", true, "英語表記"],
  ["サマンサタバサ", "サマンサタバサ", "アパレル", true, "バッグ大手"],
  ["Samantha Thavasa", "Samantha Thavasa", "アパレル", true, "英語表記"],

  // 古着・リユース大手
  ["セカンドストリート", "セカンドストリート", "アパレル", true, "リユース大手"],
  ["2nd STREET", "2nd STREET", "アパレル", true, "英語表記"],
  ["トレファクスタイル", "トレファクスタイル", "アパレル", true, "リユース大手"],
  ["TreFacStyle", "TreFacStyle", "アパレル", true, "英語表記"],
  ["RAGTAG", "RAGTAG", "アパレル", true, "古着大手"],
  ["ラグタグ", "ラグタグ", "アパレル", true, "カタカナ表記"],
  ["古着屋JAM", "古着屋JAM", "アパレル", true, "古着チェーン"],
  ["KINJI", "KINJI", "アパレル", true, "古着チェーン"],
  ["キングファミリー", "キングファミリー", "アパレル", true, "古着チェーン"],
  ["オフハウス", "オフハウス", "大型小売", true, "ハードオフ系リユース"],
  ["ハードオフ", "ハードオフ", "大型小売", true, "大型リユース"],
  ["ブックオフ", "ブックオフ", "大型小売", true, "大型リユース"],

  // ↓Ver2.5.0で追加：買取・リサイクル系全国チェーン。
  // APPAREL_FACILITY_EXCLUDE_KEYWORDSの「買取」は店名に文字通り
  // 「買取」という語が含まれる場合しか拾えないため（例:"買取大吉"は
  // 拾えるが、"おたからや"や"ジュエルカフェ"のように屋号に「買取」の
  // 文字が入らないチェーンは素通りしてしまう）、個別にマスタ登録した。
  ["買取大吉", "買取大吉", "大型小売", true, "買取店チェーン（保険。施設除外キーワード「買取」でも拾える）"],
  ["おたからや", "おたからや", "大型小売", true, "買取店チェーン（店名に「買取」の文字がなく施設除外だけでは拾えない）"],
  ["ジュエルカフェ", "ジュエルカフェ", "大型小売", true, "買取店チェーン（店名に「買取」の文字がなく施設除外だけでは拾えない）"],

  // ドラッグストア・薬局
  ["マツモトキヨシ", "マツモトキヨシ", "大型小売", true, "ドラッグストア除外"],
  ["マツモトキヨシ", "マツキヨ", "大型小売", true, "略称表記"],
  ["ウエルシア", "ウエルシア", "大型小売", true, "ドラッグストア除外"],
  ["ウエルシア", "ウェルシア", "大型小売", true, "表記ゆれ"],
  ["ツルハドラッグ", "ツルハ", "大型小売", true, "ドラッグストア除外"],
  ["サンドラッグ", "サンドラッグ", "大型小売", true, "ドラッグストア除外"],
  ["ドラッグストアコスモス", "コスモス", "大型小売", true, "ドラッグストア除外"],
  ["スギ薬局", "スギ薬局", "大型小売", true, "ドラッグストア除外"],
  ["スギドラッグ", "スギドラッグ", "大型小売", true, "ドラッグストア除外"],
  ["ココカラファイン", "ココカラファイン", "大型小売", true, "ドラッグストア除外"],
  ["クリエイトSD", "クリエイトSD", "大型小売", true, "ドラッグストア除外"],
  ["クリエイトエス・ディー", "クリエイトエスディー", "大型小売", true, "表記ゆれ"],
  ["カワチ薬品", "カワチ薬品", "大型小売", true, "ドラッグストア除外"],
  ["クスリのアオキ", "クスリのアオキ", "大型小売", true, "ドラッグストア除外"],
  ["ドラッグセイムス", "セイムス", "大型小売", true, "ドラッグストア除外"],
  ["キリン堂", "キリン堂", "大型小売", true, "ドラッグストア除外"],
  ["ゲンキー", "ゲンキー", "大型小売", true, "ドラッグストア除外"],
  ["ヤックスドラッグ", "ヤックス", "大型小売", true, "ドラッグストア除外"],

  // ホームセンター・大型小売
  ["カインズ", "カインズ", "大型小売", true, "ホームセンター除外"],
  ["カインズ", "CAINZ", "大型小売", true, "英語表記"],
  ["コメリ", "コメリ", "大型小売", true, "ホームセンター除外"],
  ["DCM", "DCM", "大型小売", true, "ホームセンター除外"],
  ["ケーヨーデイツー", "ケーヨーデイツー", "大型小売", true, "ホームセンター除外"],
  ["コーナン", "コーナン", "大型小売", true, "ホームセンター除外"],
  ["ビバホーム", "ビバホーム", "大型小売", true, "ホームセンター除外"],
  ["ジョイフル本田", "ジョイフル本田", "大型小売", true, "ホームセンター除外"],
  ["島忠", "島忠", "大型小売", true, "ホームセンター除外"],
  ["ナフコ", "ナフコ", "大型小売", true, "ホームセンター除外"],
  ["ロイヤルホームセンター", "ロイヤルホームセンター", "大型小売", true, "ホームセンター除外"],
  ["ニトリ", "ニトリ", "大型小売", true, "家具量販除外"],
  ["ヤマダデンキ", "ヤマダ", "大型小売", true, "家電量販除外"],
  ["ケーズデンキ", "ケーズデンキ", "大型小売", true, "家電量販除外"],
  ["エディオン", "エディオン", "大型小売", true, "家電量販除外"],
  ["ダイソー", "ダイソー", "大型小売", true, "100円ショップ除外"],
  ["DAISO", "DAISO", "大型小売", true, "英語表記"],
  ["セリア", "セリア", "大型小売", true, "100円ショップ除外"],
  ["Seria", "Seria", "大型小売", true, "英語表記"],
  ["キャンドゥ", "キャンドゥ", "大型小売", true, "100円ショップ除外"],
  ["Can Do", "Can Do", "大型小売", true, "英語表記"],
  ["トライアル", "トライアル", "大型小売", true, "スーパー除外"],
  ["ベイシア", "ベイシア", "大型小売", true, "スーパー除外"],
  ["業務スーパー", "業務スーパー", "大型小売", true, "スーパー除外"]
];

// =====================================================================
// 一括処理
// =====================================================================
function apparelExecuteAllProcesses() {
  apparelSetupAll();
  apparelImportCSVFiles();
  apparelExecuteNormalizeAndValidate();
  apparelExecuteDuplicateCheck();
  apparelAddKnownChainMasterGapsSilently();
  apparelExecuteChainCheck();
  apparelExecuteFacilityCheck();
  apparelExecuteWorkflowGrouping();
  apparelExecuteSplitSheets();
  apparelExecuteGenerateSalesAreaSheets();
  apparelExecuteExportSalesAreaCsvFiles();
  const summary = apparelExecuteCountSummary();

  SpreadsheetApp.getUi().alert(
    "アパレルリスト処理が完了しました。\n\n" +
    `営業対象: ${summary.totalTarget}件\n` +
    `確認対象: ${summary.totalConfirm}件\n` +
    `除外対象: ${summary.totalExclude}件\n` +
    `取得失敗: ${summary.totalFailed}件\n\n` +
    "アパレル_04_SALES_地域別タブと完成版CSVエクスポートを確認してください。\n\n" +
    "※「チェーン店疑い」（地名付きの支店名パターンで自動検出したが、マスタには" +
    "未登録の店舗）は確認対象に入っています。本当にチェーン店か目視確認してください。"
  );
}

function apparelSetupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(APPAREL_SHEETS).forEach(key => apparelGetOrCreateSheet_(ss, APPAREL_SHEETS[key]));
  apparelCreateFolders_();
  apparelAddKnownChainMasterGapsSilently();
  ss.toast("アパレル用フォルダ・タブを確認しました。", "👕 アパレル");
}

function apparelCreateFolders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  apparelGetOrCreateFolder_(parentFolder, APPAREL_FOLDER_NAMES.input);
  apparelGetOrCreateFolder_(parentFolder, APPAREL_FOLDER_NAMES.processed);
  apparelGetOrCreateFolder_(parentFolder, APPAREL_FOLDER_NAMES.export);
}

// =====================================================================
// 1. CSV取り込み
// =====================================================================
function apparelImportCSVFiles() {
  apparelCreateFolders_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = apparelGetOrCreateSheet_(ss, APPAREL_SHEETS.normalized);
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const importFolder = apparelGetOrCreateFolder_(parentFolder, APPAREL_FOLDER_NAMES.input);
  const processedFolder = apparelGetOrCreateFolder_(parentFolder, APPAREL_FOLDER_NAMES.processed);
  const files = importFolder.getFilesByType(MimeType.CSV);

  const parsedFiles = [];
  const unifiedHeader = [];
  const unifiedHeaderSet = new Set();

  while (files.hasNext()) {
    const file = files.next();
    try {
      ss.toast(`ファイル「${file.getName()}」を解析中...`, "📁 CSV一括取り込み");
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
      Logger.log(`[アパレルCSV取込エラー] ${file.getName()}: ${e.message}`);
      ss.toast(`「${file.getName()}」の読み込みに失敗しました: ${e.message}`, "⚠️ CSV取込エラー");
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

  apparelWriteRowsToExistingSheet_(targetSheet, combinedData[0], combinedData.slice(1));
  return { files: parsedFiles.length, rows: combinedData.length - 1 };
}

// =====================================================================
// 2. 正規化・基本判定
// =====================================================================
function apparelExecuteNormalizeAndValidate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APPAREL_SHEETS.normalized);
  if (!sheet || sheet.getLastRow() <= 1) return { count: 0 };

  const values = sheet.getDataRange().getValues();
  const rawHeader = apparelNormalizeHeaderRow_(values[0]);
  const rawRows = values.slice(1);
  const outputRows = [];

  rawRows.forEach(row => {
    const rawStoreName = apparelGetValue_(rawHeader, row, ["店名", "店舗名", "名前", "施設名", "タイトル", "name", "Name"]);
    const storeName = apparelStripGenreSuffixFromName_(rawStoreName);
    const genre = apparelGetValue_(rawHeader, row, ["ジャンル", "カテゴリ", "カテゴリー", "業種", "種別", "category", "types"]);
    const searchGenre = apparelGetValue_(rawHeader, row, ["検索ジャンル", "検索キーワード", "取得元ジャンル", "keyword"]);
    const sourceGenre = apparelGetValue_(rawHeader, row, ["取得元ジャンル", "元ジャンル", "sourceGenre"]);
    const address = apparelGetValue_(rawHeader, row, ["住所", "所在地", "住所１", "住所1", "address", "Address", "formatted_address"]);
    const parsedAddress = apparelParsePrefCityFromAddress_(address);
    const pref = apparelGetValue_(rawHeader, row, ["都道府県", "pref", "Prefecture"]) || parsedAddress.pref;
    const city = apparelGetValue_(rawHeader, row, ["市区町村", "市町村", "city", "City"]) || parsedAddress.city;
    const zip = apparelParseAddressDetails_(address).pcode;
    const phone = apparelNormalizePhoneDisplay_(apparelGetValue_(rawHeader, row, ["電話番号", "TEL", "Tel", "tel", "電話", "phone", "Phone", "Tel1"]));
    const url = apparelGetValue_(rawHeader, row, ["URL", "Webサイト", "ホームページ", "サイト", "website", "Website", "リンク"]);
    const media = apparelGetValue_(rawHeader, row, ["媒体", "取得元", "source", "Source"]) || "Googleマップ";
    const hpHave = apparelGetValue_(rawHeader, row, ["HP有無", "HPある？", "ホームページ有無"]);
    const businessDays = apparelGetValue_(rawHeader, row, ["営業日", "営業曜日"]);
    const holiday = apparelGetValue_(rawHeader, row, ["定休日", "休業曜日"]);
    const openA = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetValue_(rawHeader, row, ["営業開始A", "営業開始", "午前始"])));
    const closeA = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetValue_(rawHeader, row, ["営業終了A", "営業終了", "午前終"])));
    const openB = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetValue_(rawHeader, row, ["営業開始B", "午後始"])));
    const closeB = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetValue_(rawHeader, row, ["営業終了B", "午後終"])));
    const fetchStatus = apparelGetValue_(rawHeader, row, ["取得ステータス", "status"]);
    const externalReason = apparelGetValue_(rawHeader, row, ["除外理由", "理由"]);

    const normalizedGenre = apparelNormalizeSystemGenre_(genre, searchGenre, sourceGenre, storeName);
    const normalizedPhone = apparelNormalizePhoneNumberForAnalysis_(phone);
    const normalizedName = apparelSimplifyStoreName_(storeName);
    const addressStatus = apparelJudgeAddressStatus_(address, pref, city);
    const areaStatus = apparelJudgeAreaStatus_(address, pref, city);

    const basicReasons = [];
    if (!storeName) basicReasons.push("店名なし");
    if (!normalizedPhone) basicReasons.push("電話番号なし");
    if (addressStatus.status !== "住所あり") basicReasons.push(addressStatus.reason);
    if (!normalizedGenre || !apparelIsValidTargetGenre_(normalizedGenre)) basicReasons.push("ジャンル確認");
    if (fetchStatus === "失敗") basicReasons.push("取得失敗");

    outputRows.push([
      storeName,
      normalizedGenre || genre,
      searchGenre,
      sourceGenre,
      pref,
      city,
      zip,
      address,
      phone,
      url,
      media,
      apparelNormalizeHpStatus_(hpHave, url),
      businessDays,
      holiday,
      openA,
      closeA,
      openB,
      closeB,
      fetchStatus,
      externalReason,
      normalizedPhone,
      normalizedName,
      normalizedGenre,
      addressStatus.status,
      areaStatus.status,
      areaStatus.reason,
      basicReasons.length === 0 ? "対象" : "確認対象",
      apparelUniqueTextList_(basicReasons).join(" / ")
    ]);
  });

  apparelWriteRowsToExistingSheet_(sheet, APPAREL_NORMALIZED_HEADER, outputRows);
  return { count: outputRows.length };
}

// =====================================================================
// 3. 重複判定
// =====================================================================
function apparelExecuteDuplicateCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(APPAREL_SHEETS.normalized);
  const targetSheet = apparelGetOrCreateSheet_(ss, APPAREL_SHEETS.duplicate);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = apparelNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const seenPhones = new Set();
  const seenNames = new Set();
  const outputRows = [];

  rows.forEach(row => {
    const rawName = apparelGetRowValueByHeader_(header, row, "店名");
    const rawPhone = apparelNormalizePhoneNumberForAnalysis_(apparelGetRowValueByHeader_(header, row, "電話番号")).replace(/[^\d]/g, "");
    const cleanName = apparelSimplifyStoreName_(rawName);
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

  apparelWriteRowsToExistingSheet_(targetSheet, header.concat(["重複判定"]), outputRows);
  return { count: outputRows.length };
}

// =====================================================================
// 4. チェーン判定
// =====================================================================
// チェーン判定は3段階：
//   "チェーン店"      … MASTER_CHAINへの明示登録、または号店・駅前店・
//                       モール系テナント名などの強いシグナルに一致（確実）
//   "チェーン店疑い"  … 地名＋「店」、スペース/括弧区切りの「〇〇店」など、
//                       チェーンらしい命名パターンではあるが確証がない
//                       （単独店が同じような命名をしている可能性を否定できない）
//   "単独店"          … いずれにも一致しない
// "チェーン店"は従来通り即除外、"チェーン店疑い"は確認対象に回して
// 人の目で最終確認してもらう（誤って独立系の小規模店舗を営業リストから
// 永久に落とすリスクを避けるため）。
function apparelExecuteChainCheck() {
  apparelAddKnownChainMasterGapsSilently();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(APPAREL_SHEETS.masterChain);
  const sourceSheet = ss.getSheetByName(APPAREL_SHEETS.duplicate);
  const targetSheet = apparelGetOrCreateSheet_(ss, APPAREL_SHEETS.chain);
  if (!masterSheet || !sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const masterValues = masterSheet.getDataRange().getValues();
  const chainMaster = [];
  for (let i = 1; i < masterValues.length; i++) {
    const chainName = apparelTextValue_(masterValues[i][0]);
    const keyword = apparelTextValue_(masterValues[i][1]);
    const isValid = masterValues[i][3];
    if (isValid && keyword) {
      chainMaster.push({
        chainName,
        keyword,
        normalizedKeyword: apparelNormalizeForKeywordMatch_(keyword)
      });
    }
  }

  const sourceValues = sourceSheet.getDataRange().getValues();
  const header = apparelNormalizeHeaderRow_(sourceValues[0]);
  const rows = sourceValues.slice(1);
  const outputRows = [];

  rows.forEach(row => {
    const storeName = apparelGetRowValueByHeader_(header, row, "店名");
    const genre = apparelGetRowValueByHeader_(header, row, "ジャンル");
    const searchGenre = apparelGetRowValueByHeader_(header, row, "検索ジャンル");
    const sourceGenre = apparelGetRowValueByHeader_(header, row, "取得元ジャンル");
    const haystackRaw = [storeName, genre, searchGenre, sourceGenre].join(" ");
    const haystack = apparelNormalizeForKeywordMatch_(haystackRaw);

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
      const branchLevel = apparelIsLikelyBranchStoreName_(storeName);

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

  apparelWriteRowsToExistingSheet_(targetSheet, header.concat(["チェーン判定", "チェーン名", "チェーン理由"]), outputRows);
  return { count: outputRows.length };
}

function apparelAddKnownChainMasterGapsSilently() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = ss.getSheetByName(APPAREL_SHEETS.masterChain);
  if (!masterSheet) {
    masterSheet = apparelGetOrCreateSheet_(ss, APPAREL_SHEETS.masterChain);
    masterSheet.getRange(1, 1, 1, 5).setValues([["チェーン名", "キーワード", "業種", "有効", "メモ"]]);
    masterSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }

  const existingValues = masterSheet.getDataRange().getValues();
  const existingKeywords = new Set(existingValues.slice(1).map(r => apparelTextValue_(r[1]).normalize("NFC")));
  const newRows = APPAREL_KNOWN_CHAIN_GAPS.filter(row => {
    const keyword = apparelTextValue_(row[1]).normalize("NFC");
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

function apparelFixKnownChainMasterGaps() {
  const result = apparelAddKnownChainMasterGapsSilently();
  SpreadsheetApp.getUi().alert(
    result.added.length === 0
      ? "追加対象がありません（すべて登録済みです）。"
      : `【完了】${result.added.length}件のキーワードをアパレル_MASTER_CHAINへ追加しました。\n\n反映するには「4. チェーン判定」以降を再実行してください。`
  );
}

// =====================================================================
// 5. 施設判定
// =====================================================================
function apparelExecuteFacilityCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(APPAREL_SHEETS.chain) || ss.getSheetByName(APPAREL_SHEETS.duplicate) || ss.getSheetByName(APPAREL_SHEETS.normalized);
  const targetSheet = apparelGetOrCreateSheet_(ss, APPAREL_SHEETS.facility);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = apparelNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const outputRows = [];

  rows.forEach(row => {
    const facility = apparelJudgeFacilityStatus_(
      apparelGetRowValueByHeader_(header, row, "店名"),
      apparelGetRowValueByHeader_(header, row, "住所"),
      apparelGetRowValueByHeader_(header, row, "ジャンル"),
      apparelGetRowValueByHeader_(header, row, "検索ジャンル"),
      apparelGetRowValueByHeader_(header, row, "取得元ジャンル")
    );
    const sales = apparelJudgeSalesTargetStatus_(header, row, facility);
    outputRows.push(row.concat([
      facility.status,
      facility.reason,
      sales.status,
      sales.reason
    ]));
  });

  apparelWriteRowsToExistingSheet_(targetSheet, header.concat(["施設判定", "施設判定理由", "営業対象判定", "営業対象除外理由"]), outputRows);
  return { count: outputRows.length };
}

function apparelJudgeFacilityStatus_(storeName, address, genre, searchGenre, sourceGenre) {
  const haystack = apparelNormalizeForKeywordMatch_([storeName, address, genre, searchGenre, sourceGenre].join(" "));
  const excludeKeyword = APPAREL_FACILITY_EXCLUDE_KEYWORDS.find(keyword => haystack.indexOf(apparelNormalizeForKeywordMatch_(keyword)) !== -1);
  if (excludeKeyword) return { status: "除外", reason: "完全除外キーワード一致: " + excludeKeyword };

  const reviewKeyword = APPAREL_FACILITY_REVIEW_KEYWORDS.find(keyword => haystack.indexOf(apparelNormalizeForKeywordMatch_(keyword)) !== -1);
  if (reviewKeyword) return { status: "確認対象", reason: "確認対象キーワード一致: " + reviewKeyword };

  return { status: "対象", reason: "" };
}

function apparelJudgeSalesTargetStatus_(header, row, facility) {
  const reasons = [];
  let hasReview = false;
  let hasExclude = false;

  const fetchStatus = apparelGetRowValueByHeader_(header, row, "取得ステータス");
  const externalReason = apparelGetRowValueByHeader_(header, row, "除外理由");
  const basicStatus = apparelGetRowValueByHeader_(header, row, "基本データ判定");
  const basicReason = apparelGetRowValueByHeader_(header, row, "基本データ除外理由");
  const areaStatus = apparelGetRowValueByHeader_(header, row, "エリア判定");
  const areaReason = apparelGetRowValueByHeader_(header, row, "エリア判定理由");
  const dupStatus = apparelGetRowValueByHeader_(header, row, "重複判定");
  const chainStatus = apparelGetRowValueByHeader_(header, row, "チェーン判定");

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

  const joined = apparelUniqueTextList_(reasons).join(" / ");
  if (!joined) return { status: "対象", reason: "" };
  if (hasReview && !hasExclude) return { status: "確認対象", reason: joined };
  return { status: "除外", reason: joined };
}

// =====================================================================
// 6. ワークフロー分類
// =====================================================================
function apparelExecuteWorkflowGrouping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APPAREL_SHEETS.facility);
  if (!sheet || sheet.getLastRow() <= 1) return { count: 0 };

  const values = sheet.getDataRange().getValues();
  const header = apparelNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const outputRows = [];

  rows.forEach(row => {
    const workflow = apparelJudgeWorkflowGroup_(header, row);
    outputRows.push(row.concat([
      workflow.group,
      workflow.item,
      workflow.status,
      workflow.nextAction
    ]));
  });

  apparelWriteRowsToExistingSheet_(sheet, header.concat(["ワークフローグループ", "ワークフロー項目", "対応ステータス", "次アクション"]), outputRows);
  return { count: outputRows.length };
}

function apparelJudgeWorkflowGroup_(header, row) {
  const storeName = apparelGetRowValueByHeader_(header, row, "店名");
  const address = apparelGetRowValueByHeader_(header, row, "住所");
  const url = apparelGetRowValueByHeader_(header, row, "URL");
  const fetchStatus = apparelGetRowValueByHeader_(header, row, "取得ステータス");
  const dupStatus = apparelGetRowValueByHeader_(header, row, "重複判定");
  const chainStatus = apparelGetRowValueByHeader_(header, row, "チェーン判定");
  const facilityStatus = apparelGetRowValueByHeader_(header, row, "施設判定");
  const areaStatus = apparelGetRowValueByHeader_(header, row, "エリア判定");
  const addressStatus = apparelGetRowValueByHeader_(header, row, "住所判定");
  const normalizedPhone = apparelGetRowValueByHeader_(header, row, "正規化電話番号");
  const normalizedGenre = apparelGetRowValueByHeader_(header, row, "正規化ジャンル");

  if (fetchStatus === "失敗" || (!storeName && url)) {
    return { group: APPAREL_SHEETS.failed, item: "詳細取得失敗", status: "未対応", nextAction: "再取得" };
  }
  if (dupStatus && dupStatus !== "ユニーク") return { group: APPAREL_SHEETS.exclude, item: "重複除外", status: "除外確定", nextAction: "投入しない" };
  if (chainStatus === "チェーン店") return { group: APPAREL_SHEETS.exclude, item: "チェーン店除外", status: "除外確定", nextAction: "投入しない" };
  if (facilityStatus === "除外") return { group: APPAREL_SHEETS.exclude, item: "大型小売・モール除外", status: "除外確定", nextAction: "投入しない" };
  if (areaStatus === "エリア外") return { group: APPAREL_SHEETS.exclude, item: "エリア外除外", status: "除外確定", nextAction: "投入しない" };
  if (!storeName || (!address && !url) || (addressStatus === "住所未取得" && !url)) {
    return { group: APPAREL_SHEETS.exclude, item: "住所未取得除外", status: "除外確定", nextAction: "投入しない" };
  }
  // チェーン店疑い（マスタ未登録だが地名付き支店名パターンで検出）は
  // 即除外にせず確認対象へ回し、人の目で最終判断してもらう。
  if (chainStatus === "チェーン店疑い") {
    return { group: APPAREL_SHEETS.confirm, item: "チェーン店疑い確認", status: "未対応", nextAction: "支店名パターンで検出。本当にチェーン店か目視確認してください" };
  }
  if (facilityStatus === "確認対象") return { group: APPAREL_SHEETS.confirm, item: "小規模ビル確認", status: "未対応", nextAction: "テナントか路面店か確認" };
  if (!normalizedPhone && storeName && address) return { group: APPAREL_SHEETS.confirm, item: "電話番号なし確認", status: "未対応", nextAction: "電話番号補完" };
  if (!normalizedGenre || !apparelIsValidTargetGenre_(normalizedGenre)) return { group: APPAREL_SHEETS.confirm, item: "ジャンル確認", status: "未対応", nextAction: "ジャンルを目視確認" };
  if (addressStatus !== "住所あり") return { group: APPAREL_SHEETS.confirm, item: "住所確認", status: "未対応", nextAction: "住所確認" };
  if (areaStatus === "判定不可") return { group: APPAREL_SHEETS.confirm, item: "住所確認", status: "未対応", nextAction: "住所確認" };

  if (apparelIsComdeskTargetRow_(header, row)) return { group: APPAREL_SHEETS.target, item: "営業対象", status: "未対応", nextAction: "コムデスク投入" };
  return { group: APPAREL_SHEETS.confirm, item: "住所確認", status: "未対応", nextAction: "目視確認" };
}

// =====================================================================
// 7. タブ分け
// =====================================================================
function apparelExecuteSplitSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(APPAREL_SHEETS.facility);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = apparelNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const buckets = {};
  buckets[APPAREL_SHEETS.target] = [];
  buckets[APPAREL_SHEETS.confirm] = [];
  buckets[APPAREL_SHEETS.exclude] = [];
  buckets[APPAREL_SHEETS.failed] = [];

  rows.forEach(row => {
    const group = apparelGetRowValueByHeader_(header, row, "ワークフローグループ") || APPAREL_SHEETS.confirm;
    if (buckets[group]) buckets[group].push(row);
    else buckets[APPAREL_SHEETS.confirm].push(row);
  });

  Object.keys(buckets).forEach(sheetName => {
    apparelWriteRowsToSheetByName_(ss, sheetName, header, buckets[sheetName]);
  });

  return { count: rows.length };
}

// =====================================================================
// 8. 04_SALES地域別タブ生成
// =====================================================================
function apparelExecuteGenerateSalesAreaSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(APPAREL_SHEETS.target);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) return { count: 0 };

  const values = sourceSheet.getDataRange().getValues();
  const header = apparelNormalizeHeaderRow_(values[0]);
  const rows = values.slice(1);
  const areaContainers = {};

  rows.forEach(row => {
    if (!apparelIsComdeskTargetRow_(header, row)) return;
    const area = apparelGetAreaKeyForRow_(header, row);
    if (!areaContainers[area]) areaContainers[area] = [];
    areaContainers[area].push(apparelBuildComdeskRow_(header, row));
  });

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("アパレル_04_SALES_") && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });

  Object.keys(areaContainers).forEach(area => {
    const sheetName = "アパレル_04_SALES_" + apparelSafeSheetNamePart_(area);
    apparelWriteRowsToExistingSheet_(apparelGetOrCreateSheet_(ss, sheetName), APPAREL_COMDESK_HEADER, areaContainers[area]);
  });

  return { count: Object.keys(areaContainers).length };
}

function apparelGetAreaKeyForRow_(header, row) {
  const city = apparelGetRowValueByHeader_(header, row, "市区町村");
  if (city) return city;
  const pref = apparelGetRowValueByHeader_(header, row, "都道府県");
  if (pref) return pref;
  return "エリア不明";
}

function apparelExecuteGenerateSalesGenreSheets() {
  return apparelExecuteGenerateSalesAreaSheets();
}

// =====================================================================
// 9. CSV出力
// =====================================================================
function apparelExecuteExportSalesAreaCsvFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const exportFolder = apparelGetOrCreateFolder_(parentFolder, APPAREL_FOLDER_NAMES.export);
  const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
  let exported = 0;

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (!sheetName.startsWith("アパレル_04_SALES_")) return;
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return;

    const areaName = sheetName.replace("アパレル_04_SALES_", "") || "ダウンロードリスト";
    const fileName = `【アパレル営業リスト】${areaName}_${formattedDate}.csv`;
    const existingFiles = exportFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

    const blob = Utilities.newBlob("﻿" + apparelConvertArrayToCsvText_(values), "text/csv", fileName);
    exportFolder.createFile(blob);
    exported++;
  });

  ss.toast(`CSV出力完了: ${exported}ファイル`, "📂 CSV出力");
  return { exported };
}

function apparelExecuteExportSalesGenreCsvFiles() {
  return apparelExecuteExportSalesAreaCsvFiles();
}

// =====================================================================
// 10. 件数サマリー
// =====================================================================
function apparelExecuteCountSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const countRows = sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    return sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;
  };

  const totalNormalized = countRows(APPAREL_SHEETS.normalized);
  const totalTarget = countRows(APPAREL_SHEETS.target);
  const totalConfirm = countRows(APPAREL_SHEETS.confirm);
  const totalExclude = countRows(APPAREL_SHEETS.exclude);
  const totalFailed = countRows(APPAREL_SHEETS.failed);

  const areaRows = [];
  let totalSales = 0;
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith("アパレル_04_SALES_")) return;
    const count = Math.max(sheet.getLastRow() - 1, 0);
    if (count <= 0) return;
    const area = name.replace("アパレル_04_SALES_", "");
    areaRows.push([name, count]);
    totalSales += count;
  });
  areaRows.sort((a, b) => b[1] - a[1]);

  const output = [
    ["区分", "件数"],
    ["── 全体 ──", ""],
    ["アパレル_01_NORMALIZED（取込総数）", totalNormalized],
    [APPAREL_SHEETS.target, totalTarget],
    [APPAREL_SHEETS.confirm, totalConfirm],
    [APPAREL_SHEETS.exclude, totalExclude],
    [APPAREL_SHEETS.failed, totalFailed],
    ["", ""],
    ["── 04_SALES_地域別 ──", ""]
  ].concat(areaRows.length ? areaRows : [["（地域別タブなし）", 0]])
    .concat([
      ["", ""],
      ["地域別合計（01_営業対象と一致するはず）", totalSales],
      ["更新日時", Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss")]
    ]);

  const summarySheet = apparelGetOrCreateSheet_(ss, APPAREL_SHEETS.summary);
  apparelWriteRowsToExistingSheet_(summarySheet, output[0], output.slice(1));
  summarySheet.getRange(1, 1, 1, 2).setFontWeight("bold");
  summarySheet.autoResizeColumns(1, 2);
  ss.setActiveSheet(summarySheet);
  ss.moveActiveSheet(1);

  return { totalNormalized, totalTarget, totalConfirm, totalExclude, totalFailed, totalSales };
}

// =====================================================================
// 判定・正規化ヘルパー
// =====================================================================
function apparelNormalizeSystemGenre_(genre, searchGenre, sourceGenre, storeName) {
  const rawGenre = apparelTextValue_(genre);
  const searchGenreText = apparelTextValue_(searchGenre);
  const sourceGenreText = apparelTextValue_(sourceGenre);
  const nameText = apparelTextValue_(storeName);

  if (apparelIsValidTargetGenre_(rawGenre)) return rawGenre;

  const candidates = [searchGenreText, rawGenre, sourceGenreText];

  for (const candidate of candidates) {
    if (candidate && APPAREL_GENRE_MAP[candidate]) return APPAREL_GENRE_MAP[candidate];
  }

  for (const candidate of candidates) {
    if (!candidate) continue;
    const matched = Object.keys(APPAREL_GENRE_MAP).find(oldGenre => candidate.indexOf(oldGenre) !== -1);
    if (matched) return APPAREL_GENRE_MAP[matched];
  }

  if (nameText) {
    const nameKana = apparelHiraganaToKatakana_(nameText);
    const matchedTarget = APPAREL_TARGET_GENRES.find(g => nameKana.indexOf(apparelHiraganaToKatakana_(g)) !== -1);
    if (matchedTarget) return matchedTarget;

    const matchedOld = Object.keys(APPAREL_GENRE_MAP).find(oldGenre => nameKana.indexOf(apparelHiraganaToKatakana_(oldGenre)) !== -1);
    if (matchedOld) return APPAREL_GENRE_MAP[matchedOld];
  }

  return rawGenre;
}

function apparelJudgeAreaStatus_(address, pref, city) {
  const cleanAddress = apparelNormalizeAddressText_(address);
  const prefText = apparelTextValue_(pref);
  const cityText = apparelTextValue_(city);
  if (!cleanAddress) return { status: "判定不可", reason: "住所未取得" };
  if (!prefText || !cityText) return { status: "判定不可", reason: "都道府県または市区町村が空" };
  if (cleanAddress.indexOf(prefText) === -1 && cleanAddress.indexOf(cityText) === -1) {
    return { status: "判定不可", reason: "住所に都道府県/市区町村が含まれない" };
  }
  return { status: "エリア内", reason: "" };
}

function apparelJudgeAddressStatus_(address, pref, city) {
  const cleanAddress = apparelNormalizeAddressText_(address);
  if (!cleanAddress) return { status: "住所未取得", reason: "住所未取得" };
  if (!apparelTextValue_(pref) || !apparelTextValue_(city)) return { status: "住所未取得", reason: "都道府県または市区町村が空" };
  return { status: "住所あり", reason: "" };
}

function apparelIsComdeskTargetRow_(header, row) {
  const phone = apparelGetRowValueByHeader_(header, row, "正規化電話番号") || apparelNormalizePhoneNumberForAnalysis_(apparelGetRowValueByHeader_(header, row, "電話番号"));
  const genre = apparelGetRowValueByHeader_(header, row, "正規化ジャンル") || apparelGetRowValueByHeader_(header, row, "ジャンル");
  return apparelGetRowValueByHeader_(header, row, "営業対象判定") === "対象" &&
    apparelGetRowValueByHeader_(header, row, "重複判定") === "ユニーク" &&
    // "チェーン店"も"チェーン店疑い"も除外し、"単独店"のみ通す
    // （厳密に単独店であることを確認した行だけを営業対象にする）
    apparelGetRowValueByHeader_(header, row, "チェーン判定") === "単独店" &&
    apparelGetRowValueByHeader_(header, row, "施設判定") === "対象" &&
    apparelGetRowValueByHeader_(header, row, "エリア判定") === "エリア内" &&
    !!phone &&
    !!apparelGetRowValueByHeader_(header, row, "住所") &&
    apparelIsValidTargetGenre_(genre);
}

function apparelIsValidTargetGenre_(genre) {
  return APPAREL_TARGET_GENRES.indexOf(apparelTextValue_(genre)) !== -1;
}

// 支店名らしさを判定する。
// 返り値: "" (支店名パターンではない) / "high" (号店・駅前店・モール系
// テナント名など、ほぼ確実にチェーンの支店と分かる強いシグナル) /
// "heuristic" (スペース・括弧区切りの「〇〇店」、地名＋「店」など、
// チェーンの可能性が高いが単独店が偶然同じ命名をしている可能性も
//否定できないシグナル)。
// "high"は即除外、"heuristic"は確認対象に回す、という運用にしている
// （誤って独立系の小規模店舗を営業リストから永久に落としてしまうと、
// その店舗には二度と営業をかけられなくなり機会損失が大きいため）。
function apparelIsLikelyBranchStoreName_(storeName) {
  const name = apparelTextValue_(storeName).normalize("NFKC").trim();
  if (!name) return "";

  // ①単独店・個人店にありがちな「〜店」は保護（誤除外防止）
  const personalLikePatterns = [
    /洋品店$/, /衣料品店$/, /呉服店$/, /古着店$/, /下着店$/, /靴店$/, /洋服店$/,
    /専門店$/, /セレクトショップ$/, /ブティック$/, /商店$/, /本店$/, /販売店$/
  ];
  if (personalLikePatterns.some(re => re.test(name))) return "";

  // ②明らかな商業施設・モールのテナントを示す接尾辞（強いシグナル）
  if (/(モール店|イオン店|アウトレット店|SC店|ショッピングセンター店|スクエア店|パーク店)$/.test(name)) return "high";

  // ③号店・駅・インター等、チェーン店の支店名として非常に一般的な接尾辞（強いシグナル）
  if (/[0-9０-９]+号店$/.test(name)) return "high";
  if (/(駅前|駅|北口|南口|東口|西口|インター|バイパス).{0,10}店$/.test(name)) return "high";

  // ④スペースや括弧で区切られた「〇〇店」（例:"店舗名 古河店","店舗名(結城店)"）
  //   → 確証はないため"heuristic"（確認対象）
  if (/[\s\(（][^\s\(（]+店[\)）]?$/.test(name)) return "heuristic";

  // ⑤地名（都道府県・市区町村）＋「店」（例:"〇〇下妻市店","〇〇日立市店"）
  //   → 確証はないため"heuristic"（確認対象）
  if (/.+(都|道|府|県|市|区|町|村).{0,10}店$/.test(name)) return "heuristic";

  // ⑥「下妻店」「古河店」のように、行政区画文字（市区町村等）を含まない
  //   短い地名・固有名詞＋「店」で終わるケース（実務上、実際のCSVに現れる
  //   支店名の多くはこの形式）。①の保護パターン（洋品店・専門店・商店等）に
  //   一致するものは既に上で除外済みなので、ここに来るのは基本的に
  //   「短い固有名詞＋店」のみ。確証はないため"heuristic"（確認対象）。
  if (/^.{1,6}店$/.test(name)) return "heuristic";

  return "";
}

function apparelStripGenreSuffixFromName_(name) {
  const raw = apparelTextValue_(name);
  const match = raw.match(/^(.*?)\s*[（(]([^）)]{1,20})[）)]\s*$/);
  if (!match) return raw;
  const inner = apparelTextValue_(match[2]);
  const knownGenreWords = APPAREL_TARGET_GENRES.concat(Object.keys(APPAREL_GENRE_MAP));
  if (knownGenreWords.indexOf(inner) !== -1) return apparelTextValue_(match[1]);
  return raw;
}

function apparelNormalizeHpStatus_(hp, url) {
  const value = apparelTextValue_(hp).toLowerCase();
  if (value === "1" || value === "true" || value.indexOf("有") !== -1) return "1";
  if (url) return "1";
  return "0";
}

// =====================================================================
// コムデスク行生成
// =====================================================================
function apparelBuildComdeskRow_(header, row) {
  const storeName = apparelGetRowValueByHeader_(header, row, "店名");
  const fullAddr = apparelGetRowValueByHeader_(header, row, "住所");
  const addrDetails = apparelParseAddressDetails_(fullAddr);
  const pref = apparelGetRowValueByHeader_(header, row, "都道府県") || addrDetails.pref;
  const city = apparelGetRowValueByHeader_(header, row, "市区町村");
  const phone = apparelGetRowValueByHeader_(header, row, "電話番号");
  const cleanPhone = apparelNormalizePhoneNumberForAnalysis_(phone).replace(/[^\d]/g, "");
  const media = apparelGetRowValueByHeader_(header, row, "媒体") || "Googleマップ";
  const url = apparelGetRowValueByHeader_(header, row, "URL");
  const hpHave = apparelGetRowValueByHeader_(header, row, "HP有無");
  const hpStatus = (hpHave.indexOf("有") !== -1 || hpHave === "1" || hpHave.toLowerCase() === "true") ? "1" : "0";
  const bizDaysVal = apparelRemoveHolidayFromBizDays_(apparelGetRowValueByHeader_(header, row, "営業日"), apparelGetRowValueByHeader_(header, row, "定休日"));
  const holidayVal = apparelGetRowValueByHeader_(header, row, "定休日");
  const rawOpenA = apparelGetRowValueByHeader_(header, row, "営業開始A") || apparelGetRowValueByHeader_(header, row, "営業開始");
  const openAVal = apparelFormatToPureTime_(apparelToHalfWidthForTime_(rawOpenA));
  const closeAVal = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetRowValueByHeader_(header, row, "営業終了A") || apparelGetRowValueByHeader_(header, row, "営業終了")));
  const openBVal = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetRowValueByHeader_(header, row, "営業開始B")));
  const closeBVal = apparelFormatToPureTime_(apparelToHalfWidthForTime_(apparelGetRowValueByHeader_(header, row, "営業終了B")));
  const timeValues = apparelNormalizeBusinessTimeValues_(rawOpenA, openAVal, closeAVal, openBVal, closeBVal);

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
  salesRow[22] = `${areaText}tel${cleanPhone}`;
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
function apparelGetValue_(header, row, names) {
  for (const name of names) {
    const idx = header.indexOf(name);
    if (idx !== -1 && row[idx] !== undefined && row[idx] !== null && apparelTextValue_(row[idx]) !== "") {
      return apparelTextValue_(row[idx]);
    }
  }
  return "";
}

function apparelConvertArrayToCsvText_(array) {
  return array.map(row => row.map(cell => {
    const str = String(cell === null || cell === undefined ? "" : cell).replace(/"/g, '""');
    if (str.includes(",") || str.includes("\n") || str.includes("\r") || str.includes('"')) return '"' + str + '"';
    return str;
  }).join(",")).join("\r\n");
}

function apparelSimplifyStoreName_(name) {
  if (!name) return "";
  let n = String(name).normalize("NFKC").toLowerCase();
  n = n.replace(/[ぁ-ゖ]/g, m => String.fromCharCode(m.charCodeAt(0) + 0x60));
  n = n.replace(/[\s ・、。，．・！？!?()（）【】\[\]「」『』_－\-〜~]/g, "");
  n = n.replace(/(店|駅前店|北口店|南口店|東口店|西口店|インター店|本店|支店|営業所)$/, "");
  return n.trim();
}

function apparelNormalizeForKeywordMatch_(text) {
  return apparelTextValue_(text)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\s　・．.、,＆&()（）【】\[\]「」『』_－\-ー]/g, "");
}

function apparelHiraganaToKatakana_(text) {
  return apparelTextValue_(text).replace(/[ぁ-ゖ]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
}

function apparelGetOrCreateFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function apparelGetOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function apparelNormalizeHeaderRow_(header) {
  return header.map(h => String(h || "").replace(/^﻿/, "").trim());
}

function apparelGetRowValueByHeader_(header, row, name) {
  const idx = header.indexOf(name);
  return idx === -1 ? "" : apparelTextValue_(row[idx]);
}

function apparelWriteRowsToExistingSheet_(sheet, header, rows) {
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

function apparelWriteRowsToSheetByName_(ss, sheetName, header, rows) {
  const sheet = apparelGetOrCreateSheet_(ss, sheetName);
  apparelWriteRowsToExistingSheet_(sheet, header, rows);
}

function apparelNormalizePhoneDisplay_(phone) {
  return apparelTextValue_(phone)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[ー−―]/g, "-")
    .replace(/[^0-9\-]/g, "")
    .trim();
}

function apparelNormalizePhoneNumberForAnalysis_(phone) {
  const digits = apparelTextValue_(phone)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[^\d]/g, "");
  if (!digits) return "";
  if (digits.length === 11) return digits.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3");
  if (digits.startsWith("03") || digits.startsWith("06")) return digits.replace(/(\d{2})(\d{4})(\d{4})/, "$1-$2-$3");
  if (digits.length === 10) return digits.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
  return digits;
}

function apparelNormalizeAddressText_(address) {
  return apparelTextValue_(address)
    .replace(/(?:〒\d{3}-?\d{4}\s*|日本、\s*|日本\s*)/g, "")
    .replace(/\s+/g, "");
}

function apparelParsePrefCityFromAddress_(address) {
  const cleanAddress = apparelNormalizeAddressText_(address);
  const match = cleanAddress.match(/^((?:北海道|東京都|大阪府|京都府|.{2,3}県))?((?:.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村]))?/);
  return {
    pref: match && match[1] ? match[1] : "",
    city: match && match[2] ? match[2] : ""
  };
}

function apparelParseAddressDetails_(fullAddress) {
  let pcode = "";
  let pref = "";
  let addr1 = apparelTextValue_(fullAddress);
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

function apparelToHalfWidthForTime_(str) {
  return apparelTextValue_(str).replace(/[０-９：]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
}

function apparelNormalizeBusinessTimeValues_(rawOpenA, openAVal, closeAVal, openBVal, closeBVal) {
  const result = { openA: openAVal, closeA: closeAVal, openB: openBVal, closeB: closeBVal };
  if (!openBVal && !closeBVal) {
    const timeMatches = apparelToHalfWidthForTime_(rawOpenA).match(/(\d{1,2}:\d{2})/g);
    if (timeMatches && timeMatches.length >= 2) {
      result.openA = apparelFormatToPureTime_(timeMatches[0]);
      result.closeA = "";
      result.openB = "";
      result.closeB = apparelFormatToPureTime_(timeMatches[timeMatches.length - 1]);
    } else if (openAVal) {
      result.closeA = "";
      result.openB = "";
      result.closeB = closeAVal;
    }
  }
  return result;
}

function apparelRemoveHolidayFromBizDays_(bizDays, holiday) {
  let bizDaysVal = apparelTextValue_(bizDays);
  const holidayVal = apparelTextValue_(holiday);
  if (!bizDaysVal || !holidayVal) return bizDaysVal;
  ["月", "火", "水", "木", "金", "土", "日", "祝"].forEach(day => {
    if (holidayVal.indexOf(day) !== -1) {
      const regex = new RegExp(day + "[・、/]?|[・、/]?" + day, "g");
      bizDaysVal = bizDaysVal.replace(regex, "");
    }
  });
  return bizDaysVal.replace(/^[・、/]+|[・、/]+$/g, "").replace(/[・、/]{2,}/g, "・");
}

function apparelFormatToPureTime_(val) {
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

function apparelSafeSheetNamePart_(name) {
  return apparelTextValue_(name)
    .replace(/[\\\/\?\*\[\]\:]/g, "")
    .replace(/_/g, "")
    .substring(0, 30) || "エリア不明";
}

function apparelUniqueTextList_(values) {
  const seen = {};
  return values.map(apparelTextValue_).filter(value => {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function apparelTextValue_(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}
