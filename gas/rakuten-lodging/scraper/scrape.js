#!/usr/bin/env node
/**
 * 楽天トラベル 都道府県別 宿泊施設スクレイパー Ver1.0.0
 *
 * 目的：
 * これまでOctoparseで手動実行していた「都道府県を指定して宿泊施設一覧＋
 * 住所・電話番号・部屋数を取得する」作業を、都道府県名だけ指定すれば
 * ノーコードツールを使わずに完結するようにしたCLI。
 * 出力するCSVは、../code-rakuten.js（GASの正規化・分類ロジック）が
 * すでに認識している列名（店名・住所・電話番号・FAX・総部屋数・都道府県 等）
 * でそのまま出力するため、Octoparseの「フィールドN」形式の変換は不要。
 * 出力先を「楽天_CSV投入フォルダ」（Google Driveと同期しているローカルフォルダ）
 * に指定すれば、GAS側のトリガーがそのまま拾って処理する。
 *
 * 使い方：
 *   npm install
 *   node scrape.js --pref 茨城県 --out ./output
 *
 * オプション：
 *   --pref <都道府県名>   必須。例: 茨城県
 *   --out <ディレクトリ>  出力先。省略時は ./output
 *   --hyoji <件数>        1ページあたりの取得件数（既定30。楽天側の仕様に合わせる）
 *   --delay <ms>          リクエスト間隔（既定500ms。サイトへの配慮のため）
 *   --concurrency <数>    詳細ページ取得の同時実行数（既定3）
 *   --max <件数>          デバッグ用。取得件数の上限（既定は無制限）
 *
 * ★重要：この一次実装について
 * 一覧ページ（web.travel.rakuten.co.jp/portal/my/search_undecided.main）と
 * 詳細ページ（travel.rakuten.co.jp/HOTEL/{id}/{id}_std.html）の実データは
 * 茨城県で1回ずつ内容を確認した上でこのパーサーを書いているが、
 * 実際のHTMLタグ構造（class名など）までは確認できていない
 * （確認に使ったツールがHTML→テキスト変換して返すものだったため）。
 * そのため、住所・電話番号・部屋数などの抽出は「◯◯というラベル文字列の
 * 直後にある値」という正規表現ベースの抽出にしてあり、CSS/DOMセレクタには
 * 依存していない（サイト側のクラス名変更に強い一方、想定外のレイアウトの
 * 施設では取りこぼす可能性がある）。
 * 初回実行時は必ず件数の少ない都道府県や --max で少数だけ試し、
 * 出力CSVの中身を確認してから本番投入することを推奨する。
 * うまく取れていない列があれば、そのCSVを見せてもらえれば抽出ロジックを
 * 調整する。
 */

const fs = require("fs");
const path = require("path");
const axios = require("axios");
const iconv = require("iconv-lite");

// =====================================================================
// 都道府県名 → 楽天トラベル内部エリアコード(f_chu)
// =====================================================================
// 「茨城県」→"ibaragi"は実データで確認済み（一般的なローマ字"ibaraki"ではない
// 楽天独自の内部コード）。他の46都道府県は標準的なローマ字表記を初期値として
// 入れているが、未検証。検索結果が0件になった場合はエラーで停止するので、
// その場合はこの表を実際の値に修正してから再実行してほしい
// （調べ方はREADME.md参照）。
const PREF_F_CHU = {
  "北海道": "hokkaido", "青森県": "aomori", "岩手県": "iwate", "宮城県": "miyagi",
  "秋田県": "akita", "山形県": "yamagata", "福島県": "fukushima",
  "茨城県": "ibaragi", // ★確認済み
  "栃木県": "tochigi", "群馬県": "gunma", "埼玉県": "saitama", "千葉県": "chiba",
  "東京都": "tokyo", "神奈川県": "kanagawa", "新潟県": "niigata", "富山県": "toyama",
  "石川県": "ishikawa", "福井県": "fukui", "山梨県": "yamanashi", "長野県": "nagano",
  "岐阜県": "gifu", "静岡県": "shizuoka", "愛知県": "aichi", "三重県": "mie",
  "滋賀県": "shiga", "京都府": "kyoto", "大阪府": "osaka", "兵庫県": "hyogo",
  "奈良県": "nara", "和歌山県": "wakayama", "鳥取県": "tottori", "島根県": "shimane",
  "岡山県": "okayama", "広島県": "hiroshima", "山口県": "yamaguchi", "徳島県": "tokushima",
  "香川県": "kagawa", "愛媛県": "ehime", "高知県": "kochi", "福岡県": "fukuoka",
  "佐賀県": "saga", "長崎県": "nagasaki", "熊本県": "kumamoto", "大分県": "oita",
  "宮崎県": "miyazaki", "鹿児島県": "kagoshima", "沖縄県": "okinawa"
};

// =====================================================================
// CLI引数
// =====================================================================
function parseArgs(argv) {
  const args = {
    pref: null,
    out: path.join(__dirname, "output"),
    hyoji: 30,
    delay: 500,
    concurrency: 3,
    max: Infinity
  };

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--pref") args.pref = argv[++i];
    else if (a === "--out") args.out = argv[++i];
    else if (a === "--hyoji") args.hyoji = parseInt(argv[++i], 10);
    else if (a === "--delay") args.delay = parseInt(argv[++i], 10);
    else if (a === "--concurrency") args.concurrency = parseInt(argv[++i], 10);
    else if (a === "--max") args.max = parseInt(argv[++i], 10);
    else if (a === "--help" || a === "-h") args.help = true;
  }

  return args;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// =====================================================================
// HTTP取得（文字コード自動判定：Shift-JIS / UTF-8 両対応）
// =====================================================================
async function fetchHtml(url, params) {
  const res = await axios.get(url, {
    params,
    responseType: "arraybuffer",
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      "Accept-Language": "ja,en;q=0.8"
    },
    timeout: 20000,
    validateStatus: s => s < 500
  });

  const contentType = String(res.headers["content-type"] || "");
  const isShiftJis = /shift[-_]?jis/i.test(contentType);
  const buf = Buffer.from(res.data);
  const html = isShiftJis ? iconv.decode(buf, "Shift_JIS") : buf.toString("utf8");

  return { html, status: res.status, url: res.request?.res?.responseUrl || url };
}

// =====================================================================
// 一覧ページのパース（正規表現ベース。DOM構造に依存しない）
// =====================================================================
function extractTotalCount(html) {
  const m = html.match(/(\d[\d,]*)\s*件中/);
  if (!m) return null;
  return parseInt(m[1].replace(/,/g, ""), 10);
}

// ホテルIDの出現位置ごとにブロックを区切り、各ブロック内から
// 施設名・価格・レビュー・住所・アクセス・駐車場を抜き出す。
function extractHotelBlocks(html) {
  const idPattern = /HOTEL\/(\d+)\/\1\.html/g;
  const positions = [];
  let m;

  while ((m = idPattern.exec(html)) !== null) {
    positions.push({ id: m[1], index: m.index });
  }

  // 同じIDが複数回（サムネイルリンク・タイトルリンク等）出現するため、
  // ID単位で最初の出現位置だけを採用し、そこから次のユニークIDの
  // 出現位置までをそのホテルの情報ブロックとして扱う。
  const uniqueFirst = [];
  const seen = new Set();
  for (const p of positions) {
    if (!seen.has(p.id)) {
      seen.add(p.id);
      uniqueFirst.push(p);
    }
  }

  const blocks = [];
  for (let i = 0; i < uniqueFirst.length; i++) {
    const start = uniqueFirst[i].index;
    const end = i + 1 < uniqueFirst.length ? uniqueFirst[i + 1].index : Math.min(html.length, start + 4000);
    blocks.push({ id: uniqueFirst[i].id, segment: html.slice(start, end) });
  }

  return blocks.map(b => parseHotelBlock(b.id, b.segment));
}

function stripTags(text) {
  return text
    .replace(/<[^>]+>/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function parseHotelBlock(id, segment) {
  const url = `https://travel.rakuten.co.jp/HOTEL/${id}/${id}.html`;

  // 施設名：href="...HOTEL/{id}/{id}.html" の直後にあるリンクテキスト、
  // または title="..." 属性から取る。
  let name = "";
  const titleAttrMatch = segment.match(new RegExp(`HOTEL/${id}/${id}\\.html"[^>]*title="([^"]+)"`));
  if (titleAttrMatch) {
    name = stripTags(titleAttrMatch[1]);
  } else {
    const linkTextMatch = segment.match(new RegExp(`HOTEL/${id}/${id}\\.html"[^>]*>([^<]{1,120})<`));
    if (linkTextMatch) name = stripTags(linkTextMatch[1]);
  }

  // 住所（〒付き郵便番号から始まるテキスト）
  let address = "";
  const addrMatch = segment.match(/〒\d{3}-\d{4}[^<\n]{0,60}/);
  if (addrMatch) address = stripTags(addrMatch[0]);

  // アクセス
  let access = "";
  const accessMatch = segment.match(/アクセス[^\S\n]*[:：]?\s*([^<\n]{5,150})/);
  if (accessMatch) access = stripTags(accessMatch[1]);

  // 駐車場
  let parking = "";
  const parkingMatch = segment.match(/駐車場[^\S\n]*[:：]?\s*([^<\n]{2,80})/);
  if (parkingMatch) parking = stripTags(parkingMatch[1]);

  // 最安値
  let price = "";
  const priceMatch = segment.match(/([\d,]+)\s*円/);
  if (priceMatch) price = priceMatch[1] + "円〜";

  // 口コミ（点数＋件数）
  let review = "";
  const reviewMatch = segment.match(/([\d.]{1,4})\s*[（(]\s*([\d,]+)\s*件[）)]/) ||
    segment.match(/([\d,]+)\s*件.{0,20}?([\d.]{3,4})/);
  if (reviewMatch) review = `評価${reviewMatch[1]} (${reviewMatch[2]}件)`;

  return { id, url, name, address, access, parking, price, review };
}

// =====================================================================
// 詳細ページ（_std.html）のパース：TEL・FAX・総部屋数・住所（保険）
// =====================================================================
function parseDetailPage(html) {
  const result = { address: "", tel: "", fax: "", access: "", parking: "", roomCount: "", description: "" };

  // meta descriptionに「総部屋数XX室」が確実に含まれているため、最優先で使う
  const metaDescMatch = html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["']/i);
  const metaDesc = metaDescMatch ? metaDescMatch[1] : "";

  const roomFromMeta = metaDesc.match(/総部屋数\s*(\d+)\s*室/);
  if (roomFromMeta) result.roomCount = roomFromMeta[1];

  // 「基本情報」ブロック内の 住所/TEL/FAX/交通アクセス/駐車場/総部屋数 は
  // ラベル文字列の直後に値が続く構造（実データで確認済み）。
  const addrMatch = html.match(/住所[^\S\n]*[\s\S]{0,10}?(〒\d{3}-\d{4}[^<\n]{0,60})/);
  if (addrMatch) result.address = stripTags(addrMatch[1]);

  const telMatch = html.match(/TEL[^\S\n]*[\s\S]{0,15}?(\d{2,4}-\d{2,4}-\d{3,4})/);
  if (telMatch) result.tel = telMatch[1];

  const faxMatch = html.match(/FAX[^\S\n]*[\s\S]{0,15}?(\d{2,4}-\d{2,4}-\d{3,4})/);
  if (faxMatch) result.fax = faxMatch[1];

  if (!result.roomCount) {
    const roomMatch = html.match(/総部屋数[^\S\n]*[\s\S]{0,10}?(\d+)\s*室/);
    if (roomMatch) result.roomCount = roomMatch[1];
  }

  const accessMatch = html.match(/交通アクセス[^\S\n]*[\s\S]{0,10}?([^<\n]{5,150})/);
  if (accessMatch) result.access = stripTags(accessMatch[1]);

  const parkingMatch = html.match(/駐車場[^\S\n]*[\s\S]{0,10}?([^<\n]{2,100})/);
  if (parkingMatch) result.parking = stripTags(parkingMatch[1]);

  // 施設説明：宿の注目ポイント直前にある短いキャッチコピー、無ければmeta descriptionの先頭部分
  const catchMatch = html.match(/<h1[^>]*>[\s\S]{0,300}?<\/h1>[\s\S]{0,400}?([^\s<>][^<]{10,120})/);
  result.description = catchMatch ? stripTags(catchMatch[1]) : stripTags(metaDesc).slice(0, 150);

  return result;
}

// =====================================================================
// メイン処理
// =====================================================================
async function scrapePrefecture(prefName, opts) {
  const fChu = PREF_F_CHU[prefName];

  if (!fChu) {
    throw new Error(
      `都道府県名「${prefName}」に対応するコードが見つかりません。PREF_F_CHUに追加してください。`
    );
  }

  console.log(`[scrape] ${prefName} (f_chu=${fChu}) の一覧取得を開始します...`);

  const listUrl = "https://web.travel.rakuten.co.jp/portal/my/search_undecided.main";
  const baseParams = {
    f_cd: "02",
    f_dai: "japan",
    f_chu: fChu,
    f_shou: "",
    f_sai: "",
    f_teikei: ":JPGo",
    f_hyoji: opts.hyoji,
    f_sort: "hotel_kin_high"
  };

  const allHotels = [];
  let totalCount = null;
  let page = 1;

  while (true) {
    const { html } = await fetchHtml(listUrl, { ...baseParams, f_page: page });

    if (totalCount === null) {
      totalCount = extractTotalCount(html);
      console.log(`[scrape] 総件数: ${totalCount === null ? "不明（0件の可能性あり）" : totalCount + "件"}`);

      if (!totalCount) {
        console.warn(
          `[scrape] 警告: 「${prefName}」の検索結果が0件、または件数を読み取れませんでした。\n` +
          `  f_chu="${fChu}" が正しいコードか確認してください（README.md参照）。`
        );
        break;
      }
    }

    const blocks = extractHotelBlocks(html);

    if (blocks.length === 0) {
      console.log(`[scrape] page=${page} で0件のため終了します。`);
      break;
    }

    blocks.forEach(b => allHotels.push(b));
    console.log(`[scrape] page=${page}: ${blocks.length}件取得（累計 ${allHotels.length}件）`);

    if (allHotels.length >= totalCount || allHotels.length >= opts.max) break;

    page += opts.hyoji;
    await sleep(opts.delay);
  }

  const limited = allHotels.slice(0, opts.max);

  console.log(`[scrape] 一覧取得完了: ${limited.length}件。詳細ページ（住所・TEL・部屋数）の取得を開始します...`);

  const results = [];
  let cursor = 0;

  async function worker() {
    while (cursor < limited.length) {
      const idx = cursor++;
      const hotel = limited[idx];
      const detailUrl = `https://travel.rakuten.co.jp/HOTEL/${hotel.id}/${hotel.id}_std.html`;

      try {
        const { html } = await fetchHtml(detailUrl);
        const detail = parseDetailPage(html);

        results[idx] = {
          name: hotel.name,
          url: hotel.url,
          pref: prefName,
          address: detail.address || hotel.address,
          tel: detail.tel,
          fax: detail.fax,
          roomCount: detail.roomCount,
          access: detail.access || hotel.access,
          parking: detail.parking || hotel.parking,
          description: detail.description,
          price: hotel.price,
          review: hotel.review
        };

        if ((idx + 1) % 10 === 0 || idx === limited.length - 1) {
          console.log(`[scrape] 詳細取得: ${idx + 1}/${limited.length}件`);
        }
      } catch (e) {
        console.warn(`[scrape] 詳細取得失敗 (id=${hotel.id}): ${e.message}`);
        results[idx] = {
          name: hotel.name,
          url: hotel.url,
          pref: prefName,
          address: hotel.address,
          tel: "",
          fax: "",
          roomCount: "",
          access: hotel.access,
          parking: hotel.parking,
          description: "",
          price: hotel.price,
          review: hotel.review
        };
      }

      await sleep(opts.delay);
    }
  }

  const workers = [];
  for (let i = 0; i < opts.concurrency; i++) workers.push(worker());
  await Promise.all(workers);

  return results;
}

// =====================================================================
// CSV出力（../code-rakuten.js が認識する列名で出力）
// =====================================================================
const CSV_HEADER = [
  "店名", "住所", "電話番号", "FAX", "総部屋数", "都道府県",
  "URL", "施設説明", "アクセス", "駐車場", "口コミ", "料金"
];

function toCsvCell(value) {
  const str = String(value === null || value === undefined ? "" : value).replace(/"/g, '""');
  if (str.indexOf(",") !== -1 || str.indexOf("\n") !== -1 || str.indexOf('"') !== -1) {
    return `"${str}"`;
  }
  return str;
}

function buildCsv(rows) {
  const lines = [CSV_HEADER.join(",")];

  rows.forEach(r => {
    lines.push([
      r.name, r.address, r.tel, r.fax, r.roomCount, r.pref,
      r.url, r.description, r.access, r.parking, r.review, r.price
    ].map(toCsvCell).join(","));
  });

  return lines.join("\r\n");
}

// =====================================================================
// エントリポイント
// =====================================================================
async function main() {
  const opts = parseArgs(process.argv.slice(2));

  if (opts.help || !opts.pref) {
    console.log(
      "使い方: node scrape.js --pref 茨城県 [--out ./output] [--hyoji 30] [--delay 500] [--concurrency 3] [--max 999999]"
    );
    process.exit(opts.help ? 0 : 1);
  }

  const rows = await scrapePrefecture(opts.pref, opts);

  fs.mkdirSync(opts.out, { recursive: true });

  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  const fileName = `楽天_${opts.pref}_${dateStr}.csv`;
  const filePath = path.join(opts.out, fileName);

  const bom = "﻿";
  fs.writeFileSync(filePath, bom + buildCsv(rows), "utf8");

  console.log(`[scrape] 完了: ${rows.length}件を書き出しました → ${filePath}`);
}

main().catch(e => {
  console.error("[scrape] エラー:", e);
  process.exit(1);
});
