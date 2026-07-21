/**
 * code-restaurant.js の純粋関数を Node.js で検証するテストハーネス。
 * GAS環境依存API（SpreadsheetApp等）はスタブ化し、電話番号補完の
 * 正規化・スコアリング・判定ロジックとプロファイル切替を検証する。
 *
 * 実行方法: node gas/list-normalizer/test/run-tests.js
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

// ---- GASスタブ ----
const scriptProps = { SYSTEM_PROFILE: "ELECTRIC" };
const sandbox = {
  console,
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: k => (k in scriptProps ? scriptProps[k] : null),
      setProperty: (k, v) => { scriptProps[k] = v; },
      deleteProperty: k => { delete scriptProps[k]; }
    })
  },
  SpreadsheetApp: { getActiveSpreadsheet: () => { throw new Error("not stubbed"); }, getUi: () => ({ alert: () => {} }) },
  Utilities: {
    sleep: () => {},
    formatDate: () => "2026-07-21 00:00:00",
    parseCsv: () => [],
    newBlob: () => ({})
  },
  UrlFetchApp: { fetch: () => { throw new Error("network disabled in tests"); } },
  DriveApp: {},
  MimeType: { CSV: "text/csv" },
  Logger: { log: () => {} }
};
vm.createContext(sandbox);
const code = fs.readFileSync(path.join(__dirname, "..", "code-restaurant.js"), "utf8");
// const/letはvmコンテキストのグローバルに載らないため、テスト用ブリッジを追加する
const bridge = `
;this.__test = {
  PHONE_ENRICHMENT_HEADERS: PHONE_ENRICHMENT_HEADERS,
  COMMON_EXCLUDE_KEYWORDS: COMMON_EXCLUDE_KEYWORDS,
  resetProfileCache: function () { ACTIVE_SYSTEM_PROFILE_CACHE_ = null; }
};`;
vm.runInContext(code + bridge, sandbox);

// ---- テストランナー ----
let pass = 0, fail = 0;
const failures = [];
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) { pass++; } else {
    fail++;
    failures.push(`✗ ${label}\n    expected: ${JSON.stringify(expected)}\n    actual:   ${JSON.stringify(actual)}`);
  }
}
function checkTrue(label, cond) { check(label, !!cond, true); }

const g = sandbox;

// ================= 電話番号の検証・正規化（§9 / 18.1） =================
check("固定電話10桁", g.validateJpPhoneNumber_("0280-12-3456").ok, true);
check("固定電話 表示形式", g.validateJpPhoneNumber_("0280123456").display, "028-012-3456".length === 12 ? g.normalizePhoneNumberForAnalysis("0280123456") : "");
check("携帯 090", g.validateJpPhoneNumber_("090-1234-5678"), { ok: true, digits: "09012345678", display: "090-1234-5678" });
check("050番号", g.validateJpPhoneNumber_("05012345678").ok, true);
check("0120", g.validateJpPhoneNumber_("0120-123-456").ok, true);
check("0570", g.validateJpPhoneNumber_("0570-000-111").ok, true);
check("+81変換", g.validateJpPhoneNumber_("+81 90 1234 5678"), { ok: true, digits: "09012345678", display: "090-1234-5678" });
check("全角数字", g.validateJpPhoneNumber_("０９０１２３４５６７８").ok, true);
check("桁不足はNG", g.validateJpPhoneNumber_("03-1234").ok, false);
check("0始まりでないものはNG", g.validateJpPhoneNumber_("1234567890").ok, false);
check("空はNG", g.validateJpPhoneNumber_("").ok, false);

// ================= 店舗名の正規化（§8.1） =================
check("法人格の除外", g.normalizeStoreNameForMatch("株式会社マエカワ食堂"), g.normalizeStoreNameForMatch("マエカワ食堂"));
checkTrue("支店名は残す", g.normalizeStoreNameForMatch("すし処まえかわ 古河店").indexOf("古河店") !== -1);
check("全角半角統一", g.normalizeStoreNameForMatch("ＣＡＦＥ　ＭＡＥＫＡＷＡ"), g.normalizeStoreNameForMatch("cafe maekawa"));
check("支店ラベル抽出", g.extractBranchLabel_("スシロー 古河店"), "古河店");
check("支店ラベルなし", g.extractBranchLabel_("スシロー"), "");

// ================= 住所の正規化（§8.2） =================
const addr = g.normalizeAddressForMatch("〒306-0023 茨城県古河市本町１丁目２−３ サンプルビル２Ｆ");
check("都道府県分割", addr.pref, "茨城県");
check("市区町村分割", addr.city, "古河市");
check("町域分割", addr.town, "本町");
check("番地分割", addr.banchi, "1-2-3");
checkTrue("比較用住所からビル・階数を除去", addr.comparable.indexOf("ビル") === -1 && addr.comparable.indexOf("2F") === -1);
checkTrue("比較用住所に番地は残す", addr.comparable.indexOf("1-2-3") !== -1);
const addr2 = g.normalizeAddressForMatch("日本、〒306-0023 茨城県古河市本町1番地2");
check("番地表記の統一", addr2.banchi, "1-2");
const addrChiku = g.normalizeAddressForMatch("東京都中央区築地4-5-6");
check("地名中の「地」を壊さない", addrChiku.town, "築地");

// ================= 信頼度スコア（§8.3） =================
const target = {
  name: "そば処まえかわ",
  rawAddress: "茨城県古河市本町1-2-3",
  addr: g.normalizeAddressForMatch("茨城県古河市本町1-2-3"),
  genre: "蕎麦・うどん"
};
const perfect = {
  name: "そば処まえかわ", address: "茨城県古河市本町1-2-3", phone: "0280-12-3456",
  website: "https://example.com", placeId: "PID1", dataId: "", dataCid: "",
  type: "そば", sourceUrl: "https://maps.google.com/x", source: "Google Maps",
  permanentlyClosed: false, reviewOnly: false
};
const s1 = g.scorePhoneCandidate_(target, perfect);
check("完全一致スコア=100", s1.score, 100);

const config = { profile: "ELECTRIC", maxCallsPerRun: 100, minAutoAcceptScore: 85, dryRun: true, reviewMinScore: 70, minScoreGap: 15, safeStopMillis: 240000 };

// 18.1: 店舗名と住所が完全一致する候補を高信頼度で補完できる
const evalPerfect = g.evaluatePhoneCandidates_(target, [perfect], config);
check("完全一致はAUTO", evalPerfect.decision, "AUTO");
// 表示形式は既存システムと同一（10桁は3-3-4区切り）§9
check("採用電話番号の正規化", evalPerfect.phone, g.normalizePhoneNumberForAnalysis("0280123456"));

// 18.1: 同名だが住所が異なる店舗を自動採用しない
const wrongCity = Object.assign({}, perfect, { address: "茨城県結城市本町1-2-3" });
const evalWrongCity = g.evaluatePhoneCandidates_(target, [wrongCity], config);
checkTrue("市区町村不一致はAUTOにしない", evalWrongCity.decision !== "AUTO");
checkTrue("市区町村不一致が理由に含まれる", evalWrongCity.reasons.indexOf("市区町村不一致") !== -1);

// 18.1: 支店名が異なる候補を自動採用しない
const branchTarget = Object.assign({}, target, { name: "スシロー 古河店" });
const branchCand = Object.assign({}, perfect, { name: "スシロー 結城店" });
const evalBranch = g.evaluatePhoneCandidates_(branchTarget, [branchCand], config);
checkTrue("支店名不一致はAUTOにしない", evalBranch.decision !== "AUTO");

// 18.1: 1位と2位のスコア差が15点未満の場合に要確認へ送られる
const near1 = Object.assign({}, perfect, { phone: "0280-12-3456" });
const near2 = Object.assign({}, perfect, { name: "そば処まえかわ本店", phone: "0280-99-9999" });
const evalNear = g.evaluatePhoneCandidates_(target, [near1, near2], config);
checkTrue("スコア差15点未満はAUTOにしない", evalNear.decision !== "AUTO");

// 複数の電話番号が競合する場合
checkTrue("電話番号競合が理由に含まれる", evalNear.reasons.some(r => r.indexOf("競合") !== -1 || r.indexOf("スコア差") !== -1));

// 閉業候補
const closedCand = Object.assign({}, perfect, { permanentlyClosed: true });
const evalClosed = g.evaluatePhoneCandidates_(target, [closedCand], config);
checkTrue("閉業の可能性はAUTOにしない", evalClosed.decision !== "AUTO");

// 店舗名のみ一致（住所裏付けなし）
const nameOnly = Object.assign({}, perfect, { address: "" });
const evalNameOnly = g.evaluatePhoneCandidates_(target, [nameOnly], config);
checkTrue("店名のみ一致はAUTOにしない", evalNameOnly.decision !== "AUTO");

// 単一の第三者サイトのみ（§7.4）
const thirdParty = Object.assign({}, perfect, { reviewOnly: true, source: "第三者サイト" });
const evalThird = g.evaluatePhoneCandidates_(target, [thirdParty], config);
checkTrue("単一第三者サイトはAUTOにしない", evalThird.decision !== "AUTO");

// 候補なし
check("候補なしはNONE", g.evaluatePhoneCandidates_(target, [], config).decision, "NONE");

// 低スコア（69点以下）は自動採用しない
const weak = { name: "ぜんぜん違う店", address: "北海道札幌市中央区1-1", phone: "011-123-4567", website: "", placeId: "", dataId: "", dataCid: "", type: "", sourceUrl: "", source: "Google Maps", permanentlyClosed: false, reviewOnly: false };
checkTrue("低スコアはAUTOにしない", g.evaluatePhoneCandidates_(target, [weak], config).decision !== "AUTO");

// ================= プロファイル切替（§5 / 18.3） =================
function setProfile(p) {
  scriptProps.SYSTEM_PROFILE = p;
  g.__test.resetProfileCache();
}

setProfile("ELECTRIC");
check("ELECTRIC: ビル表記は確認対象", g.judgeFacilityStatus("喫茶マエカワ", "茨城県古河市本町1-2-3 第一ビル2F", "カフェ").status, "確認対象");
setProfile("AFFILIATE");
check("AFFILIATE: ビル表記のみでは除外しない", g.judgeFacilityStatus("喫茶マエカワ", "茨城県古河市本町1-2-3 第一ビル2F", "カフェ").status, "対象");

// 共通除外7種は両プロファイルで除外
const commonKeywords = ["道の駅", "カラオケ", "総合公園", "ゴルフ", "ホテル", "スーパー", "イオンモール"];
["ELECTRIC", "AFFILIATE"].forEach(p => {
  setProfile(p);
  commonKeywords.forEach(kw => {
    check(`${p}: 共通除外「${kw}」`, g.judgeFacilityStatus(`${kw}テスト店`, "茨城県古河市本町1-2-3", "").status, "除外");
  });
});

// 既存の大型施設除外が維持されている（縮小されていない）
setProfile("ELECTRIC");
["病院", "パチンコ", "ショッピングモール", "フードコート", "天然温泉"].forEach(kw => {
  check(`既存除外の維持「${kw}」`, g.judgeFacilityStatus(`${kw}前食堂`, "", "").status, "除外");
});

// 共通除外はプロファイルより優先（AFFILIATEでもホテル内ビル店舗は除外）
setProfile("AFFILIATE");
check("AFFILIATE: 共通除外がビル判定より優先", g.judgeFacilityStatus("ホテルマエカワ", "茨城県古河市本町1-2-3 タワー3F", "").status, "除外");

// SYSTEM_PROFILE未設定は明確なエラー
delete scriptProps.SYSTEM_PROFILE;
g.__test.resetProfileCache();
let profileError = "";
try { g.getActiveSystemProfile(); } catch (e) { profileError = e.message; }
checkTrue("SYSTEM_PROFILE未設定でエラー", profileError.indexOf("SYSTEM_PROFILE") !== -1);

// SERPAPI_API_KEY未設定は秘密情報を含まない明確なエラー
delete scriptProps.SERPAPI_API_KEY;
let keyError = "";
try { g.getSerpApiKey_(); } catch (e) { keyError = e.message; }
checkTrue("APIキー未設定でエラー", keyError.indexOf("SERPAPI_API_KEY") !== -1);

// ================= 補完対象の抽出（§7.1） =================
setProfile("ELECTRIC");
const header = ["店名", "住所", "電話番号", "正規化電話番号", "取得ステータス", "重複判定", "チェーン判定", "エリア判定", "施設判定", "電話補完ステータス"];
function makeRow(over) {
  const base = { "店名": "テスト店", "住所": "茨城県古河市本町1-2-3", "電話番号": "", "正規化電話番号": "", "取得ステータス": "", "重複判定": "ユニーク", "チェーン判定": "単独店", "エリア判定": "エリア内", "施設判定": "対象", "電話補完ステータス": "" };
  Object.assign(base, over || {});
  return header.map(h => base[h]);
}
check("補完対象になる", g.judgePhoneEnrichmentTarget_(header, makeRow(), "ELECTRIC").target, true);
check("電話番号ありは対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "電話番号": "0280-12-3456" }), "ELECTRIC").target, false);
check("重複行は対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "重複判定": "重複（電話番号一致）" }), "ELECTRIC").target, false);
check("チェーン店は対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "チェーン判定": "チェーン店" }), "ELECTRIC").target, false);
check("エリア外は対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "エリア判定": "エリア外" }), "ELECTRIC").target, false);
check("取得失敗は対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "取得ステータス": "失敗" }), "ELECTRIC").target, false);
check("共通除外施設は対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "施設判定": "除外" }), "ELECTRIC").target, false);
check("ELECTRIC: ビル確認候補は対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "施設判定": "確認対象" }), "ELECTRIC").target, false);
check("AFFILIATE: ビル表記ありでも対象（施設判定=対象）", g.judgePhoneEnrichmentTarget_(header, makeRow(), "AFFILIATE").target, true);
check("高信頼補完済みは再検索しない", g.judgePhoneEnrichmentTarget_(header, makeRow({ "電話補完ステータス": "高信頼補完" }), "ELECTRIC").target, false);
check("店名なしは対象外", g.judgePhoneEnrichmentTarget_(header, makeRow({ "店名": "" }), "ELECTRIC").target, false);

// ================= 安定キー（§12） =================
const keyHeader = ["店名", "住所", "GoogleマップURL"];
check("GoogleマップURL優先", g.buildStableRowKey_(keyHeader, ["店A", "住所A", "https://maps.google.com/abc"]), "URL::https://maps.google.com/abc");
checkTrue("URLなしは店名+住所キー", g.buildStableRowKey_(["店名", "住所"], ["店A", "茨城県古河市本町1-2-3"]).startsWith("NA::"));

// ================= コムデスク31列仕様の維持（§17 / 18.4） =================
check("コムデスクヘッダーは31列", g.getComdeskHeader().length, 31);
check("列順維持（先頭5列）", g.getComdeskHeader().slice(0, 5), ["UUID", "種別", "名前", "カナ", "郵便番号"]);
check("列順維持（末尾4列）", g.getComdeskHeader().slice(27), ["午前始", "午前終", "午後始", "午後終"]);
checkTrue("監査列がコムデスクヘッダーに混入しない", g.getComdeskHeader().every(c => g.__test.PHONE_ENRICHMENT_HEADERS.indexOf(c) === -1));

// ================= エラー分類（§14.3 / 18.2） =================
check("429は再試行対象", g.isSerpApiRetryable_(429), true);
check("503は再試行対象", g.isSerpApiRetryable_(503), true);
check("401は認証エラー", g.isSerpApiAuthError_(401), true);
check("403は認証エラー", g.isSerpApiAuthError_(403), true);
check("404は再試行しない", g.isSerpApiRetryable_(404), false);

// ---- 結果 ----
console.log(`\n${pass} passed / ${fail} failed`);
if (failures.length) {
  console.log("\n" + failures.join("\n\n"));
  process.exit(1);
}
