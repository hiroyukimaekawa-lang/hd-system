/**
 * Ver10.0.1 オーケストレーションの統合テスト（Node.js / GASモック環境）。
 *
 * 検証項目:
 *  - API上限到達時にコムデスクCSVを出力しない
 *  - 実行時間による安全停止時にCSVを出力しない（一括処理開始時点からの経過時間で判定）
 *  - 再開完了後に再判定〜CSV出力〜サマリーまで実行される
 *  - 401/403発生時に一括処理全体が停止する（CSV・サマリー未確定、カーソル保存）
 *  - 認証エラー行から再開できる
 *  - 情報源側で住所を確認できない通常Google検索候補を自動採用しない（要確認送り）
 *
 * 実行方法: node gas/list-normalizer/test/run-integration-tests.js
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const CODE = fs.readFileSync(path.join(__dirname, "..", "code-restaurant.js"), "utf8");

// ---------- GASモック ----------
function makeIterator(items) {
  let i = 0;
  return { hasNext: () => i < items.length, next: () => items[i++] };
}

function MockSheet(name) {
  this.name = name;
  this.data = [];
}
MockSheet.prototype.getName = function () { return this.name; };
MockSheet.prototype.getLastRow = function () { return this.data.length; };
MockSheet.prototype.clear = function () { this.data = []; };
MockSheet.prototype.clearContents = function () { this.data = []; };
MockSheet.prototype.autoResizeColumns = function () {};
MockSheet.prototype.getDataRange = function () {
  const width = this.data.reduce((w, r) => Math.max(w, r.length), 0);
  const copy = this.data.map(r => {
    const row = r.slice();
    while (row.length < width) row.push("");
    return row;
  });
  return { getValues: () => copy };
};
MockSheet.prototype.getRange = function (row, col, numRows, numCols) {
  const sheet = this;
  numRows = numRows || 1;
  numCols = numCols || 1;
  const range = {
    setValues: function (values) {
      for (let r = 0; r < values.length; r++) {
        const tr = row - 1 + r;
        while (sheet.data.length <= tr) sheet.data.push([]);
        for (let c = 0; c < values[r].length; c++) {
          const tc = col - 1 + c;
          while (sheet.data[tr].length <= tc) sheet.data[tr].push("");
          sheet.data[tr][tc] = values[r][c];
        }
      }
      return range;
    },
    setValue: function (v) { return range.setValues([[v]]); },
    getValues: function () {
      const out = [];
      for (let r = 0; r < numRows; r++) {
        const src = sheet.data[row - 1 + r] || [];
        const line = [];
        for (let c = 0; c < numCols; c++) line.push(src[col - 1 + c] !== undefined ? src[col - 1 + c] : "");
        out.push(line);
      }
      return out;
    },
    setNumberFormat: function () { return range; },
    setFontWeight: function () { return range; },
    setDataValidation: function () { return range; }
  };
  return range;
};

function buildEnv(options) {
  options = options || {};
  const env = {
    props: Object.assign({
      SYSTEM_PROFILE: "ELECTRIC",
      SERPAPI_API_KEY: "test-key",
      PHONE_ENRICHMENT_DRY_RUN: "false"
    }, options.props || {}),
    exported: [],
    alerts: [],
    logs: [],
    fetchHandler: options.fetchHandler || (() => { throw new Error("fetchHandler not set"); })
  };

  const sheets = {};
  const order = [];
  const ss = {
    getId: () => "SSID",
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => { const s = new MockSheet(n); sheets[n] = s; order.push(s); return s; },
    getSheets: () => order.slice(),
    deleteSheet: s => { delete sheets[s.getName()]; const i = order.indexOf(s); if (i !== -1) order.splice(i, 1); },
    toast: () => {},
    setActiveSheet: () => {},
    moveActiveSheet: () => {}
  };
  env.ss = ss;
  env.sheet = n => sheets[n] || null;

  function makeFolder(name) {
    const folder = {
      name: name,
      subFolders: {},
      getFoldersByName: n => makeIterator(folder.subFolders[n] ? [folder.subFolders[n]] : []),
      createFolder: n => { const f = makeFolder(n); folder.subFolders[n] = f; return f; },
      getFilesByType: () => makeIterator([]),
      getFilesByName: () => makeIterator([]),
      createFile: blob => { env.exported.push(blob); return {}; }
    };
    return folder;
  }
  const rootFolder = makeFolder("root");

  const sandbox = {
    console,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (k in env.props ? env.props[k] : null),
        setProperty: (k, v) => { env.props[k] = String(v); },
        deleteProperty: k => { delete env.props[k]; }
      })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({ alert: msg => env.alerts.push(String(msg)) }),
      newDataValidation: () => ({ requireCheckbox: () => ({ build: () => ({}) }) })
    },
    DriveApp: {
      getFileById: () => ({ getParents: () => makeIterator([rootFolder]) })
    },
    UrlFetchApp: {
      fetch: url => {
        const res = env.fetchHandler(url);
        return { getResponseCode: () => res.code, getContentText: () => (typeof res.body === "string" ? res.body : JSON.stringify(res.body)) };
      }
    },
    Utilities: {
      sleep: () => {},
      parseCsv: () => [],
      newBlob: (content, type, name) => ({ content, type, name }),
      formatDate: (d, tz, fmt) => (fmt === "yyyyMMdd" ? "20260721" : "2026-07-21 00:00:00")
    },
    MimeType: { CSV: "text/csv" },
    Logger: { log: msg => env.logs.push(String(msg)) }
  };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox);
  env.g = sandbox;

  // シードデータ: 電話番号ありの店1件＋電話番号なし（補完対象）の店1件
  const norm = ss.insertSheet("01_NORMALIZED");
  norm.data = [
    ["店名", "住所", "都道府県", "市区町村", "電話番号", "ジャンル"],
    ["らーめん一番", "茨城県古河市本町1-2-3", "茨城県", "古河市", "0280-11-1111", "ラーメン"],
    ["そば処まえかわ", "茨城県古河市中央町2-3-4", "茨城県", "古河市", "", "そば"]
  ];
  const master = ss.insertSheet("MASTER_CHAIN");
  master.data = [["チェーン名", "除外キーワード", "業種", "除外対象", "メモ"]];

  return env;
}

// 成功応答（google_maps検索で完全一致候補を返す）
function successFetchHandler(url) {
  if (url.indexOf("engine=google_maps") !== -1) {
    return {
      code: 200,
      body: {
        local_results: [{
          title: "そば処まえかわ",
          address: "茨城県古河市中央町2-3-4",
          phone: "0280-22-2222",
          website: "https://soba-maekawa.example.com",
          place_id: "PID-1",
          type: "そば屋"
        }]
      }
    };
  }
  return { code: 200, body: { organic_results: [] } };
}

// ---------- テストランナー ----------
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

// ================================================================
// テスト1: API上限到達時にCSVを出力しない（＋停止情報の表示）
// ================================================================
(function () {
  const env = buildEnv({ props: { MAX_SERPAPI_CALLS_PER_RUN: "0" } });
  env.fetchHandler = () => { throw new Error("API上限0なのに呼び出された"); };
  env.g.executeAllProcesses();

  check("T1: CSVが出力されない", env.exported.length, 0);
  checkTrue("T1: カーソルが保存される", "PHONE_ENRICHMENT_CURSOR_ELECTRIC" in env.props);
  check("T1: 後続処理未確定フラグ", env.props.PHONE_ENRICHMENT_PENDING_FINALIZE_ELECTRIC, "true");
  const alertText = env.alerts.join("\n");
  checkTrue("T1: 途中停止である旨を表示", alertText.indexOf("途中で停止") !== -1);
  checkTrue("T1: 処理済み件数を表示", alertText.indexOf("処理済み: 0件") !== -1);
  checkTrue("T1: 残件数を表示", alertText.indexOf("残り: 1件") !== -1);
  checkTrue("T1: 再開方法を表示", alertText.indexOf("再開") !== -1);
  checkTrue("T1: CSV未出力である旨を表示", alertText.indexOf("CSVはまだ出力されていません") !== -1);
  checkTrue("T1: 営業対象シートが確定していない", !env.sheet("04_SALES_ラーメン"));
  checkTrue("T1: 件数サマリーが確定していない", !env.sheet("00_件数サマリー"));

  // ================================================================
  // テスト3: 再開完了後に再判定〜CSV出力〜サマリーまで実行される
  // ================================================================
  env.props.MAX_SERPAPI_CALLS_PER_RUN = "100";
  env.fetchHandler = successFetchHandler;
  env.g.resumePhoneEnrichmentFromTrigger();

  checkTrue("T3: 再開完了後にCSVが出力される", env.exported.length > 0);
  checkTrue("T3: ファイル名に運用区分を含む", env.exported.every(f => f.name.indexOf("コムデスク_電気営業_") === 0));
  checkTrue("T3: カーソルがクリアされる", !("PHONE_ENRICHMENT_CURSOR_ELECTRIC" in env.props));
  checkTrue("T3: 後続処理未確定フラグがクリアされる", !("PHONE_ENRICHMENT_PENDING_FINALIZE_ELECTRIC" in env.props));
  const normData = env.sheet("01_NORMALIZED").getDataRange().getValues();
  const header = normData[0];
  const phoneIdx = header.indexOf("電話番号");
  const sobaRow = normData.find(r => r[0] === "そば処まえかわ");
  checkTrue("T3: 補完電話番号が01_NORMALIZEDへ反映される", sobaRow && String(sobaRow[phoneIdx]).replace(/[^\d]/g, "") === "0280222222");
  checkTrue("T3: 件数サマリーが更新される", !!env.sheet("00_件数サマリー"));
  const salesSheets = env.ss.getSheets().filter(s => s.getName().indexOf("04_SALES_") === 0);
  checkTrue("T3: 営業対象ジャンルタブが生成される", salesSheets.length >= 2); // ラーメン＋和食(蕎麦・うどん統合)
  checkTrue("T3: 補完完了ログ", env.logs.join("\n").indexOf("コムデスクCSV出力〜件数サマリーが完了") !== -1);
})();

// ================================================================
// テスト2: 実行時間停止時にCSVを出力しない（一括処理開始時点からの経過時間）
// ================================================================
(function () {
  const env = buildEnv({});
  env.fetchHandler = () => { throw new Error("時間切れなのに呼び出された"); };
  // 判定パイプラインを先に実行して04_FACILITY_CHECKを作る
  env.g.runJudgmentPipeline_();
  // 一括処理の開始時刻として「大昔」を渡す → 経過時間超過で即・安全停止
  const summary = env.g.executePhoneEnrichment(Date.now() - 10 * 60 * 1000);

  check("T2: 時間超過でfinished=false", summary.finished, false);
  check("T2: API呼び出しなし", summary.apiCalls, 0);
  checkTrue("T2: カーソルが保存される", "PHONE_ENRICHMENT_CURSOR_ELECTRIC" in env.props);
  check("T2: 残件数を報告", summary.remainingCount, 1);

  // 未完了サマリーでは後続処理が実行されない
  const result = env.g.maybeFinalizePipeline_(summary, false);
  check("T2: 後続処理がスキップされる", result, null);
  check("T2: CSVが出力されない", env.exported.length, 0);
  checkTrue("T2: 件数サマリーが確定していない", !env.sheet("00_件数サマリー"));

  // 開始時刻が現在なら停止しない（同一環境・成功応答で完走することを確認）
  env.fetchHandler = successFetchHandler;
  delete env.props.PHONE_ENRICHMENT_CURSOR_ELECTRIC;
  const summary2 = env.g.executePhoneEnrichment(Date.now());
  check("T2: 経過時間内なら完走する", summary2.finished, true);
})();

// ================================================================
// テスト4: 401/403で一括処理全体が停止する
// ================================================================
(function () {
  const env = buildEnv({});
  env.fetchHandler = () => ({ code: 401, body: { error: "invalid key" } });
  env.g.executeAllProcesses();

  check("T4: CSVが出力されない", env.exported.length, 0);
  checkTrue("T4: 営業対象シートが確定しない", env.ss.getSheets().every(s => s.getName().indexOf("04_SALES_") !== 0));
  checkTrue("T4: 件数サマリーが確定しない", !env.sheet("00_件数サマリー"));
  checkTrue("T4: カーソルが保存される（対象行から再開可能）", "PHONE_ENRICHMENT_CURSOR_ELECTRIC" in env.props);
  check("T4: 後続処理未確定フラグ", env.props.PHONE_ENRICHMENT_PENDING_FINALIZE_ELECTRIC, "true");
  const alertText = env.alerts.join("\n");
  checkTrue("T4: 認証エラーである旨を表示", alertText.indexOf("認証エラー") !== -1);
  checkTrue("T4: APIキー修正の案内を表示", alertText.indexOf("SERPAPI_API_KEY") !== -1);
  checkTrue("T4: CSV未出力である旨を表示", alertText.indexOf("CSVはまだ出力されていません") !== -1);

  // 対象行のステータスにAPIエラーが記録される
  const normData = env.sheet("01_NORMALIZED").getDataRange().getValues();
  const header = normData[0];
  const statusIdx = header.indexOf("電話補完ステータス");
  const sobaRow = normData.find(r => r[0] === "そば処まえかわ");
  check("T4: 対象行にAPIエラーを記録", sobaRow ? sobaRow[statusIdx] : "", "APIエラー");

  // ================================================================
  // テスト5: APIキー修正後、認証エラー行から再開できる
  // ================================================================
  const cursor = env.props.PHONE_ENRICHMENT_CURSOR_ELECTRIC;
  env.props.SERPAPI_API_KEY = "fixed-key";
  env.fetchHandler = url => {
    checkTrue("T5: 修正後キーで呼び出される", url.indexOf("api_key=fixed-key") !== -1);
    return successFetchHandler(url);
  };
  env.g.resumePhoneEnrichmentFromTrigger();

  checkTrue("T5: 認証エラー行から再開して補完される", (() => {
    const data = env.sheet("01_NORMALIZED").getDataRange().getValues();
    const h = data[0];
    const row = data.find(r => r[0] === "そば処まえかわ");
    return row && String(row[h.indexOf("電話番号")]).replace(/[^\d]/g, "") === "0280222222";
  })());
  checkTrue("T5: 再開後にCSVが出力される", env.exported.length > 0);
  checkTrue("T5: カーソルがクリアされる", !("PHONE_ENRICHMENT_CURSOR_ELECTRIC" in env.props));
  checkTrue("T5: カーソルは停止行を指していた", cursor === "2" || cursor === "1"); // 04_FACILITY_CHECKの対象行位置
})();

// ================================================================
// テスト6: 情報源側で住所を確認できない通常Google検索候補は自動採用しない
// ================================================================
(function () {
  const env = buildEnv({});
  env.fetchHandler = url => {
    if (url.indexOf("engine=google_maps") !== -1) {
      return { code: 200, body: {} }; // Google Mapsでは見つからない
    }
    // 通常Google検索: 店名と電話番号はあるが住所が確認できないスニペット
    return {
      code: 200,
      body: {
        organic_results: [{
          title: "そば処まえかわ - 電話番号情報",
          snippet: "そば処まえかわ TEL 0280-33-3333 の店舗情報ページです。",
          link: "https://directory.example.com/soba-maekawa"
        }]
      }
    };
  };
  env.g.runJudgmentPipeline_();
  const summary = env.g.executePhoneEnrichment(Date.now());

  check("T6: 自動採用されない", summary.autoCount, 0);
  check("T6: 要確認へ送られる", summary.reviewCount, 1);
  const normData = env.sheet("01_NORMALIZED").getDataRange().getValues();
  const h = normData[0];
  const row = normData.find(r => r[0] === "そば処まえかわ");
  check("T6: 元の電話番号は空のまま", row[h.indexOf("電話番号")], "");
  check("T6: ステータスは要確認", row[h.indexOf("電話補完ステータス")], "要確認");
  const review = env.sheet("確認_電話番号候補");
  checkTrue("T6: 確認シートに候補が記録される", review && review.data.length >= 2);
  const reviewRow = review.data[1];
  checkTrue("T6: 判定理由に住所未確認を含む", String(reviewRow[6]).indexOf("住所を確認できず") !== -1);
  checkTrue("T6: 候補側店名に対象の値を無条件代入しない（住所欄が空）", reviewRow[4] === "");
})();

// ---- 結果 ----
console.log(`\n${pass} passed / ${fail} failed`);
if (failures.length) {
  console.log("\n" + failures.join("\n\n"));
  process.exit(1);
}
