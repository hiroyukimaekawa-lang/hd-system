/**
 * offscreen.js (改修版: 複数営業時間枠A/Bパース + 高精度HP判定 + 実店舗電話番号優先抽出対応)
 */

const activeTasks = new Map();
const CHUNK_SIZE = 5;
const DELAY_BETWEEN_CHUNKS = 800;
const DELAY_LIST_FETCH = 600;

const genreLinksResolvers = new Map();

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// =====================================================================
// 【新規】住所文字列から都道府県と市区町村を抽出する関数
// =====================================================================
function parseAddress(address) {
  let cleanAddress = address.replace(/(?:〒\d{3}-\d{4}\s*|日本、\s*)/g, '').trim();
  const regex = /^((?:北海道|東京都|大阪府|京都府|.{2,3}県))?((?:.+?郡.+?[町村]|.+?市.+?区|.+?[市区町村]))?(.+)?$/;
  const m = cleanAddress.match(regex);
  if (!m) return { prefecture: '', city: '' };
  return { prefecture: m[1] || '', city: m[2] || '' };
}

// 【新規】セルのテキストを改行タグ(br)を保持して取得
function _extractCellLines(node) {
  if (!node) return '';
  const clone = node.cloneNode(true);
  clone.querySelectorAll('br').forEach(br => br.replaceWith(document.createTextNode('\n')));
  return clone.textContent.replace(/[ \t\u3000]+/g, ' ').trim();
}

// =====================================================================
// 【新規】050予約専用番号を回避し、実店舗の固定電話等を優先抽出する関数
// =====================================================================
function selectBestPhoneNumber(rawText) {
  if (!rawText) return '';
  
  // テキスト内から電話番号のパターン（数字とハイフンの組み合わせ）をすべて抽出
  const matches = rawText.match(/(?:\d{2,5}-\d{1,4}-\d{3,4}|\d{10,11})/g);
  if (!matches || matches.length === 0) {
    return rawText.replace(/[^\d\-]/g, ''); // マッチしなければ記号除去のみ
  }
  
  // 050から始まらない番号（実店舗の固定電話や携帯など）を最優先でフィルタリング
  const non050Numbers = matches.filter(num => !num.startsWith('050'));
  if (non050Numbers.length > 0) {
    return non050Numbers[0].replace(/[^\d\-]/g, ''); // 最初の固定電話などを採用
  }
  
  // 050番号しか存在しない場合は、最初の050番号を返す
  return matches[0].replace(/[^\d\-]/g, '');
}

function getSiteType(url) {
  if (/tabelog\.com/.test(url)) return 'tabelog';
  if (/hotpepper\.jp/.test(url)) return 'hotpepper';
  return null;
}

function resolveUrl(href, baseUrl) {
  if (!href) return '';
  try {
    return new URL(href, baseUrl).href;
  } catch (e) {
    return href;
  }
}

function normalizeListUrl(url) {
  return String(url || '').split('#')[0];
}

function addUniqueLink(links, href) {
  const normalized = normalizeListUrl(href).split('?')[0];
  if (normalized && !links.includes(normalized)) links.push(normalized);
}

function isTabelogRestaurantUrl(href) {
  if (!href || !/tabelog\.com/.test(href)) return false;
  let url;
  try {
    url = new URL(href);
  } catch (_) {
    return false;
  }

  const path = url.pathname;
  if (!/\/\d{8}\/?$/.test(path)) return false;
  if (/\/(?:rvwr|rstLst|dtlrvwlst|dtlmenu|party|map|photo|coupon|award)\//.test(path)) return false;
  return /\/[a-z]+\/A\d+\/A\d+\/\d{8}\/?$/.test(path)
    || /\/\d{8}\/?$/.test(path);
}

function isLikelyNextAnchor(a) {
  if (!a) return false;
  const text = (a.textContent || '').replace(/\s+/g, '');
  const label = `${a.getAttribute('aria-label') || ''} ${a.getAttribute('title') || ''}`.replace(/\s+/g, '');
  const rel = (a.getAttribute('rel') || '').toLowerCase();
  const cls = a.className || '';
  return rel === 'next'
    || /次|次へ|Next|next/.test(text)
    || /次|次へ|Next|next/.test(label)
    || /next|pa_next|pagination-next/.test(String(cls));
}

function deriveTabelogNextUrl(currentUrl, pageNum) {
  try {
    const url = new URL(currentUrl);
    const nextPage = pageNum + 1;
    const parts = url.pathname.split('/').filter(Boolean);
    const rstIdx = parts.indexOf('rstLst');
    if (rstIdx === -1) return null;

    const last = parts[parts.length - 1];
    if (/^\d+$/.test(last)) {
      parts[parts.length - 1] = String(nextPage);
    } else {
      parts.push(String(nextPage));
    }
    url.pathname = '/' + parts.join('/') + '/';
    return url.href;
  } catch (e) {
    return null;
  }
}

function deriveHotpepperNextUrl(currentUrl, pageNum) {
  try {
    const url = new URL(currentUrl);
    const nextPage = pageNum + 1;
    const parts = url.pathname.split('/').filter(Boolean);
    const pnIdx = parts.findIndex(p => /^pn\d+$/i.test(p));
    if (pnIdx !== -1) {
      parts[pnIdx] = `pn${nextPage}`;
    } else {
      parts.push(`pn${nextPage}`);
    }
    url.pathname = '/' + parts.join('/') + '/';
    return url.href;
  } catch (e) {
    return null;
  }
}

function sendToBackground(tabId, type, payload = {}) {
  chrome.runtime.sendMessage({
    target: 'background',
    tabId,
    type,
    payload
  }).catch(() => { });
}

function extractMetadata(doc, siteType) {
  const meta = { area: '', industry: '' };
  if (siteType === 'tabelog') {
    meta.area = doc.querySelector('.list-condition__item--area')?.textContent?.trim()
      || doc.querySelector('.c-link-arrow--back')?.textContent?.trim() || '';
    meta.industry = doc.querySelector('.list-condition__item--genre')?.textContent?.trim() || '';
  } else if (siteType === 'hotpepper') {
    meta.area = doc.querySelector('.current-area')?.textContent?.trim() || '';
    meta.industry = doc.querySelector('.current-genre')?.textContent?.trim() || '';
  }
  return meta;
}

const DAY_MAP = {
  '月曜日':'月','火曜日':'火','水曜日':'水','木曜日':'木',
  '金曜日':'金','土曜日':'土','日曜日':'日',
  'Monday':'月','Tuesday':'火','Wednesday':'水','Thursday':'木',
  'Friday':'金','Saturday':'土','Sunday':'日',
  'Mon':'月','Tue':'火','Wed':'水','Thu':'木','Fri':'金','Sat':'土','Sun':'日'
};

const ALL_DAYS = ['月','火','水','木','金','土','日'];

const HOLIDAY_NOISE_PATTERNS = [
  /お問い?合わせ(ください|下さい)?/g, /詳細はお電話(にて)?/g,
  /コロナ.*?(\n|$)/g, /感染症.*?(\n|$)/g, /変更(に)?なる場合.*?(\n|$)/g,
  /変更の可能性.*?(\n|$)/g, /ご確認ください/g, /店舗にお問い?合わせ/g,
  /予告なく.*?(\n|$)/g, /※.*?(\n|$)/g, /\(※.*?\)/g, /（※.*?）/g,
  /毎週/g, /隔週/g, /第[一二三四五1-5]・?/g, /営業時間.*?(\n|$)/g,
  /ご来店前.*?(\n|$)/g, /店舗にご確認.*?(\n|$)/g, /変更となる場合.*?(\n|$)/g,
];

function normalizeDayText(text) {
  if (!text) return '';
  let result = text;
  for (const [long, short] of Object.entries(DAY_MAP)) {
    result = result.replaceAll(long, short);
  }
  return result;
}

function extractDaySet(text) {
  if (!text) return new Set();
  const normalized = normalizeDayText(text);
  const daySet = new Set();
  const rangePattern = /([月火水木金土日])[〜～~－\-ーー–—]([月火水木金土日])/g;
  let m;
  while ((m = rangePattern.exec(normalized)) !== null) {
    const start = ALL_DAYS.indexOf(m[1]);
    const end   = ALL_DAYS.indexOf(m[2]);
    if (start !== -1 && end !== -1) {
      for (let i = start; i <= end; i++) daySet.add(ALL_DAYS[i]);
    }
  }
  const withoutRange = normalized.replace(/[月火水木金土日][〜～~－\-ーー–—][月火水木金土日]/g, '  ');
  const listPattern = /[月火水木金土日]/g;
  let m2;
  while ((m2 = listPattern.exec(withoutRange)) !== null) {
    daySet.add(m2[0]);
  }
  return daySet;
}

function cleanHolidayText(text) {
  if (!text) return '';
  let result = text;
  for (const pattern of HOLIDAY_NOISE_PATTERNS) {
    result = result.replace(pattern, '');
  }
  return result.trim();
}

function resolveFinalHoliday(rawHolidayText, businessDaySet) {
  const cleaned = cleanHolidayText(rawHolidayText);
  if (/無休|年中無休/.test(cleaned)) return '無休';
  if (/^[\-－ー\s]+$/.test(cleaned)) return '-';
  const holidayDaySet = extractDaySet(cleaned);
  if (holidayDaySet.size > 0) return ALL_DAYS.filter(d => holidayDaySet.has(d)).join('・');
  if (businessDaySet && businessDaySet.size > 0) {
    const calculated = ALL_DAYS.filter(d => !businessDaySet.has(d));
    if (calculated.length === 0 || businessDaySet.size === 7) return '無休';
    if (calculated.length > 0 && calculated.length < 7) return calculated.join('・');
  }
  return '';
}

function normalizeBusinessHours(hoursText) {
  if (!hoursText) {
    return { holiday: '', businessDays: '', openTimeA: '', closeTimeA: '', openTimeB: '', closeTimeB: '' };
  }

  const textFlat = hoursText.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();

  let rawHolidayText = '';
  if (/年中無休|無休/.test(textFlat)) {
    rawHolidayText = '無休';
  } else {
    const holidayMatch = textFlat.match(/【定休日】([^【]+)/);
    if (holidayMatch) rawHolidayText = holidayMatch[1].trim();
    if (!rawHolidayText) {
      const squareMatch = textFlat.match(/■定休日\s*[\r\n]+([^\r\n■【]+)/);
      if (squareMatch) rawHolidayText = squareMatch[1].trim();
    }
    if (!rawHolidayText) {
      const colonMatch = textFlat.match(/定休日[：:]([^【■\r\n]+)/);
      if (colonMatch) rawHolidayText = colonMatch[1].trim();
    }
  }

  let hoursBlock = '';
  const hoursMatch = hoursText.match(/【営業時間】([\s\S]+?)(?:【定休日】|$)/);
  if (hoursMatch) {
    hoursBlock = hoursMatch[1];
  } else {
    hoursBlock = hoursText
      .replace(/【定休日】[\s\S]*/g, '')
      .replace(/■定休日[\s\S]*/g, '')
      .replace(/定休日[：:][\s\S]*/g, '');
  }

  const lines = hoursBlock.split('\n').map(l => l.trim()).filter(Boolean);
  const dayToTimes = Array.from({ length: 7 }, () => []);
  let currentDays = [...ALL_DAYS];
  
  const timeRangePattern = /(\d{1,2})[：:](\d{2})\s*[〜～\-–―]\s*(\d{1,2})[：:](\d{2})/g;

  for (const line of lines) {
    const hasDayChar = /[月火水木金土日]/.test(line);
    const timeMatches = [...line.matchAll(/(\d{1,2})[：:](\d{2})/g)];

    if (hasDayChar && timeMatches.length === 0) {
      const extractedDays = extractDaySet(line);
      if (extractedDays.size > 0) {
        currentDays = [...extractedDays];
      }
    } else if (timeMatches.length >= 2) {
      let match;
      timeRangePattern.lastIndex = 0;
      const pairs = [];
      while ((match = timeRangePattern.exec(line)) !== null) {
        pairs.push({
          open: `${match[1].padStart(2, '0')}:${match[2]}`,
          close: `${match[3].padStart(2, '0')}:${match[4]}`
        });
      }
      if (pairs.length === 0 && timeMatches.length >= 2) {
        for (let idx = 0; idx < timeMatches.length - 1; idx += 2) {
          pairs.push({
            open: `${timeMatches[idx][1].padStart(2, '0')}:${timeMatches[idx][2]}`,
            close: `${timeMatches[idx+1][1].padStart(2, '0')}:${timeMatches[idx+1][2]}`
          });
        }
      }
      if (pairs.length > 0) {
        currentDays.forEach(d => {
          const dayIdx = ALL_DAYS.indexOf(d);
          if (dayIdx !== -1) dayToTimes[dayIdx].push(...pairs);
        });
      }
    }
  }

  const filterUniqueTimes = (timesList) => {
    const seen = new Set();
    return timesList.filter(t => {
      const key = `${t.open}-${t.close}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const todayIdx = (new Date().getDay() + 6) % 7;
  let todayTimes = filterUniqueTimes(dayToTimes[todayIdx]);

  if (todayTimes.length <= 1) {
    let richerTimes = [];
    for (let i = 0; i < 7; i++) {
      const u = filterUniqueTimes(dayToTimes[i]);
      if (u.length > richerTimes.length) richerTimes = u;
    }
    if (richerTimes.length <= 1) {
      const allPairs = [];
      timeRangePattern.lastIndex = 0;
      let match;
      while ((match = timeRangePattern.exec(hoursBlock)) !== null) {
        allPairs.push({
          open: `${match[1].padStart(2, '0')}:${match[2]}`,
          close: `${match[3].padStart(2, '0')}:${match[4]}`
        });
      }
      const uAll = filterUniqueTimes(allPairs);
      if (uAll.length > richerTimes.length) richerTimes = uAll;
    }
    if (richerTimes.length >= 2) {
      todayTimes = richerTimes;
    }
  }

  todayTimes.sort((a, b) => a.open.localeCompare(b.open));

  let finalOpenTimeA = '';
  let finalCloseTimeA = '';
  let finalOpenTimeB = '';
  let finalCloseTimeB = '';

  if (todayTimes.length > 0) {
    finalOpenTimeA = todayTimes[0].open;
    finalCloseTimeA = todayTimes[0].close;
    if (finalCloseTimeA === '24:00' || finalCloseTimeA === '00:00') finalCloseTimeA = '24:00';
  }
  if (todayTimes.length > 1) {
    finalOpenTimeB = todayTimes[1].open;
    finalCloseTimeB = todayTimes[1].close;
    if (finalCloseTimeB === '24:00' || finalCloseTimeB === '00:00') finalCloseTimeB = '24:00';
  }

  const activeDays = [];
  for (let i = 0; i < 7; i++) {
    if (dayToTimes[i].length > 0) activeDays.push(ALL_DAYS[i]);
  }
  
  const businessDaySet = new Set(activeDays);
  const finalHoliday = resolveFinalHoliday(rawHolidayText, businessDaySet);
  let finalBusinessDays = finalHoliday === '無休' ? '月・火・水・木・金・土・日' : activeDays.join('・');

  return { holiday: finalHoliday, businessDays: finalBusinessDays, openTimeA: finalOpenTimeA, closeTimeA: finalCloseTimeA, openTimeB: finalOpenTimeB, closeTimeB: finalCloseTimeB };
}

function tabelogGetLinks(doc, baseUrl) {
  const links = [];
  const primary = doc.querySelectorAll([
    '.list-rst__rst-name-target',
    '.js-rst-cassette-wrap .list-rst__name a',
    'a.list-rst__name-main',
    '.list-rst__name a[href]',
    '.list-rst__rst-name a[href]',
    '.list-rst__rst-name-main a[href]',
    '.rst-name a[href]',
    '.rst__name a[href]',
    'h3 a[href]',
    'a[href*="/A"][href*="/A"][href*="/1"]',
    '.rstname a[href]',
    'a[href*="/rstLst/"][href*="/dtl"]'
  ].join(', '));
  primary.forEach(a => {
    const rawHref = a.getAttribute('href') || '';
    const href = resolveUrl(rawHref, baseUrl).split('?')[0];
    if (isTabelogRestaurantUrl(href)) addUniqueLink(links, href);
  });
  if (links.length === 0) {
    doc.querySelectorAll('a[href]').forEach(a => {
      const rawHref = a.getAttribute('href') || '';
      const href = resolveUrl(rawHref, baseUrl).split('?')[0];
      if (isTabelogRestaurantUrl(href)) addUniqueLink(links, href);
    });
  }
  return links;
}

function tabelogGetNextUrl(doc, baseUrl, pageNum = 1) {
  const nextBtn = doc.querySelector('a.c-pagination__arrow--next')
    || doc.querySelector('.c-pagination__arrow--next a')
    || doc.querySelector('a[rel="next"]')
    || Array.from(doc.querySelectorAll('a[href]')).find(isLikelyNextAnchor);
  if (nextBtn && !nextBtn.classList.contains('is-disabled')) {
    const rawHref = nextBtn.getAttribute('href') || '';
    return resolveUrl(rawHref, baseUrl);
  }
  return deriveTabelogNextUrl(baseUrl, pageNum);
}

function hotpepperGetLinks(doc, baseUrl) {
  const links = [];
  const anchors = doc.querySelectorAll([
    '.shopDetailTop a',
    '.shopName a',
    'h3.shopName a',
    'a.shopDetailLink',
    '.list-cassette__unit a',
    'a[href*="/strJ"]'
  ].join(', '));
  anchors.forEach(a => {
    const rawHref = a.getAttribute('href') || '';
    let href = resolveUrl(rawHref, baseUrl).split('?')[0].split('#')[0];
    if (/^https?:\/\/(www\.)?hotpepper\.jp\/(strJ[A-Z0-9]+|A[A-Z0-9]+)\/?$/.test(href)) {
      if (!href.endsWith('/')) href += '/';
      addUniqueLink(links, href);
    }
  });
  if (links.length === 0) {
    doc.querySelectorAll('a[href]').forEach(a => {
      const rawHref = a.getAttribute('href') || '';
      let href = resolveUrl(rawHref, baseUrl).split('?')[0].split('#')[0];
      if (/^https?:\/\/(www\.)?hotpepper\.jp\/(strJ[A-Z0-9]+|A[A-Z0-9]+)\/?$/.test(href)) {
        if (!href.endsWith('/')) href += '/';
        addUniqueLink(links, href);
      }
    });
  }
  return links;
}

function hotpepperGetNextUrl(doc, baseUrl, pageNum = 1) {
  const pagerContainers = doc.querySelectorAll('.pageLinkLinearBasic, .pagination, .pager, .page-list, .pageList, .page-link');
  let nextBtn = null;
  for (const container of pagerContainers) {
    const anchors = Array.from(container.querySelectorAll('a'));
    nextBtn = anchors.find(isLikelyNextAnchor);
    if (nextBtn) break;
  }
  if (!nextBtn) {
    const anchors = Array.from(doc.querySelectorAll('a[href]'));
    nextBtn = anchors.find(isLikelyNextAnchor);
  }
  if (nextBtn) {
    const rawHref = nextBtn.getAttribute('href') || '';
    return resolveUrl(rawHref, baseUrl);
  }
  return deriveHotpepperNextUrl(baseUrl, pageNum);
}

function validateWebsiteUrl(url) {
  if (!url) return '無';
  const urlLower = url.toLowerCase().trim();
  const portalDomains = [
    'tabelog.com', 'hotpepper.jp', 'gorp.jp', 'gnavi.co.jp', 'retty.me',
    'favy.jp', 'favy.me', 'facebook.com', 'instagram.com', 'twitter.com', 'x.com', 'ameblo.jp'
  ];
  const isPortal = portalDomains.some(domain => urlLower.includes(domain));
  return (isPortal || urlLower === '') ? '無' : '有';
}

async function fetchAndParseDetail(link, siteType, context = {}, timeoutMs = 10000, signal = null) {
  try {
    const controller = new AbortController();
    const combinedSignal = signal || controller.signal;
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    const res = await fetch(link, { signal: combinedSignal });
    clearTimeout(timer);

    if (res.status === 403 || res.status === 429) {
      return { isBlocked: true, url: link };
    }

    const html = await res.text();

    // 食べログ・ホットペッパー等のWAF（Cloudflare等）ブロック検知
    if (/アクセスが拒否されました|一時的に制限|アクセスが集中|Cloudflare|Robot Check|Security Check/i.test(html)) {
      return { isBlocked: true, url: link };
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    let name = '', genre = '', address = '', phone = '', hasWebsite = '無', rawHours = '';

    if (siteType === 'tabelog') {
      name = doc.querySelector('.display-name, h1.display-name, h2.display-name, [property="v:itemreviewed"], h1')?.textContent?.trim()
        || doc.title.split('|')[0].trim();
      address = doc.querySelector('p.rstinfo-table__address, .rstinfo-table__address, [rel="v:addr"], [property="v:address"]')?.textContent?.trim() || '';

      let tHours = '';
      let tClosed = '';

      const businessItems = doc.querySelectorAll('.rstinfo-table__business-item');
      if (businessItems.length > 0) {
        const hoursArray = [];
        const closedArray = [];
        businessItems.forEach(item => {
          const txt = _extractCellLines(item);
          if (!txt) return;
          if (txt.includes('定休日')) {
            const closedDayMatch = txt.match(/^([月火水木金土日・\s]+)定休日/);
            if (closedDayMatch) closedArray.push(closedDayMatch[1].trim());
            else closedArray.push(txt);
          } else {
            hoursArray.push(txt);
          }
        });
        tHours = hoursArray.join('\n');
        if (closedArray.length > 0) tClosed = closedArray.join('\n');
      }

      if (!tClosed) {
        const allBusinessText = Array.from(doc.querySelectorAll('.rstinfo-table__business-item')).map(el => el.textContent).join(' ');
        if (/年中無休|無休/.test(allBusinessText)) tClosed = '無休';
      }

      // 電話番号の退避・分別用
      let tclPhone = ''; // お問い合せ専用番号 / 電話番号 (実店舗最優先)
      let rsvPhone = ''; // 予約・お問い合わせ / ネット予約用

      doc.querySelectorAll('.rstinfo-table__table th, table th').forEach(th => {
        const t = th.textContent.trim();
        const tdText = th.nextElementSibling?.textContent || '';
        if (t.includes('ジャンル')) genre = th.nextElementSibling?.textContent?.trim() || genre;
        if (t.includes('住所') && !address) address = th.nextElementSibling?.textContent?.trim() || '';
        
        // 【改修】食べログの電話番号属性切り分け
        if (t.includes('お問い合せ専用番号') || t.includes('お問い合わせ専用') || t === '電話番号') {
          tclPhone = tdText;
        } else if (t.includes('予約') || t.includes('お問い合わせ')) {
          if (!rsvPhone) rsvPhone = tdText;
        }
      });

      if (tClosed && /年中無休/.test(tClosed)) tClosed = '無休';

      // ページ内に元々埋め込まれている基本テキスト
      const defaultTelText = doc.querySelector('.rstinfo-table__tel-num')?.textContent || '';
      
      // お問い合せ専用(tclPhone) → 一般情報(default) → 予約用(rsvPhone) の順でテキストを結合して精査
      const combinedPhoneText = `${tclPhone}\n${defaultTelText}\n${rsvPhone}`;
      phone = selectBestPhoneNumber(combinedPhoneText);

      if (!phone) {
        const telAnchor = doc.querySelector('a[href^="tel:"]');
        if (telAnchor) phone = telAnchor.getAttribute('href').replace('tel:', '').trim().replace(/[^\d\-]/g, '');
      }

      const hpLink = doc.querySelector('.homepage a, a.c-link-arrow[href*="rst_site_url"], .rstinfo-table__link');
      if (hpLink) {
        hasWebsite = validateWebsiteUrl(hpLink.getAttribute('href'));
      } else {
        const hasHpText = doc.documentElement.textContent.includes('お店のホームページ');
        hasWebsite = hasHpText ? '有' : '無';
      }

      address = address.replace(/大きな地図を見る/g, '').replace(/周辺のお店を探す/g, '').replace(/\s+/g, ' ').trim();

      if (tHours) rawHours += `【営業時間】\n${tHours}\n`;
      if (tClosed) rawHours += `【定休日】\n${tClosed}`;
      rawHours = rawHours.trim();

    } else if (siteType === 'hotpepper') {
      const shopInner = doc.querySelector('.shopInner.meiryoFont') || doc.querySelector('.shopDetailInnerTop') || doc;
      name = shopInner.querySelector('.shopName')?.textContent?.trim()
        || doc.querySelector('h1')?.textContent?.trim()
        || doc.title.split('|')[0].trim();

      let businessHours = '';
      let regularHoliday = '';
      let hpPhoneText = '';

      doc.querySelectorAll('table tr').forEach(tr => {
        const cells = tr.querySelectorAll('td');
        if (cells.length < 2) return;
        const label = cells[0].textContent.trim();
        const valFirstLine = cells[1].textContent.trim().split('\n')[0].replace(/\s+/g, ' ').trim();
        const valFull = cells[1].textContent.replace(/\s+/g, ' ').trim();

        if (label === '店名' && (!name || name === doc.title.split('|')[0].trim())) name = valFirstLine || name;
        if (label === '住所' && !address) address = valFull;
        if (label === '電話' || label === 'TEL') hpPhoneText += `\n${cells[1].textContent}`;
        if ((label === 'ジャンル' || label === '料理') && !genre) genre = valFull;
        if (label === '営業時間' && !businessHours) businessHours = _extractCellLines(cells[1]);
        if (label === '定休日' && !regularHoliday) regularHoliday = _extractCellLines(cells[1]);
      });

      shopInner.querySelectorAll('th').forEach(th => {
        const t = th.textContent?.split('\n')[0].trim() || '';
        const td = th.nextElementSibling;
        if (!td) return;
        const tdText = td.textContent.trim().split('\n')[0].trim();
        const tdFull = td.textContent.replace(/\s+/g, ' ').trim();

        if (t.includes('店名') && (!name || name === doc.title.split('|')[0].trim())) name = tdText || name;
        if (t.includes('住所') && !address) address = tdFull || '';
        if (t.includes('電話') || t.includes('TEL') || t.includes('問い合わせ')) hpPhoneText += `\n${td.textContent}`;
        if ((t.includes('ジャンル') || t.includes('料理')) && !genre) genre = tdFull || '';
        if (t.includes('営業時間') && !businessHours) businessHours = _extractCellLines(td);
        if (t.includes('定休日') && !regularHoliday) regularHoliday = _extractCellLines(td);
      });

      if (!businessHours && !regularHoliday) {
        doc.querySelectorAll('dl').forEach(dl => {
          const dt = dl.querySelector('dt');
          const dd = dl.querySelector('dd');
          if (!dt || !dd) return;
          const t = dt.textContent.trim();
          if (t.includes('営業時間') && !businessHours) businessHours = _extractCellLines(dd);
          if (t.includes('定休日') && !regularHoliday) regularHoliday = _extractCellLines(dd);
          if (t.includes('住所') && !address) address = dd.textContent.replace(/\s+/g, ' ').trim();
          if ((t.includes('ジャンル') || t.includes('料理')) && !genre) genre = dd.textContent.replace(/\s+/g, ' ').trim();
        });
      }

      if (!address) address = shopInner.querySelector('.shopDetailInfoAddress')?.textContent?.trim() || shopInner.querySelector('.address')?.textContent?.trim() || '';
      
      // 【改修】溜まったテーブル内のテキスト群からベストな固定電話を優先抽出
      phone = selectBestPhoneNumber(hpPhoneText);
      if (!phone) phone = shopInner.querySelector('.shopDetailInfoTel')?.textContent?.trim() || shopInner.querySelector('.tel')?.textContent?.trim() || shopInner.querySelector('a[href^="tel:"]')?.textContent?.trim() || '';

      const telLinkNode = doc.querySelector('.telLink');
      if (telLinkNode || !phone || phone.includes('電話番号を表示する') || phone.startsWith('050')) {
        try {
          let telUrl = telLinkNode ? telLinkNode.getAttribute('href') : '';
          if (telUrl && !telUrl.startsWith('http')) {
            if (telUrl.startsWith('/')) telUrl = new URL(link).origin + telUrl;
            else telUrl = (link.endsWith('/') ? link.slice(0, -1) : link) + '/' + telUrl;
          }
          if (!telUrl) telUrl = (link.endsWith('/') ? link : link + '/') + 'tel/';
          await sleep(500);
          const telRes = await fetch(telUrl, { signal: signal || AbortSignal.timeout(8000) });
          const telDoc = new DOMParser().parseFromString(await telRes.text(), 'text/html');
          const telNode = telDoc.querySelector('.telephoneNumber, .tel, .telephone, a[href^="tel:"]');
          if (telNode) {
            let rawTel = telNode.textContent;
            let extractedSub = selectBestPhoneNumber(rawTel);
            if (extractedSub) {
              phone = extractedSub;
            } else if (telNode.tagName === 'A' && telNode.getAttribute('href')?.startsWith('tel:')) {
              phone = telNode.getAttribute('href').replace('tel:', '').trim();
            } else {
              phone = rawTel.replace(/[^\d\-]/g, '');
            }
          }
        } catch (e) { }
      }

      const hpLink = doc.querySelector('a[href^="http"]:not([href*="hotpepper"])');
      if (hpLink) {
        hasWebsite = validateWebsiteUrl(hpLink.getAttribute('href'));
      } else {
        const hasHpText = doc.documentElement.textContent.includes('お店のホームページ');
        hasWebsite = hasHpText ? '有' : '無';
      }

      address = address.replace(/地図を見る/g, '').replace(/\s+/g, ' ').replace(/\n/g, '').trim();
      phone = phone.replace(/[^\d\-]/g, '');
      name = name.replace(/\n/g, '').trim();
      genre = genre.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();

      if (businessHours) rawHours += `【営業時間】\n${businessHours}\n`;
      if (regularHoliday) rawHours += `【定休日】\n${regularHoliday}`;
      rawHours = rawHours.trim();
    }

    const parsedAddr = parseAddress(address);
    const searchArea = context.area || '';
    if (searchArea) {
      const cleanArea = searchArea.replace(/駅周辺|エリア|付近/g, '').trim();
      if (/[市区町村]$/.test(cleanArea) && parsedAddr.city && !parsedAddr.city.includes(cleanArea) && !cleanArea.includes(parsedAddr.city)) {
        return null; 
      }
    }

    return { name, genre, address, phone, rawHours, hasWebsite, url: link, source: siteType };

  } catch (e) {
    return null;
  }
}

async function runCrawlTask(tabId) {
  const task = activeTasks.get(tabId);
  if (!task) return;

  const isPopularSubtask = typeof tabId === 'string' && tabId.includes('_pg_');
  let collected = 0;
  let pageNum = 1;
  let currentListUrl = task.listUrl;

  try {
    while (task.running && collected < task.maxItems) {
      const siteType = getSiteType(currentListUrl);
      if (!siteType) {
        sendToBackground(tabId, 'ERROR', { message: '対応サイトではありません' });
        break;
      }

      const siteName = siteType === 'tabelog' ? '食べログ' : 'ホットペッパー';
      const genreLabel = task.metadata?.industry ? `["${task.metadata.industry}"]` : '';
      sendToBackground(tabId, 'PAGE_START', { page: pageNum, collected, siteName, genreLabel });

      await sleep(DELAY_LIST_FETCH);
      const res = await fetch(currentListUrl);
      const html = await res.text();
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');

      if (pageNum === 1) {
        const meta = extractMetadata(doc, siteType);
        if (meta.area || meta.industry) task.metadata = { ...task.metadata, ...meta };
        if (!task.metadata.area || !task.metadata.industry) {
          const parts = doc.title.split(' ');
          if (parts.length >= 2) {
            if (!task.metadata.area) task.metadata.area = parts[0];
            if (!task.metadata.industry) task.metadata.industry = parts[1];
          }
        }
      }

      const getLinks = siteType === 'tabelog' ? tabelogGetLinks : hotpepperGetLinks;
      let links = getLinks(doc, currentListUrl) || [];

      const existingUrls = new Set(task.results.filter(Boolean).map(r => r.url));
      links = links.filter(l => !existingUrls.has(l.split('?')[0]));

      const remaining = task.maxItems - collected;
      links = links.slice(0, remaining);

      if (links.length === 0) {
        sendToBackground(tabId, 'INFO', { message: `${genreLabel} ${pageNum}ページ目: 新規リンクなし → 終了` });
        break;
      }

      sendToBackground(tabId, 'INFO', { message: `📋 ${genreLabel} ${pageNum}ページ目: ${links.length}件を取得中...` });

      const tasksToRun = links.map((link, idx) => ({ link, originalIndex: collected + idx }));
      
      const siteConcurrency = siteType === 'tabelog' ? task.tabelogConcurrency : task.hotpepperConcurrency;
      const siteDelay = siteType === 'tabelog' ? task.tabelogDelay : task.hotpepperDelay;

      let activeConcurrency = siteConcurrency;
      let activeDelay = siteDelay;

      task.stats.activeConcurrency = activeConcurrency;
      task.stats.activeDelay = activeDelay;

      function sendProgress(latestName = '') {
        const currentCollected = task.results.filter(Boolean).length;
        const targetTabId = typeof tabId === 'string' && tabId.includes('_pg_')
          ? parseInt(tabId.split('_pg_')[0])
          : tabId;

        let successCount = task.stats.successCount;
        let failedCount = task.stats.failedCount;
        let retryingCount = task.stats.retryingCount;
        let collectedCount = currentCollected;

        if (typeof tabId === 'string' && tabId.includes('_pg_')) {
          const parentKey = targetTabId + '_popular';
          const parent = activeTasks.get(parentKey);
          if (parent) {
            successCount += parent.stats.successCount;
            failedCount += parent.stats.failedCount;
            collectedCount += parent.results.filter(Boolean).length;
          }
        }

        sendToBackground(targetTabId, 'PROGRESS', {
          collected: collectedCount,
          maxItems: task.maxItems,
          latest: latestName,
          page: pageNum,
          successCount: successCount,
          failedCount: failedCount,
          retryingCount: retryingCount,
          isThrottling: task.stats.isThrottling,
          activeConcurrency: task.stats.activeConcurrency,
          activeDelay: task.stats.activeDelay
        });
      }

      let currentIndex = 0;
      let activeCount = 0;
      const queuePromises = [];

      async function worker() {
        while (currentIndex < tasksToRun.length && task.running) {
          const currentCollected = task.results.filter(Boolean).length;
          if (currentCollected >= task.maxItems) break;

          const item = tasksToRun[currentIndex++];
          if (!item) break;

          activeCount++;
          task.stats.activeConcurrency = activeCount;

          const delay = task.stats.isThrottling ? 3000 : activeDelay;
          if (delay > 0) {
            await sleep(delay);
          }

          if (!task.running) {
            activeCount--;
            task.stats.activeConcurrency = activeCount;
            break;
          }

          try {
            let detail = null;
            let retries = 0;
            const maxRetriesVal = task.maxRetries ?? 2;

            while (retries <= maxRetriesVal && task.running) {
              if (retries > 0) {
                task.stats.retryingCount++;
                sendProgress();
                await sleep(retries * 1000);
                task.stats.retryingCount--;
                if (!task.running) break;
              }

              try {
                const timeoutMs = (task.fetchTimeout || 10) * 1000;
                detail = await fetchAndParseDetail(item.link, siteType, { area: task.metadata.area, listUrl: currentListUrl }, timeoutMs, task.abortController?.signal);

                if (detail && detail.isBlocked) {
                  throw new Error('BLOCKED');
                }
                break;
              } catch (err) {
                if (err.name === 'AbortError') {
                  console.warn(`[offscreen] タイムアウト: ${item.link}`);
                } else {
                  console.warn(`[offscreen] フェッチエラー: ${item.link}`, err.message);
                }
                retries++;
                if (retries > maxRetriesVal) {
                  throw err;
                }
              }
            }

            if (detail && detail.name && task.running) {
              const normalized = normalizeBusinessHours(detail.rawHours || '');
              const parsedAddr = parseAddress(detail.address);

              const finalDetail = {
                name: detail.name,
                genre: detail.genre,
                sourceGenre: detail.genre,
                prefecture: parsedAddr.prefecture,
                city: parsedAddr.city,
                address: detail.address,
                phone: detail.phone || '',
                regularHoliday: normalized.holiday || '無休',
                businessDays: normalized.businessDays || '',
                openTimeA: normalized.openTimeA || '',
                closeTimeA: normalized.closeTimeA || '',
                openTimeB: normalized.openTimeB || '',
                closeTimeB: normalized.closeTimeB || '',
                rawHours: detail.rawHours || '',
                url: detail.url,
                hasWebsite: detail.hasWebsite,
                source: detail.source === 'tabelog' ? '食べログ' : 'ホットペッパー',
                sourceUrl: currentListUrl,
                scrapedAt: new Date().toISOString()
              };

              task.results[item.originalIndex] = finalDetail;
              task.stats.successCount++;

              if (task.stats.isThrottling) {
                task.stats.isThrottling = false;
                sendToBackground(tabId, 'INFO', { message: '正常な通信を検知したため、通常速度に復帰します' });
              }

              collected = task.results.filter(Boolean).length;
              sendProgress(detail.name);
            } else {
              task.stats.failedCount++;
              sendProgress();
            }

          } catch (err) {
            console.error(`[offscreen] ${item.link} 最終失敗:`, err.message);
            task.stats.failedCount++;

            if (err.message === 'BLOCKED' || err.message.includes('403') || err.message.includes('429')) {
              if (!task.stats.isThrottling) {
                task.stats.isThrottling = true;
                sendToBackground(tabId, 'INFO', { message: '⚠️ アクセス拒否(403/429/ブロック画面)を検知しました。自動減速運転に移行します（並行数1、待機3000ms）' });
              }
            }
            sendProgress();
          } finally {
            activeCount--;
            task.stats.activeConcurrency = activeCount;
          }
        }
      }

      const workersCount = task.stats.isThrottling ? 1 : activeConcurrency;
      for (let w = 0; w < workersCount; w++) {
        queuePromises.push(worker());
      }
      await Promise.all(queuePromises);

      if (!task.running || collected >= task.maxItems) break;

      const getNextUrl = siteType === 'tabelog' ? tabelogGetNextUrl : hotpepperGetNextUrl;
      const nextUrl = getNextUrl(doc, currentListUrl, pageNum);
      if (!nextUrl) {
        sendToBackground(tabId, 'INFO', { message: `${genreLabel} 最終ページに達しました` });
        break;
      }

      currentListUrl = nextUrl;
      pageNum++;
    }
  } catch (err) {
    console.error('バックグラウンド処理エラー:', err);
    sendToBackground(tabId, 'ERROR', { message: err.message });
  } finally {
    task.running = false;

    // undefined をフィルタリングして順序を詰める
    task.results = task.results.filter(Boolean);

    if (!isPopularSubtask) {
      sendToBackground(tabId, 'DONE', {
        collected: task.results.length,
        results: task.results,
        metadata: task.metadata,
        successCount: task.stats.successCount,
        failedCount: task.stats.failedCount,
        retryingCount: task.stats.retryingCount,
        isThrottling: task.stats.isThrottling,
        activeConcurrency: 0,
        activeDelay: 0
      });
    }

    const mediaName = task.metadata.media === 'tabelog'
      ? '食べログ'
      : (task.metadata.media === 'hotpepper' ? 'ホットペッパー' : 'サイト');
    const area = task.metadata.area || '';
    const industry = task.metadata.industry || '';
    const count = task.results.length;

    let title = count === 0 ? '取得完了 (該当なし)' : '取得完了';
    let message = `${area} ${industry} (${mediaName}) の取得が完了しました。計 ${count} 件`;

    if (task.stopRequested) {
      title = '取得停止';
      message = `${area} ${industry} (${mediaName}) の取得を停止しました。計 ${count} 件取得済み`;
    } else if (collected >= task.maxItems) {
      title = '取得完了 (上限到達)';
    }

    if (!isPopularSubtask) {
      chrome.runtime.sendMessage({ target: 'background', type: 'SHOW_NOTIFICATION', title, message });
    }
    if (!isPopularSubtask && task.results.length > 0) {
      chrome.runtime.sendMessage({
        target: 'background',
        type: 'DOWNLOAD_CSV',
        results: task.results,
        metadata: task.metadata,
        tabId
      });
    }
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.target !== 'offscreen') return;

  const tabId = message.tabId;

  if (message.type === 'GENRE_LINKS_FROM_CONTENT_RESULT') {
    const resolver = genreLinksResolvers.get(message.tabId);
    if (resolver) {
      resolver(message.links || []);
      genreLinksResolvers.delete(message.tabId);
    }
    sendResponse && sendResponse({ ok: true });
    return true;
  }

  if (message.action === 'START_CRAWL') {
    if (activeTasks.get(tabId)?.running) {
      sendResponse({ ok: false, error: 'このタブで既に実行中です' });
      return;
    }
    const siteType = getSiteType(message.listUrl);
    const controller = new AbortController();
    activeTasks.set(tabId, {
      running: true,
      tabId,
      listUrl: message.listUrl,
      results: [],
      maxItems: message.maxItems || Infinity,
      metadata: { media: siteType, area: '', industry: '' },
      abortController: controller,
      tabelogConcurrency: message.tabelogConcurrency || 5,
      tabelogDelay: message.tabelogDelay || 800,
      hotpepperConcurrency: message.hotpepperConcurrency || 6,
      hotpepperDelay: message.hotpepperDelay || 500,
      maxRetries: message.maxRetries ?? 2,
      fetchTimeout: message.fetchTimeout || 10,
      stats: {
        successCount: 0,
        failedCount: 0,
        retryingCount: 0,
        isThrottling: false,
        activeConcurrency: 0,
        activeDelay: 0
      }
    });
    runCrawlTask(tabId);
    sendResponse({ ok: true });
    return;
  }

  if (message.action === 'STOP_CRAWL') {
    const task = activeTasks.get(tabId);
    if (task) {
      task.stopRequested = true;
      task.running = false;
      task.abortController?.abort();
    }
    const popularTask = activeTasks.get(tabId + '_popular');
    if (popularTask) {
      popularTask.stopRequested = true;
      popularTask.running = false;
      popularTask.abortController?.abort();
    }
    const popularPrefix = `${tabId}_pg_`;
    for (const [key, runningTask] of activeTasks.entries()) {
      if (typeof key === 'string' && key.startsWith(popularPrefix)) {
        runningTask.stopRequested = true;
        runningTask.running = false;
        runningTask.abortController?.abort();
      }
    }
    sendResponse({ ok: true });
    return;
  }

  if (message.action === 'GET_RESULTS') {
    const popularTask = activeTasks.get(tabId + '_popular');
    const task = popularTask || activeTasks.get(tabId);
    if (task) {
      sendResponse({
        results: task.results || [],
        running: task.running || false,
        metadata: task.metadata || {}
      });
    } else {
      sendResponse({ results: [], running: false, metadata: {} });
    }
    return;
  }

  if (message.action === 'START_POPULAR_GENRE_CRAWL') {
    if (activeTasks.get(tabId)?.running || activeTasks.get(tabId + '_popular')?.running) {
      sendResponse({ ok: false, error: 'このタブで既に実行中です' });
      return;
    }
    runPopularGenreCrawl(tabId, message.listUrl, message.maxItems || Infinity, {
      tabelogConcurrency: message.tabelogConcurrency,
      tabelogDelay: message.tabelogDelay,
      hotpepperConcurrency: message.hotpepperConcurrency,
      hotpepperDelay: message.hotpepperDelay,
      maxRetries: message.maxRetries,
      fetchTimeout: message.fetchTimeout
    });
    sendResponse({ ok: true });
    return;
  }
});

async function extractGenreLinks(listUrl, siteType, tabId) {
  if (tabId != null) {
    try {
      const liveLinks = await new Promise((resolve) => {
        const timer = setTimeout(() => {
          genreLinksResolvers.delete(tabId);
          resolve([]);
        }, 5000);

        genreLinksResolvers.set(tabId, (links) => {
          clearTimeout(timer);
          resolve(links);
        });

        chrome.runtime.sendMessage({
          target: 'background',
          type: 'GET_GENRE_LINKS_FROM_CONTENT',
          tabId,
          siteType
        }).catch(() => {
          clearTimeout(timer);
          genreLinksResolvers.delete(tabId);
          resolve([]);
        });
      });

      if (liveLinks.length > 0) return liveLinks;
    } catch (e) {
      console.warn('[extractGenreLinks] ライブDOM問い合わせ失敗:', e);
    }
  }

  try {
    const res = await fetch(listUrl);
    const html = await res.text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const links = [];

    if (siteType === 'tabelog') {
      const collectTabelogGenreLinks = (selector) => {
        doc.querySelectorAll(selector).forEach(a => {
          const href = resolveUrl(a.getAttribute('href') || '', listUrl).split('?')[0].split('#')[0];
          const name = a.textContent.trim().replace(/\s+/g, ' ');
          if (href && name && /tabelog\.com/.test(href) && !links.some(l => l.url === href)) {
            links.push({ name, url: href });
          }
        });
      };

      [
        '#js-leftnavi-genre-scroll .list-balloon__btn-list a[href]',
        '.list-balloon__btn-list a[href]'
      ].forEach(collectTabelogGenreLinks);

      if (links.length === 0) [
        '.list-sidebar__item-target a[href]',
        '.list-sidebar a[href*="/rstLst/"]',
        'a[href*="/rstLst/"]'
      ].forEach(collectTabelogGenreLinks);

    } else if (siteType === 'hotpepper') {
      doc.querySelectorAll('.reselectionList li a[href]').forEach(a => {
        const href = resolveUrl(a.getAttribute('href') || '', listUrl).split('?')[0].split('#')[0];
        const name = a.textContent.trim().replace(/\s+/g, ' ');
        if (!href || !name) return;
        if (!/hotpepper\.jp/.test(href)) return;
        if (!/\/G\d+/.test(href)) return;
        if (links.some(l => l.url === href)) return;
        links.push({ name, url: href });
      });
    }

    return links;
  } catch (e) {
    console.error('[extractGenreLinks] fetchフォールバック失敗:', e);
    return [];
  }
}

async function runPopularGenreCrawl(tabId, listUrl, maxItemsPerGenre, speedConfig = {}) {
  const siteType = getSiteType(listUrl);
  if (!siteType) {
    sendToBackground(tabId, 'ERROR', { message: '対応サイトではありません' });
    return;
  }

  const parentTaskKey = tabId + '_popular';
  activeTasks.set(parentTaskKey, {
    running: true,
    results: [],
    metadata: { media: siteType, area: '', industry: '人気ジャンル一括' },
    ...speedConfig,
    stats: {
      successCount: 0,
      failedCount: 0,
      retryingCount: 0,
      isThrottling: false,
      activeConcurrency: 0,
      activeDelay: 0
    }
  });

  try {
    sendToBackground(tabId, 'INFO', { message: 'ジャンルリンクを抽出中...' });
    let genreLinks = [];
    try {
      genreLinks = await extractGenreLinks(listUrl, siteType, tabId);
    } catch (e) {
      sendToBackground(tabId, 'ERROR', { message: `ジャンルリンク取得失敗: ${e.message}` });
      activeTasks.delete(parentTaskKey);
      return;
    }

    if (genreLinks.length === 0) {
      const reason = siteType === 'tabelog'
        ? 'ジャンルリンクが見つかりません。食べログの検索結果ページを開いた状態で実行してください。'
        : 'ジャンルリンクが見つかりません。ホットペッパーのエリアページを開いた状態で実行してください。';
      sendToBackground(tabId, 'ERROR', { message: reason });
      activeTasks.delete(parentTaskKey);
      return;
    }

    sendToBackground(tabId, 'INFO', {
      message: `${genreLinks.length}ジャンルを検出: ${genreLinks.map(g => g.name).join('、')}`
    });

    const allResults = [];

    for (let i = 0; i < genreLinks.length; i++) {
      const parentTask = activeTasks.get(parentTaskKey);
      if (!parentTask || !parentTask.running) {
        sendToBackground(tabId, 'INFO', { message: '停止リクエストにより中断しました' });
        break;
      }

      const { name, url } = genreLinks[i];
      sendToBackground(tabId, 'INFO', {
        message: `🏷️ [ジャンル ${i + 1}/${genreLinks.length}]「${name}」の取得を開始します`
      });

      const tempId = `${tabId}_pg_${i}`;
      const subController = new AbortController();
      activeTasks.set(tempId, {
        running: true,
        tabId,
        listUrl: url,
        results: [],
        maxItems: maxItemsPerGenre,
        metadata: { media: siteType, area: '', industry: name },
        abortController: subController,
        ...speedConfig,
        stats: {
          successCount: 0,
          failedCount: 0,
          retryingCount: 0,
          isThrottling: false,
          activeConcurrency: 0,
          activeDelay: 0
        }
      });

      await runCrawlTask(tempId);

      const finishedTask = activeTasks.get(tempId);
      if (finishedTask?.results?.length) {
        const taggedResults = finishedTask.results.map(r => ({
          ...r,
          sourceGenre: name
        }));
        allResults.push(...taggedResults);
      }
      if (finishedTask && parentTask) {
        parentTask.stats.successCount += finishedTask.stats.successCount || 0;
        parentTask.stats.failedCount += finishedTask.stats.failedCount || 0;
      }
      activeTasks.delete(tempId);

      sendToBackground(tabId, 'INFO', {
        message: `✅ [ジャンル ${i + 1}/${genreLinks.length}]「${name}」完了 → 累計 ${allResults.length} 件`
      });

      if (i < genreLinks.length - 1) {
        const pt = activeTasks.get(parentTaskKey);
        if (pt && pt.running) await sleep(2000);
      }
    }

    const pt = activeTasks.get(parentTaskKey);
    const cleanResults = allResults.filter(Boolean);
    if (pt) {
      pt.results = cleanResults;
      pt.running = false;
    }

    const metaArea = cleanResults[0]?.address?.replace(/\s+/g, '').slice(0, 6) || '';
    const finalMetadata = { media: siteType, area: metaArea, industry: '人気ジャンル一括' };

    sendToBackground(tabId, 'DONE', {
      collected: cleanResults.length,
      results: cleanResults,
      metadata: finalMetadata,
      successCount: cleanResults.length,
      failedCount: pt?.stats?.failedCount || 0,
      retryingCount: 0,
      isThrottling: false,
      activeConcurrency: 0,
      activeDelay: 0
    });

    if (cleanResults.length > 0) {
      chrome.runtime.sendMessage({
        target: 'background',
        type: 'DOWNLOAD_CSV',
        results: cleanResults,
        metadata: finalMetadata,
        tabId
      });
    }

    chrome.runtime.sendMessage({
      target: 'background',
      type: 'SHOW_NOTIFICATION',
      title: '人気ジャンル一括取得 完了',
      message: `計 ${allResults.length} 件取得しました`
    });

  } catch (err) {
    console.error('[runPopularGenreCrawl] エラー:', err);
    sendToBackground(tabId, 'ERROR', { message: `人気ジャンル一括取得エラー: ${err.message}` });
  } finally {
    activeTasks.delete(parentTaskKey);
  }
}

chrome.runtime.sendMessage({ target: 'background', type: 'OFFSCREEN_READY' }).catch(() => { });
