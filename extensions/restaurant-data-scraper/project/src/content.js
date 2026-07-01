/**
 * content.js
 */

function addGenreLink(links, anchor) {
  const href = (anchor.href || '').split('?')[0].split('#')[0];
  const name = anchor.textContent.trim().replace(/\s+/g, ' ');
  if (!href || !name) return;
  if (!/tabelog\.com/.test(href)) return;
  if (links.some(l => l.url === href)) return;
  links.push({ name, url: href });
}

async function revealTabelogGenreBalloon() {
  const targets = Array.from(document.querySelectorAll('.list-sidebar__item-target'));
  const genreTarget = targets.find(el => /ジャンル|料理/.test(el.textContent || '')) || targets[0];
  if (!genreTarget) return;

  ['mouseover', 'mouseenter', 'mousemove'].forEach(type => {
    genreTarget.dispatchEvent(new MouseEvent(type, {
      bubbles: true,
      cancelable: true,
      view: window
    }));
  });

  await new Promise(resolve => setTimeout(resolve, 350));
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.action === 'PING') {
    sendResponse({ ok: true, url: window.location.href });
    return true;
  }

  if (message.action === 'GET_GENRE_LINKS') {
    (async () => {
      const siteType = message.siteType;
      const links = [];

      if (siteType === 'tabelog') {
        await revealTabelogGenreBalloon();

        const primarySelectors = [
          '#js-leftnavi-genre-scroll .list-balloon__btn-list a[href]',
          '.list-balloon__btn-list a[href]'
        ];
        primarySelectors.forEach(selector => {
          document.querySelectorAll(selector).forEach(a => addGenreLink(links, a));
        });

        if (links.length === 0) [
          '.list-sidebar__item-target a[href]',
          '.list-sidebar a[href*="/rstLst/"]',
          'a[href*="/rstLst/"]'
        ].forEach(selector => {
          document.querySelectorAll(selector).forEach(a => addGenreLink(links, a));
        });

      } else if (siteType === 'hotpepper') {
        // .jscDropDownSideInner.boxSide は複数存在し、最初はエリア用(.linkReselectionList)
        // ジャンル用(.reselectionList)は2番目以降にあるため、document全体から直接検索する
        document.querySelectorAll('.reselectionList li a[href]').forEach(a => {
          const href = a.href.split('?')[0].split('#')[0];
          const name = a.textContent.trim().replace(/\s+/g, ' ');
          if (!href || !name) return;
          if (!/hotpepper\.jp/.test(href)) return;
          if (!/\/G\d+/.test(href)) return;
          if (links.some(l => l.url === href)) return;
          links.push({ name, url: href });
        });
      }

      sendResponse({ links });
    })();
    return true;
  }
});
