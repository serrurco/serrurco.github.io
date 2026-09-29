'use strict';

// Fragments are not sent to the web server. Never fetch or log their contents.
// All destinations are fixed Outlook routes; no third-party scripts are loaded.
function renderOutlookLinks() {
  const args = new URLSearchParams(location.hash.slice(1));
  const rest = args.get('message') || '';
  const entry = args.get('entry') || '';
  const web = args.get('web') || '';
  const validRest = /^[A-Za-z0-9_+/-]{16,2048}={0,2}$/.test(rest);
  const validEntry = /^(?:[A-Fa-f0-9]{2}){16,2048}$/.test(entry);

  // A second message can reuse this page through a hash-only navigation.
  // Remove old targets first so invalid or incomplete links cannot retain them.
  for (const id of ['ios', 'windows', 'web']) {
    const a = document.getElementById(id);
    a.hidden = true;
    a.removeAttribute('href');
  }
  document.getElementById('pc-note').hidden = true;
  document.getElementById('intro').textContent = 'Choose the Outlook app on this device.';

  function enable(id, href) {
    const a = document.getElementById(id);
    a.href = href;
    a.hidden = false;
  }
  if (validRest) {
    enable('ios', 'ms-outlook://emails/message/open?restID=' + encodeURIComponent(rest));
  }
  if (validEntry) {
    enable('windows', 'outlook:' + entry.toUpperCase());
    document.getElementById('pc-note').hidden = false;
  }
  try {
    const fallback = new URL(web);
    const item = fallback.searchParams.get('ItemID') || '';
    if (['https://outlook.office365.com', 'https://outlook.office.com'].includes(fallback.origin)
        && fallback.pathname === '/owa/' && !fallback.username && !fallback.password
        && /^[A-Za-z0-9_+/-]{16,4096}={0,2}$/.test(item)) {
      // Keep only the read-message parameters; do not forward arbitrary queries.
      const safe = new URL('/owa/', fallback.origin);
      safe.searchParams.set('ItemID', item);
      safe.searchParams.set('exvsurl', '1');
      safe.searchParams.set('viewmodel', 'ReadMessageItem');
      enable('web', safe.href);
    }
  } catch (_) { /* No verified web fallback was provided. */ }
  if (!validRest && !validEntry) {
    document.getElementById('intro').textContent = 'This link is missing a valid message reference. Please request a fresh link.';
  }
}

window.addEventListener('hashchange', renderOutlookLinks);
renderOutlookLinks();
