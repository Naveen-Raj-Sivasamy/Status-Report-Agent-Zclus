// Shared UI chrome for every renderer page that adds `<script src="footer.js">`
// (currently index.html and manage.html — the two screens someone's
// actually working in day to day, as opposed to the one-time onboarding
// screens): a small brand mark dropped into whatever title element the
// page already has, and a persistent "Contact Admin" footer bar that
// expands in place into a short ticket form. Deliberately self-contained
// with inline styles rather than depending on either page's own CSS
// classes, since it's meant to look and behave identically wherever it's
// dropped in with the one script tag — see the Requester/User
// Affected/Issue/Explanation fields' matching _SupportTickets tab and
// submitAdminContact action in Code.gs (Resolution is admin-only, filled
// in later straight in the sheet, so it's not a field here).
(function () {
  const BRAND = '#6e1b2c';
  const BRAND_DARK = '#5a1522';
  const ACCENT = '#ffb700';

  // A document-with-a-checkmark-badge mark, in the app's own maroon +
  // amber (the same gold used by .amber-btn/.update-link elsewhere) —
  // deliberately not reusing fab.png/tray.png's unrelated blue-teal, so
  // the in-app corner mark actually matches the theme it sits on.
  const LOGO_SVG =
    '<svg viewBox="0 0 32 32" width="100%" height="100%" xmlns="http://www.w3.org/2000/svg">' +
    '<defs><linearGradient id="kcLogoGrad" x1="0" y1="0" x2="32" y2="32">' +
    '<stop offset="0" stop-color="' + BRAND + '"/><stop offset="1" stop-color="' + BRAND_DARK + '"/>' +
    '</linearGradient></defs>' +
    '<rect width="32" height="32" rx="8" fill="url(#kcLogoGrad)"/>' +
    '<rect x="8" y="6" width="13" height="17" rx="2" fill="#ffffff"/>' +
    '<rect x="10.5" y="10" width="8" height="1.6" rx="0.8" fill="' + BRAND + '" opacity="0.55"/>' +
    '<rect x="10.5" y="13.4" width="8" height="1.6" rx="0.8" fill="' + BRAND + '" opacity="0.55"/>' +
    '<rect x="10.5" y="16.8" width="5" height="1.6" rx="0.8" fill="' + BRAND + '" opacity="0.55"/>' +
    '<circle cx="23" cy="23" r="6.5" fill="' + ACCENT + '" stroke="' + BRAND + '" stroke-width="1"/>' +
    '<path d="M20.2 23.2l1.8 1.8 3.4-3.6" fill="none" stroke="#ffffff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>' +
    '</svg>';

  function injectLogo() {
    const titleEl = document.querySelector('.titlebar-title') || document.querySelector('header h1');
    if (!titleEl || titleEl.dataset.kcLogoDone) return;
    titleEl.dataset.kcLogoDone = 'true';
    const mark = document.createElement('span');
    mark.innerHTML = LOGO_SVG;
    mark.style.cssText = 'display:inline-flex; width:20px; height:20px; margin-right:8px; vertical-align:middle; flex-shrink:0;';
    titleEl.insertBefore(mark, titleEl.firstChild);
  }

  function field(labelText, opts) {
    opts = opts || {};
    const wrap = document.createElement('div');
    wrap.style.cssText = 'margin-bottom:8px; text-align:left;';
    const label = document.createElement('div');
    label.textContent = labelText + (opts.required ? ' *' : '');
    label.style.cssText = 'font-size:11px; font-weight:700; margin-bottom:3px; color:inherit;';
    wrap.appendChild(label);
    const input = document.createElement(opts.textarea ? 'textarea' : 'input');
    if (!opts.textarea) input.type = 'text';
    else input.rows = 3;
    input.style.cssText =
      'width:100%; box-sizing:border-box; padding:6px 8px; border-radius:6px; border:1px solid rgba(0,0,0,0.2);' +
      'font-family:inherit; font-size:12.5px; resize:vertical;';
    wrap.appendChild(input);
    return { wrap, input };
  }

  function injectFooter() {
    if (document.getElementById('kc-admin-contact-bar')) return;

    const bar = document.createElement('div');
    bar.id = 'kc-admin-contact-bar';
    bar.style.cssText =
      'flex-shrink:0; display:flex; align-items:center; justify-content:center; gap:6px;' +
      'padding:7px 10px; font-size:11px; background:rgba(0,0,0,0.15); color:#e8c9cf;' +
      '-webkit-app-region:no-drag;';
    const label = document.createElement('span');
    label.textContent = 'Something not working?';
    const link = document.createElement('button');
    link.type = 'button';
    link.id = 'kc-contact-admin-link';
    link.textContent = 'Contact Admin';
    link.style.cssText = 'all:unset; cursor:pointer; color:' + ACCENT + '; font-weight:700; text-decoration:underline;';
    bar.appendChild(label);
    bar.appendChild(link);

    const panel = document.createElement('div');
    panel.id = 'kc-admin-contact-panel';
    panel.style.cssText =
      'flex-shrink:0; padding:12px 16px 16px; background:rgba(0,0,0,0.12); color:#ffffff; text-align:center;' +
      '-webkit-app-region:no-drag;';
    panel.hidden = true;

    const requester = field('Requester', { required: true });
    const userAffected = field('User Affected (leave blank if same as you)');
    const issue = field('Issue', { required: true, textarea: true });
    const explanation = field('Explanation', { textarea: true });
    [requester, userAffected, issue, explanation].forEach((f) => panel.appendChild(f.wrap));

    const msgEl = document.createElement('div');
    msgEl.style.cssText = 'font-size:11px; margin:2px 0 8px; min-height:14px;';
    panel.appendChild(msgEl);

    const actions = document.createElement('div');
    actions.style.cssText = 'display:flex; gap:8px; justify-content:flex-end;';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText =
      'padding:6px 12px; border-radius:6px; border:1px solid rgba(255,255,255,0.4); background:transparent; color:#fff; cursor:pointer;';
    const sendBtn = document.createElement('button');
    sendBtn.type = 'button';
    sendBtn.textContent = 'Send to Admin';
    sendBtn.style.cssText =
      'padding:6px 12px; border-radius:6px; border:none; background:' + ACCENT + '; color:' + BRAND + '; font-weight:700; cursor:pointer;';
    actions.appendChild(cancelBtn);
    actions.appendChild(sendBtn);
    panel.appendChild(actions);

    function closePanel() {
      panel.hidden = true;
      msgEl.textContent = '';
      msgEl.style.color = '';
    }

    link.addEventListener('click', () => {
      panel.hidden = !panel.hidden;
      if (!panel.hidden && !requester.input.value && window.statusApp && window.statusApp.getYourName) {
        window.statusApp.getYourName().then((n) => {
          if (n && !requester.input.value) requester.input.value = n;
        }).catch(() => {});
      }
    });
    cancelBtn.addEventListener('click', closePanel);

    sendBtn.addEventListener('click', async () => {
      const payload = {
        requester: requester.input.value.trim(),
        userAffected: userAffected.input.value.trim(),
        issue: issue.input.value.trim(),
        explanation: explanation.input.value.trim(),
      };
      if (!payload.requester || !payload.issue) {
        msgEl.style.color = '#ffd7d7';
        msgEl.textContent = 'Requester and Issue are required.';
        return;
      }
      sendBtn.disabled = true;
      const original = sendBtn.textContent;
      sendBtn.textContent = 'Sending…';
      msgEl.style.color = '';
      msgEl.textContent = '';
      try {
        const result = await window.statusApp.submitAdminContact(payload);
        if (!result || result.ok === false) throw new Error((result && result.error) || 'Send failed.');
        msgEl.style.color = '#c8ffd8';
        msgEl.textContent = 'Sent — an admin will follow up.';
        issue.input.value = '';
        explanation.input.value = '';
        setTimeout(closePanel, 1400);
      } catch (err) {
        msgEl.style.color = '#ffd7d7';
        msgEl.textContent = err.message || String(err);
      } finally {
        sendBtn.disabled = false;
        sendBtn.textContent = original;
      }
    });

    // #card (index.html) is the element whose scrollHeight the OS window
    // is kept sized to — appending there, in normal flow, means the panel
    // opening/closing is picked up by that same ResizeObserver for free,
    // no separate floating-overlay/clipping logic needed. manage.html has
    // no #card (it's a plain scrollable popup), so .scroll-area/body there.
    const host = document.getElementById('card') || document.querySelector('.scroll-area') || document.body;
    host.appendChild(bar);
    host.appendChild(panel);
  }

  function init() {
    injectLogo();
    injectFooter();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
