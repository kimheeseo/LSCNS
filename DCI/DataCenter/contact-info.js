
(() => {
  'use strict';

  const ID = 'bom-contact-info';

  function textOf(el) {
    return (el && (el.innerText || el.textContent) || '').replace(/\s+/g, ' ').trim();
  }

  function findResultPanel() {
    const candidates = Array.from(document.querySelectorAll('section, article, aside, div'))
      .filter((el) => {
        const t = textOf(el);
        return t.includes('설계 결과') &&
               (t.includes('READY') || t.includes('REVIEW')) &&
               t.includes('Leaf') &&
               t.includes('Spine');
      })
      .sort((a, b) => a.querySelectorAll('*').length - b.querySelectorAll('*').length);

    if (candidates.length) return candidates[0];

    const heading = Array.from(document.querySelectorAll('h1, h2, h3, h4, [role="heading"]'))
      .find((el) => textOf(el).includes('설계 결과'));

    if (!heading) return null;

    return heading.closest('.card, .panel, .result, .results, section, article, aside') ||
           heading.parentElement;
  }

  function mountContactInfo() {
    if (document.getElementById(ID)) return true;

    const panel = findResultPanel();
    if (!panel) return false;

    const box = document.createElement('div');
    box.id = ID;
    box.setAttribute('role', 'note');
    box.innerHTML =
      '<p class="bom-contact-title">본 툴은 DC BOM Design Tool입니다.</p>' +
      '<p class="bom-contact-line">관련 문의사항은 <a href="mailto:harrykim9463@gmail.com">harrykim9463@gmail.com</a>으로 연락해 주세요.</p>';

    panel.appendChild(box);
    return true;
  }

  function start() {
    if (mountContactInfo()) return;

    const observer = new MutationObserver(() => {
      if (mountContactInfo()) observer.disconnect();
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });

    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (mountContactInfo() || attempts >= 20) {
        clearInterval(timer);
        if (document.getElementById(ID)) observer.disconnect();
      }
    }, 250);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();
