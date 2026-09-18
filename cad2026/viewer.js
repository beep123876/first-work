'use strict';
(async function openFieldViewer() {
  const runningScript = document.currentScript;
  const root = new URL('./', runningScript.src);
  const mobile = document.documentElement.dataset.viewer === 'light';
  const source = new URL(mobile ? '../iseongsanseong-live/light/index.html' : '../iseongsanseong-live/index.html', root);
  const map = new URL('../iseongsanseong-live/base.jpg', root);
  const title = mobile ? '2026 하남이성산성문화제 현장용 배치도' : '2026 하남이성산성문화제 현장 배치도';
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 25000);
  try {
    const response = await fetch(source, { cache: 'no-store', credentials: 'same-origin', signal: controller.signal });
    if (!response.ok || new URL(response.url).origin !== location.origin) throw new Error('배치도 읽기 오류');
    let html = await response.text();
    clearTimeout(timer);
    if (!/<html[\s>]/i.test(html) || !/<script[\s>]/i.test(html)) throw new Error('배치도 형식 오류');
    // Keep the original viewer, geometry and event handlers; repair the image URL for project hosting.
    html = html.replace(/(["'])\/base\.jpg\1/g, function (_match, quote) { return quote + map.pathname + quote; });
    html = html.replace(/<title>[\s\S]*?<\/title>/i, '<title>' + title + '</title>');
    const safeBase = new URL('./', source).href.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
    const additions = '<base href="' + safeBase + '"><meta name="robots" content="noindex,nofollow,noarchive"><meta name="application-name" content="하남이성산성문화제 현장 배치도"><meta name="referrer" content="same-origin"><meta property="og:title" content="' + title + '"><meta property="og:description" content="행사장 시설 및 운영인력 배치 확인"><link rel="icon" href="data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 64 64%27%3E%3Crect width=%2764%27 height=%2764%27 rx=%2712%27 fill=%27%23253c4f%27/%3E%3Cpath d=%27M14 48V27l18-14 18 14v21M24 48V32h16v16%27 fill=%27none%27 stroke=%27white%27 stroke-width=%275%27/%3E%3C/svg%3E">';
    html = html.replace(/<head([^>]*)>/i, '<head$1>' + additions);
    document.open();
    document.write(html);
    document.close();
  } catch (error) {
    clearTimeout(timer);
    const status = document.getElementById('load-status');
    if (status) status.textContent = '배치도를 불러오지 못했습니다. 통신 상태를 확인한 뒤 새로고침해 주세요.';
    const retry = document.getElementById('retry');
    if (retry) retry.hidden = false;
    console.error('배치도 불러오기 실패', error);
  }
})();
