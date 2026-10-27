// HTML 슬라이드 → PDF(16:9, 장당 1페이지) + 슬라이드별 PNG (검수용)
// 사용법: node ppt-studio/tools/export_html.js <deck.html> [출력폴더]
// - 사전 설치된 Playwright Chromium 사용. 애니메이션은 끈 상태(최종 화면)로 캡처한다.
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

(async () => {
  const src = path.resolve(process.argv[2]);
  const out = path.resolve(process.argv[3] || path.join(path.dirname(src), 'render'));
  fs.mkdirSync(out, { recursive: true });
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.goto('file://' + src);
  await page.evaluate(() => document.fonts.ready);
  const n = await page.locator('.slide').count();
  for (let i = 1; i <= n; i++) {
    await page.evaluate((k) => {
      document.querySelectorAll('.slide').forEach((s, j) => s.classList.toggle('active', j === k - 1));
    }, i);
    await page.addStyleTag({ content: '.hud{display:none!important}' });
    await page.screenshot({ path: path.join(out, `slide-${String(i).padStart(2, '0')}.png`) });
  }
  const pdf = path.join(out, path.basename(src).replace(/\.html?$/, '') + '.pdf');
  await page.emulateMedia({ media: 'print', reducedMotion: 'reduce' });
  await page.pdf({ path: pdf, width: '1280px', height: '720px', printBackground: true, preferCSSPageSize: true });
  await browser.close();
  console.log(`PDF: ${pdf}\nPNG: ${n}장 → ${out}`);
})();
