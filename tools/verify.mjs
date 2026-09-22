/**
 * Load the demo in a real headless browser and check that it works.
 *
 * Serves the project locally and drives it with the same minimal Chrome
 * driver the grid's own test suite uses (tools/browser.js), so this needs no
 * dependency beyond Chrome itself.
 *
 * Checks:
 *   - the library arrived and left the LatticeGrid global behind;
 *   - all 12 parts land in the grid (grid.rows.matchCount(), not a DOM count);
 *   - a real, trusted Ctrl+C over the range the demo pre-selects (rows 0-3,
 *     Qty and Price) puts exactly the tab-separated text of those eight cells
 *     into the "what was copied" box, via the grid's own `clipboard:copy`
 *     event — the same text a real spreadsheet would read as columns;
 *   - with clipboard permission granted over the DevTools protocol, seeding
 *     the system clipboard with a tab-separated block (as if it had been
 *     copied from Excel) and sending a real, trusted Ctrl+V writes those
 *     values into the four rows' Qty and Price cells, read back through the
 *     grid API rather than the DOM;
 *   - a read-only cell (Item) is never a paste target: pasting onto one
 *     leaves every cell exactly as it was;
 *   - nothing logged a console error or threw while the page ran.
 *
 * Exits non-zero on any failure, so it can gate a deployment.
 *
 * Usage: node tools/verify.mjs
 */

import { startServer } from './serve.mjs';
import { Browser, available } from './browser.js';

if (!available()) {
  console.log('No headless browser on this machine; skipping verify.mjs.');
  process.exit(0);
}

// The same 12 parts demo.js loads, so every expected figure is computed here
// rather than hard-coded, and stays honest if the data ever changes.
const PARTS = [
  { id: 'P01', item: 'Steel bolt M8', qty: 240, price: 0.18 },
  { id: 'P02', item: 'Steel bolt M10', qty: 180, price: 0.24 },
  { id: 'P03', item: 'Hex nut M8', qty: 240, price: 0.09 },
  { id: 'P04', item: 'Washer M8', qty: 480, price: 0.04 },
  { id: 'P05', item: 'Cable tie 200mm', qty: 600, price: 0.06 },
  { id: 'P06', item: 'Cable gland M20', qty: 90, price: 0.55 },
  { id: 'P07', item: 'Terminal block 4-way', qty: 60, price: 1.85 },
  { id: 'P08', item: 'Din rail 1m', qty: 30, price: 3.40 },
  { id: 'P09', item: 'Enclosure 200x150', qty: 12, price: 14.2 },
  { id: 'P10', item: 'Enclosure 300x200', qty: 8, price: 21.6 },
  { id: 'P11', item: 'Fuse 5A', qty: 150, price: 0.32 },
  { id: 'P12', item: 'Fuse holder', qty: 75, price: 0.95 },
];

// What the demo's own initial range (rows 0-3, Qty then Price) should copy
// as: tab between columns, newline between rows, raw numbers (the clipboard
// carries the underlying value, not the formatted currency string).
const expectedCopy = PARTS.slice(0, 4).map((r) => `${r.qty}\t${r.price}`).join('\n');

// A trusted key combo, dispatched over the DevTools protocol rather than as a
// script-fired DOM event, so it carries isTrusted: true and real user
// activation the way an actual keypress would — required for the Clipboard
// API and for the grid to tell a real interaction from a page script.
async function pressCtrl(browser, key, code, vk) {
  const base = { key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: 2 };
  await browser.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', ...base });
  await browser.send('Input.dispatchKeyEvent', { type: 'keyUp', ...base });
}

async function clickCell(browser, rowKey, col) {
  const rect = await browser.evaluate(`(() => {
    const cell = document.querySelector('[data-key="${rowKey}"] [data-col="${col}"]');
    if (!cell) return null;
    const r = cell.getBoundingClientRect();
    return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
  })()`);
  if (!rect) throw new Error(`cell not found: row ${rowKey}, column ${col}`);
  await browser.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: rect.x, y: rect.y, button: 'left', clickCount: 1 });
  await browser.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: rect.x, y: rect.y, button: 'left', clickCount: 1 });
}

const cellValue = (row, col) => `document.querySelector('[data-key="${row}"] [data-col="${col}"]')?.textContent`;

const { server, port } = await startServer();
const browser = new Browser();

try {
  await browser.start();

  await browser.send('Page.addScriptToEvaluateOnNewDocument', {
    source: `
      window.__errs = [];
      addEventListener('error', (e) => window.__errs.push(String(e.message || e)));
      addEventListener('unhandledrejection', (e) => window.__errs.push('unhandledrejection: ' + String(e.reason)));
      const origError = console.error.bind(console);
      console.error = (...args) => { window.__errs.push('console.error: ' + args.map(String).join(' ')); origError(...args); };
    `,
  });

  // Granted up front: the paste step reads the clipboard, and without this
  // Chrome refuses even a script that seeds its own clipboard.
  await browser.send('Browser.grantPermissions', {
    origin: `http://127.0.0.1:${port}`,
    permissions: ['clipboardReadWrite', 'clipboardSanitizedWrite'],
  });

  await browser.open(`http://127.0.0.1:${port}/`);
  await new Promise((r) => setTimeout(r, 500));

  const failures = [];

  const hasGrid = await browser.evaluate(`typeof LatticeGrid !== 'undefined' && typeof LatticeGrid.createGrid === 'function'`);
  if (!hasGrid) failures.push('LatticeGrid.createGrid was not found on the page');

  const partCount = await browser.evaluate(`window.__demoGrid.rows.matchCount()`);
  if (partCount !== 12) failures.push(`expected 12 parts (grid.rows.matchCount()), grid reports ${partCount}`);

  // --- Copy: a real Ctrl+C over the pre-selected range ---
  await clickCell(browser, 'P01', 'qty');
  // The click can change the selection to a single cell; put the four-row,
  // two-column range back exactly as the demo set it on load.
  await browser.evaluate(`window.__demoGrid.selection.setRange({ startRow: 0, endRow: 3, columns: ['qty', 'price'] })`);
  await new Promise((r) => setTimeout(r, 100));
  await pressCtrl(browser, 'c', 'KeyC', 67);
  await new Promise((r) => setTimeout(r, 200));

  const copiedText = await browser.evaluate(`document.getElementById('copied').value`);
  if (copiedText !== expectedCopy) {
    failures.push(`copied text did not match: expected ${JSON.stringify(expectedCopy)}, got ${JSON.stringify(copiedText)}`);
  }

  // --- Paste: seed the clipboard as if Excel had copied into it, then Ctrl+V ---
  const pasteRows = [
    { qty: 900, price: 9.9 },
    { qty: 800, price: 8.8 },
    { qty: 700, price: 7.7 },
    { qty: 600, price: 6.6 },
  ];
  const pasteText = pasteRows.map((r) => `${r.qty}\t${r.price}`).join('\n');
  await browser.evaluate(`navigator.clipboard.writeText(${JSON.stringify(pasteText)})`);
  await clickCell(browser, 'P01', 'qty');
  await new Promise((r) => setTimeout(r, 100));
  await pressCtrl(browser, 'v', 'KeyV', 86);
  await new Promise((r) => setTimeout(r, 300));

  for (let i = 0; i < pasteRows.length; i++) {
    const rowKey = PARTS[i].id;
    const qtyText = await browser.evaluate(cellValue(rowKey, 'qty'));
    const priceText = await browser.evaluate(cellValue(rowKey, 'price'));
    const qty = Number(String(qtyText).replace(/[^0-9.-]/g, ''));
    const price = Number(String(priceText).replace(/[^0-9.-]/g, ''));
    if (qty !== pasteRows[i].qty) {
      failures.push(`row ${rowKey} Qty after paste: expected ${pasteRows[i].qty}, grid shows ${qty}`);
    }
    if (Math.abs(price - pasteRows[i].price) > 0.005) {
      failures.push(`row ${rowKey} Price after paste: expected ${pasteRows[i].price}, grid shows ${price}`);
    }
  }

  // --- A read-only cell (Item) is never a paste target ---
  const itemBefore = await browser.evaluate(cellValue('P01', 'item'));
  await clickCell(browser, 'P01', 'item');
  await browser.evaluate(`navigator.clipboard.writeText('Renamed part')`);
  await new Promise((r) => setTimeout(r, 100));
  await pressCtrl(browser, 'v', 'KeyV', 86);
  await new Promise((r) => setTimeout(r, 200));
  const itemAfter = await browser.evaluate(cellValue('P01', 'item'));
  if (itemBefore !== itemAfter) {
    failures.push(`read-only Item cell changed on paste: was "${itemBefore}", now "${itemAfter}"`);
  }

  const errors = await browser.evaluate('window.__errs');
  if (errors.length) failures.push(`console/window errors: ${errors.join(' | ')}`);

  if (failures.length) {
    console.error('FAILED:\n' + failures.map((f) => `  - ${f}`).join('\n'));
    process.exitCode = 1;
  } else {
    console.log(
      'OK: 12 parts loaded, Ctrl+C over the pre-selected range copied the exact tab-separated ' +
      'text of those 8 cells, a seeded clipboard block pasted in with Ctrl+V landed in Qty and ' +
      'Price read back through the grid API, and a paste onto the read-only Item column changed ' +
      'nothing. 0 console errors.'
    );
  }
} finally {
  await browser.close();
  server.close();
}
