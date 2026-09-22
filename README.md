# How to copy and paste between a data grid and Excel

Select a range of cells and press Ctrl+C: the grid puts the block on the system clipboard as
tab-separated text, the same form Excel, Numbers and Sheets all read as columns and rows, and this
page shows exactly what was copied. Copy a block from a real spreadsheet and press Ctrl+V over a
cell here and it pastes in, cell by cell, through each column's own type, tiling from the active
cell the way Excel's own paste does.

Live demo: https://toclocoinc.github.io/lattice-grid-howto-copy-and-paste-with-excel/

**Read the how-to:** https://www.latticegrid.dev/docs/how-to/copy-and-paste-with-excel/

## The full source

Two files: `index.html` loads the grid and declares the mount point, `demo.js` configures and creates it. Copy both as they are below and it runs.

### index.html

```html
<!doctype html>
<html lang="en-GB">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>How to copy and paste between a data grid and Excel</title>
    <meta
      name="description"
      content="Select a range of cells, copy it with Ctrl+C and see the tab-separated text Excel and Sheets read as columns and rows. Paste a block copied from Excel back in and it lands cell by cell, typed. Built with Lattice Grid loaded by script tag, no install and no build."
    />
    <link rel="icon" href="data:," />
    <!--
      The grid's stylesheet, from jsDelivr. The address names the exact
      release, 1.68.2, and carries the hash of the file it expects, so the
      page can never quietly pick up a different build than the one it was
      checked against.
    -->
    <link
      rel="stylesheet"
      href="https://cdn.jsdelivr.net/npm/@toclocoinc/lattice-grid@1.68.2/lattice-grid.min.css"
      integrity="sha384-mcpd7S8C5nz58bZDAXdYH6rzezEhfN7B4u2SlW426dSe20GnkxTu4TygyOILnCth"
      crossorigin="anonymous"
    />
    <style>
      body { margin: 0; font-family: system-ui, sans-serif; background: #f4f6f9; color: #131a24; }
      header { padding: 1.5rem 1.5rem 0.5rem; max-width: 960px; margin: 0 auto; }
      header p { color: #4a5568; }
      header a { color: #2d6bff; }
      main { max-width: 960px; margin: 0 auto; padding: 0 1.5rem 2.5rem; }
      #grid { height: 300px; margin: 0 0 1rem; }
      #stat { font-size: 0.9rem; color: #4a5568; margin: 0 0 1.5rem; }
      label { display: block; font-size: 0.9rem; font-weight: 600; margin: 0 0 0.4rem; }
      #copied {
        width: 100%; box-sizing: border-box; height: 110px; resize: vertical;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 0.85rem;
        padding: 0.6rem 0.75rem; border: 1px solid #ccd3dc; border-radius: 4px;
        background: #fff; color: #131a24; white-space: pre;
      }
      .caveats { font-size: 0.9rem; color: #4a5568; }
    </style>
  </head>
  <body>
    <header>
      <h1>How to copy and paste between a data grid and Excel</h1>
      <p>
        Drag across a block of cells, or extend it with Shift and the arrow keys, then press
        Ctrl+C (Cmd+C on a Mac). The grid puts the block on the clipboard as tab-separated text,
        the same form Excel, Numbers and Sheets all read as columns and rows, and the box below
        shows exactly what was copied. Copy a block from a real spreadsheet and press Ctrl+V over
        a cell here and it pastes in, cell by cell, through each column's own type. Read the
        <a href="https://www.latticegrid.dev/docs/how-to/copy-and-paste-with-excel/">full how-to</a>
        on latticegrid.dev.
      </p>
    </header>
    <main>
      <div id="grid"></div>
      <p id="stat"></p>
      <label for="copied">What was copied (tab-separated text)</label>
      <textarea id="copied" readonly></textarea>
      <p class="caveats">
        Two things to know: the browser asks permission before a page can read or write the
        system clipboard, so the first copy or paste may prompt; and a read-only cell is never a
        paste target, so pasting onto one changes nothing.
      </p>
    </main>

    <!--
      The library, as a classic script tag. No npm install, no bundler, no
      type="module": the file runs as it arrives and leaves the LatticeGrid
      global behind.
    -->
    <script
      src="https://cdn.jsdelivr.net/npm/@toclocoinc/lattice-grid@1.68.2/lattice-grid.min.js"
      integrity="sha384-vCzLyFYn0T0lz/vkdH4x0JpJZkOazZgI2LiGui7lm5uerdZd0Z46G9hr3Aq1FFPS"
      crossorigin="anonymous"
    ></script>
    <script src="./demo.js"></script>
  </body>
</html>
```

### demo.js

```js
/**
 * Copy a range with Ctrl+C and paste a block from Excel back in with Ctrl+V.
 *
 * Range selection is on by default, and it is what Ctrl+C and Ctrl+V act on:
 * the grid writes the selected rectangle to the system
 * clipboard as tab-separated values on copy, and on paste reads the same
 * shape back in, tiling it from the active cell through each column's own
 * type. The `clipboard:copy` event carries the exact text that was written,
 * which is what fills the box below so a reader can see the round trip
 * rather than take it on faith.
 */

// Tied to toclocoinc.github.io only; has no effect anywhere else and needs
// no key at all to run this page from a local copy.
LatticeGrid.setLicence(
  'LG1.eyJ2IjoxLCJwIjoibGF0dGljZS1ncmlkIiwidCI6IlRPQ0xPQ08gSW5jIC0gcHVibGljIGRlbW9zIiwiZSI6IjIwMzAtMDEtMDEiLCJkIjpbInRvY2xvY29pbmMuZ2l0aHViLmlvIl19.9De42ua3aCGpiMB6EVRP7Tv-upUlDI-0T07rlSPzvCrsqg8t4YJi7SRnStEpAg48uzmcG7il1fR_TfwkUE7iCA'
);

const PARTS = [
  { id: 'P01', item: 'Steel bolt M8', qty: 240, price: 0.18 },
  { id: 'P02', item: 'Steel bolt M10', qty: 180, price: 0.24 },
  { id: 'P03', item: 'Hex nut M8', qty: 240, price: 0.09 },
  { id: 'P04', item: 'Washer M8', qty: 480, price: 0.04 },
  { id: 'P05', item: 'Cable tie 200mm', qty: 600, price: 0.06 },
  { id: 'P06', item: 'Cable gland M20', qty: 90, price: 0.55 },
  { id: 'P07', item: 'Terminal block 4-way', qty: 60, price: 1.85 },
  { id: 'P08', item: 'Din rail 1m', qty: 30, price: 3.40 },
  { id: 'P09', item: 'Enclosure 200x150', qty: 12, price: 14.20 },
  { id: 'P10', item: 'Enclosure 300x200', qty: 8, price: 21.60 },
  { id: 'P11', item: 'Fuse 5A', qty: 150, price: 0.32 },
  { id: 'P12', item: 'Fuse holder', qty: 75, price: 0.95 },
];

const grid = LatticeGrid.createGrid(document.getElementById('grid'), {
  rowKey: 'id',
  selection: { mode: 'multiple', ranges: true },
  edit: true,
  columns: [
    { field: 'item', title: 'Item', layout: { flex: 1, min: 180 } },
    { field: 'qty', title: 'Qty', type: 'number', edit: true, layout: { width: 90 } },
    {
      field: 'price', title: 'Unit price', type: 'number', edit: true,
      format: { style: 'currency', currency: 'USD', decimals: 2 },
      layout: { width: 130 },
    },
  ],
  rows: PARTS,
});
window.__demoGrid = grid; // read by tools/verify.mjs

// A range already selected on load, so there is something to copy straight away.
grid.selection.setRange({ startRow: 0, endRow: 3, columns: ['qty', 'price'] });

const copied = document.getElementById('copied');
grid.on('clipboard:copy', (e) => {
  copied.value = e.ok ? e.text : 'the browser refused clipboard access';
});

document.getElementById('stat').textContent =
  'Drag across cells, or extend the selection with Shift and the arrow keys. ' +
  'Ctrl/Cmd+C copies the range; Ctrl/Cmd+V pastes a block copied from Excel or ' +
  'Sheets back into the grid, starting at the active cell.';
```

Range selection is on by default; Ctrl+C follows from it with no extra wiring, and Ctrl+V and
Ctrl+D (fill down) follow too once the target columns are editable. `grid.export.clipboard(opts)`
and `grid.export.rangeText(opts)`
do the same copy from code; `grid.edit.pasteInto(anchor, text, extent)` does the same paste. A
read-only cell is never a paste target: anchoring a paste on one changes nothing.

## Running it yourself

Open `index.html` in a browser, or serve the folder with any static file server. The grid loads
from jsDelivr by script tag, so there is no install and no build step. It runs keyless on
`localhost`; the licence key in `demo.js` is bound to `toclocoinc.github.io` and has no effect
anywhere else.

## Licence

MIT, see [LICENSE](./LICENSE). Lattice Grid itself is licensed separately per domain:
https://www.latticegrid.dev/pricing/
