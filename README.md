# How to copy and paste between a data grid and Excel

Select a range of cells and press Ctrl+C: the grid puts the block on the system clipboard as
tab-separated text, the same form Excel, Numbers and Sheets all read as columns and rows, and this
page shows exactly what was copied. Copy a block from a real spreadsheet and press Ctrl+V over a
cell here and it pastes in, cell by cell, through each column's own type, tiling from the active
cell the way Excel's own paste does.

Live demo: https://toclocoinc.github.io/lattice-grid-howto-copy-and-paste-with-excel/

**Read the how-to:** https://www.latticegrid.dev/docs/how-to/copy-and-paste-with-excel/

## The snippet

```js
const grid = LatticeGrid.createGrid(document.getElementById('grid'), {
  rowKey: 'id',
  selection: { mode: 'multiple', ranges: true },
  edit: true,
  columns: [
    { field: 'item', title: 'Item' },
    { field: 'qty', title: 'Qty', type: 'number', edit: true },
    { field: 'price', title: 'Unit price', type: 'number', edit: true },
  ],
  rows, // twelve parts
});

grid.on('clipboard:copy', (e) => {
  console.log(e.text); // the tab-separated text that was copied
});
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
