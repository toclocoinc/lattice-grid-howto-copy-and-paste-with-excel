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
