'use strict';
var xlsx = require('xlsx');

var src = 'C:/Users/Admin/Downloads/_2026_Katana-Wasp Inventory Sync.xlsx';
var wb = xlsx.readFile(src);

// ── Read Katana tab → map SKU → { name, uom } ────────────────────────────────
var kRows = xlsx.utils.sheet_to_json(wb.Sheets['Katana'], { header: 1 });
var kMap = {};
for (var ki = 1; ki < kRows.length; ki++) {
  var kr = kRows[ki];
  var sku = String(kr[0] || '').trim();
  if (!sku || sku.indexOf(' > ') > -1) continue; // skip empty + separator rows
  if (kMap[sku]) continue; // first occurrence per SKU only
  kMap[sku] = {
    name:    String(kr[1] || '').trim(),
    katUom:  String(kr[7] || '').trim()
  };
}

// ── Read Wasp tab → map SKU → uom ────────────────────────────────────────────
var wRows = xlsx.utils.sheet_to_json(wb.Sheets['Wasp'], { header: 1 });
var wMap = {};
for (var wi = 1; wi < wRows.length; wi++) {
  var wr = wRows[wi];
  var wSku = String(wr[0] || '').trim();
  if (!wSku || wSku.indexOf(' > ') > -1) continue;
  if (wMap[wSku]) continue;
  wMap[wSku] = String(wr[8] || '').trim();
}

// ── Compare — only SKUs where UOM differs ────────────────────────────────────
var output = [['SKU', 'Name', 'Katana UOM', 'WASP UOM']];
var seen = {};

var skus = Object.keys(kMap);
for (var si = 0; si < skus.length; si++) {
  var s = skus[si];
  if (seen[s]) continue;
  seen[s] = true;

  var katUom  = kMap[s].katUom;
  var waspUom = wMap[s] || '(not in WASP)';

  if (katUom.toLowerCase() === waspUom.toLowerCase()) continue; // same — skip

  output.push([s, kMap[s].name, katUom, waspUom]);
}

// ── Write output xlsx ─────────────────────────────────────────────────────────
var outWb = xlsx.utils.book_new();
var outWs = xlsx.utils.aoa_to_sheet(output);

// Column widths
outWs['!cols'] = [
  { wch: 20 }, // SKU
  { wch: 35 }, // Name
  { wch: 15 }, // Katana UOM
  { wch: 15 }  // WASP UOM
];

xlsx.utils.book_append_sheet(outWb, outWs, 'UOM Comparison');

var outPath = 'C:/Users/Admin/Downloads/uom-comparison.xlsx';
xlsx.writeFile(outWb, outPath);

console.log('Done. ' + (output.length - 1) + ' mismatches written to:');
console.log(outPath);
