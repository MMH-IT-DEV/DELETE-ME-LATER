#!/usr/bin/env node
/**
 * sheet-scan.js — Read a Google Sheet tab from the terminal
 *
 * Usage:
 *   node scripts/sheet-scan.js --tab "Katana"
 *   node scripts/sheet-scan.js --tab "Adjustments Log"
 *   node scripts/sheet-scan.js --tab "Activity"
 *   node scripts/sheet-scan.js --list              (show all tab names)
 *   node scripts/sheet-scan.js --tab "Katana" --raw  (JSON output)
 *
 * Config: config/sheet-scan.json
 *   { "url": "https://script.google.com/macros/s/.../exec", "token": "your-secret" }
 */

'use strict';

var path = require('path');
var fs = require('fs');
var https = require('https');
var http = require('http');

// ── Load config ───────────────────────────────────────────────────────────────
var configPath = path.join(__dirname, '..', 'config', 'sheet-scan.json');
var config;
try {
  config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
} catch (e) {
  console.error('Config not found: ' + configPath);
  console.error('Create it with: { "url": "<web app URL>", "token": "<SCAN_TOKEN value>" }');
  process.exit(1);
}

if (!config.url || !config.token) {
  console.error('config/sheet-scan.json must have "url" and "token" fields.');
  process.exit(1);
}

// ── Parse args ────────────────────────────────────────────────────────────────
var args = process.argv.slice(2);
var tab = null;
var rawOutput = false;
var listMode = false;
var sheetArg = '';

for (var i = 0; i < args.length; i++) {
  if (args[i] === '--tab' && args[i + 1])   { tab = args[++i]; }
  if (args[i] === '--raw')                   { rawOutput = true; }
  if (args[i] === '--list')                  { listMode = true; }
  if (args[i] === '--sheet' && args[i + 1]) { sheetArg = args[++i]; }
}

if (!tab && !listMode) {
  console.error('Usage:');
  console.error('  node scripts/sheet-scan.js --tab "Tab Name" [--sheet sync|debug] [--raw]');
  console.error('  node scripts/sheet-scan.js --list [--sheet sync|debug]');
  process.exit(1);
}

// ── Pick endpoint based on --sheet ────────────────────────────────────────────
// sync  → engin-src GAS (bound to Sync Sheet, uses doGetScan)
// debug → debug GAS (bound to Debug Sheet, uses doGet)  [default]
var baseUrl = (sheetArg === 'sync' && config.syncUrl) ? config.syncUrl : config.url;

// ── Build request URL ─────────────────────────────────────────────────────────
var tabParam = listMode ? '__list__' : tab;
var requestUrl = baseUrl + '?token=' + encodeURIComponent(config.token) + '&tab=' + encodeURIComponent(tabParam);

// ── HTTP GET with redirect follow ─────────────────────────────────────────────
function get(reqUrl, redirectCount, callback) {
  if (redirectCount > 5) { return callback(new Error('Too many redirects')); }

  var parsed = new URL(reqUrl);
  var lib = parsed.protocol === 'https:' ? https : http;

  var options = {
    hostname: parsed.hostname,
    path: parsed.pathname + parsed.search,
    method: 'GET',
    headers: { 'User-Agent': 'sheet-scan/1.0' }
  };

  var req = lib.request(options, function(res) {
    // Follow redirects (GAS web apps always redirect)
    if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
      return get(res.headers.location, redirectCount + 1, callback);
    }

    var body = '';
    res.on('data', function(chunk) { body += chunk; });
    res.on('end', function() { callback(null, body, res.statusCode); });
  });

  req.on('error', callback);
  req.end();
}

// ── Run ───────────────────────────────────────────────────────────────────────
process.stderr.write('Fetching' + (listMode ? ' tab list' : ': "' + tab + '"') + ' ...\n\n');

get(requestUrl, 0, function(err, body, statusCode) {
  if (err) {
    console.error('Request failed:', err.message);
    process.exit(1);
  }

  var data;
  try {
    data = JSON.parse(body);
  } catch (e) {
    console.error('Could not parse response as JSON.');
    console.error('Raw response:', body.substring(0, 500));
    process.exit(1);
  }

  if (data.error) {
    console.error('Error from GAS (' + statusCode + '):', data.error);
    process.exit(1);
  }

  // ── List mode ──────────────────────────────────────────────────────────────
  if (listMode) {
    console.log('Tabs in spreadsheet ' + data.spreadsheetId + ':\n');
    for (var t = 0; t < data.tabs.length; t++) {
      console.log('  ' + (t + 1) + '. ' + data.tabs[t]);
    }
    console.log('');
    return;
  }

  var rows = data.rows || [];

  if (rows.length === 0) {
    console.log('(tab is empty)');
    return;
  }

  // ── Raw JSON output ────────────────────────────────────────────────────────
  if (rawOutput) {
    console.log(JSON.stringify(rows, null, 2));
    return;
  }

  // ── Pretty table ───────────────────────────────────────────────────────────
  var header = rows[0];
  var dataRows = rows.slice(1);

  // Calculate column widths (cap at 40)
  var colWidths = header.map(function(h, ci) {
    var w = String(h || '').length;
    dataRows.forEach(function(row) {
      var cell = String(row[ci] !== undefined ? row[ci] : '');
      if (cell.length > w) w = cell.length;
    });
    return Math.min(w, 40);
  });

  function pad(str, len) {
    str = String(str !== undefined ? str : '').substring(0, len);
    while (str.length < len) str += ' ';
    return str;
  }

  var separator = colWidths.map(function(w) { return '-'.repeat(w + 2); }).join('+');
  var headerRow = header.map(function(h, ci) { return ' ' + pad(h, colWidths[ci]) + ' '; }).join('|');

  console.log('Tab:  ' + tab);
  console.log('Rows: ' + dataRows.length + ' (+ 1 header)\n');
  console.log(separator);
  console.log(headerRow);
  console.log(separator);

  dataRows.forEach(function(row) {
    var line = header.map(function(_, ci) {
      return ' ' + pad(row[ci], colWidths[ci]) + ' ';
    }).join('|');
    console.log(line);
  });

  console.log(separator);
  console.log('\nDone.\n');
});
