#!/usr/bin/env node
'use strict';

/**
 * Read a supplier PO PDF and produce verified line items.
 *
 *   node import-po.js "path/to/PO.pdf"              print the parsed table
 *   node import-po.js "path/to/PO.pdf" --json       machine-readable output
 *   node import-po.js "path/to/PO.pdf" --gas        emit a PENDING_PO_IMPORT_ block
 *   node import-po.js "path/to/PO.pdf" --save-text  keep the raw OCR text alongside the PDF
 *
 * Pipeline: rasterise (locally, via pdfjs) -> OCR (Google Drive) -> parse -> verify.
 * Nothing is written to the sheet; this prints what it found so a human can check it.
 */

const fs = require('fs');
const path = require('path');

const { rasterisePdf } = require('./lib/rasterise');
const { ocrPages } = require('./lib/drive-ocr');
const { parsePoText, verify } = require('./lib/parse-po');

function fmt(n, dp = 0) {
  if (n === null || n === undefined) return '—';
  return n.toLocaleString('en-IN', { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

function printTable(parsed, check) {
  const W = 78;
  console.log('\n' + '═'.repeat(W));
  console.log(`  PO ${parsed.poNumber || '(number not read)'}`);
  console.log(`  Date ${parsed.poDate || '—'}    Indent ${parsed.indent || '—'}`);
  console.log('═'.repeat(W));
  console.log('  Sr  Item Code    Description                    Qty       Rate     Amount');
  console.log('  ' + '─'.repeat(W - 4));
  for (const l of parsed.lines) {
    console.log(
      '  ' +
        String(l.sr).padEnd(4) +
        l.code.padEnd(13) +
        (l.description || '').slice(0, 28).padEnd(30) +
        fmt(l.qty).padStart(8) +
        fmt(l.rate, 2).padStart(9) +
        fmt(l.amount).padStart(11)
    );
  }
  console.log('  ' + '─'.repeat(W - 4));
  console.log(
    '  ' + 'TOTAL'.padEnd(47) + fmt(check.sumQty).padStart(8) + ''.padStart(9) + fmt(check.sumAmt).padStart(11)
  );
  if (parsed.printed.qty !== null || parsed.printed.amount !== null) {
    console.log(
      '  ' +
        'printed on PO'.padEnd(47) +
        fmt(parsed.printed.qty).padStart(8) +
        ''.padStart(9) +
        fmt(parsed.printed.amount).padStart(11)
    );
  }
  console.log('═'.repeat(W));
}

function emitGasBlock(parsed) {
  const iso = (() => {
    if (!parsed.poDate) return 'YYYY-MM-DD';
    const [d, m, y] = parsed.poDate.split('/');
    return `${y}-${m}-${d}`;
  })();
  const lines = parsed.lines
    .map((l) => `    { code: '${l.code}', qty: ${l.qty}, rate: ${l.rate} },`)
    .join('\n');
  console.log('\n// Paste into clasp-jpermits/Code.js, replacing PENDING_PO_IMPORT_:');
  console.log('var PENDING_PO_IMPORT_ = {');
  console.log(`  poNumber: '${parsed.poNumber}',`);
  console.log(`  poDate: '${iso}',`);
  console.log('  lines: [');
  console.log(lines);
  console.log('  ],');
  console.log('};');
}

async function main() {
  const args = process.argv.slice(2);
  const pdfPath = args.find((a) => !a.startsWith('--'));
  if (!pdfPath) {
    console.error('Usage: node import-po.js "path/to/PO.pdf" [--json] [--gas] [--save-text]');
    process.exit(2);
  }
  if (!fs.existsSync(pdfPath)) {
    console.error(`File not found: ${pdfPath}`);
    process.exit(2);
  }

  const asJson = args.includes('--json');
  const log = asJson ? () => {} : (...a) => console.log(...a);

  log(`\nReading ${path.basename(pdfPath)}`);
  const pdf = fs.readFileSync(pdfPath);

  log('  rasterising at 300 DPI …');
  const pages = await rasterisePdf(pdf, {
    onPage: (n, total, bytes) => log(`    page ${n}/${total} → ${(bytes / 1024).toFixed(0)} KB`),
  });

  log('  running OCR via Google Drive …');
  const text = await ocrPages(pages, {
    onPage: (n, total, chars) => log(`    page ${n}/${total} → ${chars} chars`),
  });

  if (args.includes('--save-text')) {
    const out = pdfPath.replace(/\.pdf$/i, '.ocr.txt');
    fs.writeFileSync(out, text, 'utf8');
    log(`  raw OCR text → ${out}`);
  }

  const parsed = parsePoText(text);
  if (!parsed.ok) {
    console.error(`\nParse failed: ${parsed.error}`);
    console.error('Re-run with --save-text to inspect what OCR actually returned.');
    process.exit(1);
  }

  const check = verify(parsed);

  if (asJson) {
    console.log(JSON.stringify({ ...parsed, check }, null, 2));
    process.exit(check.pass ? 0 : 1);
  }

  printTable(parsed, check);

  for (const w of check.warnings) console.log(`  ⚠ ${w}`);
  if (check.pass) {
    console.log('\n  ✓ VERIFIED — line arithmetic and printed totals agree.');
  } else {
    console.log('\n  ✗ CHECK FAILED — do not enter these figures without reading the PDF:');
    for (const p of check.problems) console.log(`      • ${p}`);
  }

  if (args.includes('--gas')) emitGasBlock(parsed);

  console.log('');
  process.exit(check.pass ? 0 : 1);
}

main().catch((err) => {
  console.error('\nError:', err.message);
  process.exit(1);
});
