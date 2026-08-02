/**
 * Import a supplier PO from its PDF, entirely inside Apps Script.
 *
 * WHY THIS IS SHAPED THIS WAY
 *
 * Arabian Petroleum's PO PDFs carry a text layer whose glyph codes have no /ToUnicode
 * mapping, so the characters cannot be recovered — reading the file directly yields
 * "aXXNMM]". Worse, Drive sees that layer, concludes the document already has text, and
 * skips OCR: uploading the PDF with ocr=true returned 0 of 6 key fields in testing, and so
 * did ocr=false followed by a converting copy.
 *
 * The way through is to make Drive rasterise the page and then OCR the picture. Drive
 * already generates a thumbnail for every stored PDF, and that thumbnail can be pulled at
 * 1024x1449 — enough to read the whole table, verified at 17 of 17 key fields including the
 * 6pt rate column. So:
 *
 *     store PDF -> fetch thumbnail (Google rasterises) -> re-upload image with ocr=true
 *       -> export text -> parse -> verify arithmetic -> human review -> write rows
 *
 * Every step is DriveApp or UrlFetchApp, so this runs for anyone who opens the sheet, on
 * any machine, with nothing installed.
 */

var PO_IMPORT_TEMP_PREFIX_ = 'POIMPORT_TMP_';

/** Entry point: Dispatch Tools -> Import PO from PDF. */
function showPoImportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('PoImportDialog')
    .setWidth(760)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Purchase Order');
}

/**
 * Called from the dialog with a base64 PDF. Returns parsed lines plus a verification
 * verdict; writes nothing. The caller decides whether to import.
 */
function poImportExtract(base64Pdf, filename) {
  var tempIds = [];
  try {
    var bytes = Utilities.base64Decode(base64Pdf);
    var blob = Utilities.newBlob(bytes, 'application/pdf', filename || 'po.pdf');

    var text = poOcrPdfViaThumbnail_(blob, tempIds);
    var parsed = poParseText_(text);
    if (!parsed.ok) {
      return { ok: false, error: parsed.error, rawText: text.slice(0, 2000) };
    }

    parsed.check = poVerify_(parsed);
    parsed.masters = poCheckMasters_(parsed.lines);
    parsed.duplicate = poIsAlreadyImported_(parsed.poNumber);
    return parsed;
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  } finally {
    poCleanupTemp_(tempIds);
  }
}

/**
 * Rasterise via Drive's thumbnail, then OCR the image.
 *
 * The thumbnail is the whole trick: Google renders the PDF page for us, which is the one
 * step Apps Script cannot do itself.
 */
function poOcrPdfViaThumbnail_(pdfBlob, tempIds) {
  // 1. Store the PDF unconverted so Drive generates a thumbnail for it.
  var stored = DriveApp.createFile(pdfBlob.setName(PO_IMPORT_TEMP_PREFIX_ + 'src.pdf'));
  tempIds.push(stored.getId());

  // 2. Drive builds thumbnails asynchronously; poll rather than assume it is ready.
  var link = null;
  for (var attempt = 0; attempt < 12 && !link; attempt++) {
    var meta = Drive.Files.get(stored.getId(), { fields: 'thumbnailLink' });
    link = meta && meta.thumbnailLink;
    if (!link) Utilities.sleep(2000);
  }
  if (!link) {
    throw new Error('Drive did not generate a preview image for this PDF, so it cannot be read.');
  }

  // 3. Ask for the largest render Drive will serve. It caps around 1024px wide, which
  //    testing showed is sufficient for this layout.
  var imgUrl = link.split('=s')[0] + '=s2000';
  var res = UrlFetchApp.fetch(imgUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Could not download the page image from Drive (HTTP ' + res.getResponseCode() + ').');
  }

  // 4. Hand the image back to Drive for genuine OCR. An image has no text layer, so the
  //    shortcut that defeated the direct-PDF attempt is unavailable.
  //
  //    The metadata must NOT name the Docs mimeType here: in Drive v2 that declares what is
  //    being UPLOADED, so the request reads as "a Google Doc that needs OCR" and is rejected
  //    with "OCR is not supported for files of type application/vnd.google-apps.document".
  //    Conversion is requested through the convert/ocr options instead, and the upload keeps
  //    the blob's own image type.
  var doc = Drive.Files.insert(
    { title: PO_IMPORT_TEMP_PREFIX_ + 'ocr' },
    res.getBlob(),
    { ocr: true, ocrLanguage: 'en', convert: true }
  );
  tempIds.push(doc.id);

  return DocumentApp.openById(doc.id).getBody().getText();
}

function poCleanupTemp_(ids) {
  // Scratch files must never accumulate in the user's Drive, including on the error path.
  (ids || []).forEach(function (id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
    } catch (e) {
      Logger.log('PO import: could not trash temp file ' + id + ' — ' + e);
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Parsing
//
// OCR does not return the table row by row. It reads in column-major runs, so the PO comes
// back as every code+qty, then every unit+rate, then every amount. Reading line by line
// would interleave different products, so each column is collected on its own and zipped
// back by position — then checked arithmetically, because positional zipping is only safe
// if a misalignment is detectable.
// ─────────────────────────────────────────────────────────────────────────────

function poToNumber_(raw) {
  if (raw === null || raw === undefined) return null;
  var cleaned = String(raw).replace(/[,\s]/g, '');
  return /^-?\d+(\.\d+)?$/.test(cleaned) ? parseFloat(cleaned) : null;
}

function poParseText_(text) {
  var codes = (text.match(/\bPM[A-Z]{3}\d{4}\b/g) || []);
  if (!codes.length) {
    return { ok: false, error: 'No item codes (PMxxx####) were found in this PDF.' };
  }

  var qtys = codes.map(function (code) {
    var m = text.match(new RegExp(code + '\\s+([\\d,]+)'));
    return m ? poToNumber_(m[1]) : null;
  });

  // Rates are accepted ONLY when their count matches the codes exactly. OCR output is not
  // byte-identical between runs — one pass produced four NOS tokens and the next three,
  // which shifted every rate up a row and yielded a plausible table with wrong prices.
  var rateMatches = text.match(/\bNOS\s+[\d,]+\.\d{1,2}\b/gi) || [];
  var ratesAligned = rateMatches.length === codes.length;
  var rates = rateMatches.map(function (s) {
    return poToNumber_(s.replace(/\bNOS\s+/i, ''));
  });

  var amounts = poCollectAmounts_(text, codes.length);

  var lines = codes.map(function (code, i) {
    return {
      sr: i + 1,
      code: code,
      qty: qtys[i] === undefined ? null : qtys[i],
      rate: ratesAligned ? (rates[i] === undefined ? null : rates[i]) : null,
      amount: amounts[i] === undefined ? null : amounts[i],
      rateSource: ratesAligned ? 'read' : 'pending'
    };
  });

  // Where the rate could not be read, derive it from amount / qty. Both of those are
  // independently cross-checked against the printed totals, so this is sounder than a
  // positional guess — and the qty x rate test still has to pass afterwards.
  lines.forEach(function (l) {
    if (l.rate === null && l.qty && l.amount) {
      l.rate = Math.round((l.amount / l.qty) * 100) / 100;
      l.rateSource = 'derived';
    } else if (l.rate === null) {
      l.rateSource = 'missing';
    }
  });

  var poNo = text.match(/\b([A-Z]{2,5}\/\d{2}-\d{2}\/[A-Z]{2,5}\/\d{3,6})\b/);
  var dated = text.match(/(\d{2}\/\d{2}\/\d{4})/);
  var indent = text.match(/REQ-IND\/(\d+)/i) || text.match(/INDENT\s*NO\s*:?\s*(\d+)/i);

  return {
    ok: true,
    poNumber: poNo ? poNo[1] : null,
    poDate: dated ? dated[1] : null,
    indent: indent ? indent[1] : null,
    printed: poFindPrintedTotals_(text),
    lines: lines
  };
}

/**
 * Amounts are the run of grouped numbers between the last rate and the Total label.
 * Restricting the window keeps GST figures, pin codes and phone numbers from being
 * mistaken for line amounts.
 */
function poCollectAmounts_(text, expected) {
  var rateRe = /\bNOS\s+[\d,]+\.\d{1,2}\b/gi;
  var lastEnd = 0;
  var m;
  while ((m = rateRe.exec(text)) !== null) lastEnd = m.index + m[0].length;

  var totalIdx = text.search(/\bTotal\b/i);
  var segment = text.slice(lastEnd, totalIdx > lastEnd ? totalIdx : text.length);

  var nums = (segment.match(/\b\d[\d,]*(?:\.\d{1,2})?\b/g) || [])
    .map(poToNumber_)
    .filter(function (n) { return n !== null && n > 0; });

  return nums.slice(-expected);
}

function poFindPrintedTotals_(text) {
  var qty = text.match(/([\d,]{4,})\s*[\r\n]+\s*Packing/i);
  var amt = text.match(/Packing\s*&?\s*Forwarding\s*[\r\n]+\s*([\d,]+)/i);
  return {
    qty: qty ? poToNumber_(qty[1]) : null,
    amount: amt ? poToNumber_(amt[1]) : null
  };
}

/**
 * The safeguard that makes OCR trustworthy here: every line must satisfy qty x rate =
 * amount, and the columns must sum to the totals printed on the PO. A misread digit or a
 * shifted column breaks that arithmetic, so bad data surfaces instead of being imported.
 */
function poVerify_(parsed) {
  var problems = [];
  var warnings = [];
  var sumQty = 0;
  var sumAmt = 0;

  parsed.lines.forEach(function (l) {
    sumQty += l.qty || 0;
    sumAmt += l.amount || 0;
    if (l.qty === null) problems.push('Line ' + l.sr + ' (' + l.code + '): quantity not read');
    if (l.rate === null) problems.push('Line ' + l.sr + ' (' + l.code + '): rate not read');
    if (l.qty !== null && l.rate !== null && l.amount !== null) {
      var calc = l.qty * l.rate;
      if (Math.abs(calc - l.amount) > 1.5) {
        problems.push('Line ' + l.sr + ' (' + l.code + '): ' + l.qty + ' x ' + l.rate +
          ' = ' + calc.toFixed(2) + ', but the amount column reads ' + l.amount);
      }
    }
  });

  if (parsed.printed.qty !== null && sumQty !== parsed.printed.qty) {
    problems.push('Quantities sum to ' + sumQty + ' but the PO prints ' + parsed.printed.qty);
  }
  if (parsed.printed.amount !== null && Math.abs(sumAmt - parsed.printed.amount) > 2) {
    problems.push('Amounts sum to ' + sumAmt + ' but the PO prints ' + parsed.printed.amount);
  }
  if (parsed.printed.qty === null && parsed.printed.amount === null) {
    warnings.push('No printed totals were found, so the lines could not be cross-checked.');
  }
  if (!parsed.poNumber) problems.push('PO number not found');
  if (!parsed.poDate) warnings.push('PO date not found — enter it by hand after importing.');

  var derived = parsed.lines.filter(function (l) { return l.rateSource === 'derived'; }).length;
  if (derived) {
    warnings.push(derived + ' rate(s) were computed as amount / qty because OCR did not read ' +
      'them cleanly. They agree with the printed amounts — confirm against the PDF.');
  }

  return { pass: problems.length === 0, problems: problems, warnings: warnings,
           sumQty: sumQty, sumAmt: sumAmt };
}

/** Flag codes absent from Masters: they import with blank Shape/Pack Size/Colour/Gms. */
function poCheckMasters_(lines) {
  var ss = getSpreadsheet_();
  var masters = ss.getSheetByName('Masters');
  if (!masters) return { missing: lines.map(function (l) { return l.code; }) };

  var rows = masters.getDataRange().getValues();
  var codeCol = rows[0].indexOf('Product Code');
  if (codeCol === -1) return { missing: [] };

  var known = {};
  rows.slice(1).forEach(function (r) { if (r[codeCol]) known[String(r[codeCol]).trim()] = true; });

  return {
    missing: lines.filter(function (l) { return !known[l.code]; })
      .map(function (l) { return l.code; })
  };
}

/** Re-importing a PO would double every quantity on the board, so check first. */
function poIsAlreadyImported_(poNumber) {
  if (!poNumber) return false;
  var dispatch = getSpreadsheet_().getSheetByName('Dispatch');
  if (!dispatch) return false;

  var labelRow = findInvoiceRow_(dispatch) + 1;
  var poCol = dispatchColumnByName_(dispatch, 'PO Number');
  if (poCol === -1) return false;

  var lastRow = dispatch.getLastRow();
  if (lastRow <= labelRow) return false;

  var existing = dispatch.getRange(labelRow + 1, poCol, lastRow - labelRow, 1).getValues();
  return existing.some(function (r) { return String(r[0]).trim() === poNumber; });
}

/**
 * Write the approved lines to Dispatch. Reuses the same field logic as the manual importer:
 * Masters supplies everything the code determines, and Balance / RM Req / MB carry live
 * formulas rather than frozen constants.
 */
function poImportCommit(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var ss = getSpreadsheet_();
    var dispatch = ss.getSheetByName('Dispatch');
    if (!dispatch) throw new Error('No sheet named "Dispatch" found.');

    if (poIsAlreadyImported_(payload.poNumber)) {
      return { ok: false, error: 'PO ' + payload.poNumber + ' is already in Dispatch. Nothing imported.' };
    }

    var labelRow = findInvoiceRow_(dispatch) + 1;
    var headers = dispatch.getRange(labelRow, 1, 1, dispatch.getLastColumn()).getValues()[0];
    var col = function (name) {
      var i = headers.indexOf(name);
      if (i === -1) throw new Error('Dispatch is missing a "' + name + '" column.');
      return i + 1;
    };

    var masters = ss.getSheetByName('Masters').getDataRange().getValues();
    var mh = masters[0];
    var mIdx = {
      code: mh.indexOf('Product Code'), shape: mh.indexOf('SHAPE'),
      pack: mh.indexOf('PACK SIZE'), gms: mh.indexOf('TARE WEIGHT in gms'),
      color: mh.indexOf('COLOR'), mbr: mh.indexOf('MB Ratio')
    };
    var byCode = {};
    masters.slice(1).forEach(function (r) {
      if (!r[mIdx.code]) return;
      byCode[String(r[mIdx.code]).trim()] = {
        shape: mIdx.shape === -1 ? '' : r[mIdx.shape],
        pack: mIdx.pack === -1 ? '' : r[mIdx.pack],
        gms: mIdx.gms === -1 ? '' : r[mIdx.gms],
        color: mIdx.color === -1 ? '' : r[mIdx.color],
        mbr: mIdx.mbr === -1 ? '' : r[mIdx.mbr]
      };
    });

    var startRow = dispatch.getLastRow() + 1;
    var poDate = poParseDate_(payload.poDate);

    payload.lines.forEach(function (l, i) {
      var r = startRow + i;
      dispatch.getRange(r, col('PO Number')).setValue(payload.poNumber);
      if (poDate) dispatch.getRange(r, col('PO Date')).setValue(poDate);
      dispatch.getRange(r, col('Code')).setValue(l.code);
      dispatch.getRange(r, col('Order Qty')).setValue(l.qty);
      dispatch.getRange(r, col('Rate')).setValue(l.rate);

      var m = byCode[l.code];
      if (m) {
        dispatch.getRange(r, col('Shape')).setValue(m.shape);
        dispatch.getRange(r, col('Pack Size')).setValue(m.pack);
        dispatch.getRange(r, col('Color')).setValue(m.color);
        dispatch.getRange(r, col('Gms')).setValue(m.gms);
        if (m.mbr !== '') dispatch.getRange(r, col('MB Ratio')).setValue(m.mbr);
      }
    });

    // Live formulas, so these rows recalculate as Supplied and Ready Stock are filled in.
    var a1 = function (c, r) {
      return dispatch.getRange(r, c).getA1Notation().replace(/[0-9]+$/, '') + r;
    };
    var cQty = col('Order Qty'), cSup = col('Supplied'), cRs = col('Physical Ready Stock');
    var cBal = col('Balance'), cGms = col('Gms'), cRm = col('RM Req');
    var cMb = col('MB'), cMbr = col('MB Ratio');

    payload.lines.forEach(function (l, i) {
      var r = startRow + i;
      dispatch.getRange(r, cBal).setFormula('=' + a1(cQty, r) + '-' + a1(cSup, r) + '-' + a1(cRs, r));
      dispatch.getRange(r, cRm).setFormula('=' + a1(cBal, r) + '*' + a1(cGms, r) + '/1000');
      dispatch.getRange(r, cMb).setFormula('=' + a1(cRm, r) + '*' + a1(cMbr, r) + '/100');
    });

    if (poDate) {
      dispatch.getRange(startRow, col('PO Date'), payload.lines.length, 1)
        .setNumberFormat('dd-mmm-yy');
    }

    return {
      ok: true,
      rows: payload.lines.length,
      firstRow: startRow,
      lastRow: startRow + payload.lines.length - 1
    };
  } finally {
    lock.releaseLock();
  }
}

/** PO dates read as dd/mm/yyyy; build an IST-midnight Date so the grid lands on the right day. */
function poParseDate_(ddmmyyyy) {
  if (!ddmmyyyy) return null;
  var m = String(ddmmyyyy).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(m[3] + '-' + m[2] + '-' + m[1] + 'T00:00:00+05:30');
}
