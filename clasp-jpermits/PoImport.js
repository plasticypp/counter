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
  // A user who authorised the script before it used UrlFetchApp/Drive has not granted those
  // scopes, and google.script.run from a dialog never re-prompts: it just fails with "You do
  // not have permission to call UrlFetchApp.fetch". Catch that here and hand over the link.
  var auth = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (auth.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
    var prompt = HtmlService.createHtmlOutput(
      '<p style="font-family:Arial;font-size:14px">PO import needs one more Google permission ' +
      'for your account.</p><p style="font-family:Arial;font-size:14px"><a href="' +
      auth.getAuthorizationUrl() + '" target="_blank">Click here to allow it</a>, then run ' +
      '<b>Import PO from PDF</b> again.</p>').setWidth(420).setHeight(160);
    SpreadsheetApp.getUi().showModalDialog(prompt, 'Permission needed');
    return;
  }
  var html = HtmlService.createHtmlOutputFromFile('PoImportDialog')
    .setWidth(760)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Purchase Order');
}

/**
 * Called from the dialog with a base64 PDF. Returns parsed lines plus a verification
 * verdict; writes nothing. The caller decides whether to import.
 *
 * The parsed lines are also kept server-side under a one-time token. Import takes only that
 * token and the line numbers to keep, so what gets written is exactly what was read from the
 * PDF — the page is open to anyone with the link, and must not be able to post its own lines.
 */
function poImportExtract(base64Pdf, filename) {
  var tempIds = [];
  // Google's own errors ("Invalid argument") do not say which call failed, so every step
  // records itself and the message names it.
  var at = { step: 'decoding the upload' };
  try {
    var bytes = Utilities.base64Decode(base64Pdf);
    var blob = Utilities.newBlob(bytes, 'application/pdf', filename || 'po.pdf');

    at.step = 'counting pages';
    var pages = poPageCount_(blob);
    var text = poOcrPdfViaThumbnail_(blob, tempIds, at);
    at.step = 'reading the line items';
    var parsed = poParseText_(text);
    if (!parsed.ok) {
      return { ok: false, error: parsed.error, rawText: text.slice(0, 2000) };
    }

    at.step = 'checking the figures';
    parsed.pages = pages;
    parsed.check = poVerify_(parsed);
    parsed.masters = poCheckMasters_(parsed.lines);
    parsed.duplicate = poIsAlreadyImported_(parsed.poNumber);
    parsed.token = poRememberExtract_(parsed);
    return parsed;
  } catch (err) {
    var msg = String(err && err.message ? err.message : err);
    // Drive's OCR quota shows up two ways: a rate-limit error, or the image silently left
    // unconverted (which then reads as "Invalid argument"). Both mean the same thing.
    if (/rate limit exceeded for OCR|instead of a text document/i.test(msg)) {
      return { ok: false, error: 'Google\'s limit on reading PDFs has been reached for now. ' +
        'Nothing was imported. Try again in a few hours. (' + msg + ')' };
    }
    return { ok: false, error: 'Failed while ' + at.step + ': ' + msg };
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
function poOcrPdfViaThumbnail_(pdfBlob, tempIds, at) {
  at = at || {};
  // 1. Store the PDF unconverted so Drive generates a thumbnail for it.
  at.step = 'saving the PDF to Drive';
  var stored = DriveApp.createFile(pdfBlob.setName(PO_IMPORT_TEMP_PREFIX_ + 'src.pdf'));
  tempIds.push(stored.getId());

  // 2. Drive builds thumbnails asynchronously; poll rather than assume it is ready.
  at.step = 'waiting for the Drive preview image';
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
  at.step = 'downloading the preview image';
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
  at.step = 'running OCR on the preview image';
  var doc = Drive.Files.insert(
    { title: PO_IMPORT_TEMP_PREFIX_ + 'ocr' },
    res.getBlob(),
    { ocr: true, ocrLanguage: 'en', convert: true }
  );
  tempIds.push(doc.id);

  at.step = 'reading the OCR text';
  return poOcrText_(doc);
}

// DocumentApp.openById fails with a bare "Invalid argument" on some OCR results (seen from the
// web page, 2026-10-15). Drive's own plain-text export reads the same document without the
// Docs service, so it is the fallback; if Drive did not convert at all, say what it returned.
function poOcrText_(doc) {
  try {
    return DocumentApp.openById(doc.id).getBody().getText();
  } catch (docErr) {
    var meta = Drive.Files.get(doc.id, { fields: 'mimeType,exportLinks' });
    var link = meta.exportLinks && meta.exportLinks['text/plain'];
    if (!link) {
      throw new Error('Google returned the page as ' + meta.mimeType + ' instead of a text document, ' +
        'so OCR did not run (' + docErr.message + ').');
    }
    var res = UrlFetchApp.fetch(link, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      throw new Error('Could not read the OCR text (HTTP ' + res.getResponseCode() + ').');
    }
    return res.getContentText();
  }
}

// Only page 1 is rasterised, so a longer PO would silently lose its later lines. Count the
// page objects so that case is reported. 0 means the count is hidden inside compressed object
// streams and is unknown; the printed-totals check still catches missing lines then.
function poPageCount_(pdfBlob) {
  var raw = pdfBlob.getDataAsString('ISO-8859-1');
  return (raw.match(/\/Type\s*\/Page(?![a-zA-Z])/g) || []).length;
}

var PO_IMPORT_CACHE_PREFIX_ = 'poimp_';
var PO_IMPORT_CACHE_SECONDS_ = 1800;

function poRememberExtract_(parsed) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put(PO_IMPORT_CACHE_PREFIX_ + token, JSON.stringify({
    poNumber: parsed.poNumber, poDate: parsed.poDate, lines: parsed.lines
  }), PO_IMPORT_CACHE_SECONDS_);
  return token;
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

  // Quantity, and where present the rate, sit inline right after each code:
  //     PMCON0155 970 NOS 10.69   PMCON0105 31,248 NOS 13.94   PMCON0117 36,772 NOS   ...
  // Reading them per code — rather than harvesting each column separately — keeps every
  // value tied to the product it belongs to, which is what a column-wise read cannot
  // guarantee once OCR drops a repeat.
  //
  // Read at EACH occurrence of the code, not at the first one in the text: a PO that lists the
  // same code twice used to give both lines the first line's quantity. Only when nothing
  // follows this occurrence does it fall back to the first place the code has a quantity.
  var qtys = [];
  var inlineRates = [];
  var lineAt = /^PM[A-Z]{3}\d{4}\s+([\d,]+)\s*(?:NOS)?\s*([\d,]+\.\d{1,2})?/i;
  var codeRe = /\bPM[A-Z]{3}\d{4}\b/g;
  var hit;
  while ((hit = codeRe.exec(text)) !== null) {
    var m = text.slice(hit.index).match(lineAt) ||
      text.match(new RegExp(hit[0] + '\\s+([\\d,]+)\\s*(?:NOS)?\\s*([\\d,]+\\.\\d{1,2})?', 'i'));
    qtys.push(m ? poToNumber_(m[1]) : null);
    inlineRates.push(m && m[2] ? poToNumber_(m[2]) : null);
  }

  // THE RATE COLUMN CANNOT BE READ POSITIONALLY.
  //
  // Measured over four runs of this PO: codes and quantities come back perfectly every
  // time, but the rate list is consistently SHORT. This PO charges 13.94 on two different
  // lines and OCR collapses the repeat, returning [10.69, 13.94, 18.40] for four products.
  // It is not a misread that better matching would fix — the duplicate is simply absent, so
  // zipping by position silently assigns line 4's rate to line 3.
  //
  // Amount / qty recovers it exactly (512602 / 36772 = 13.94). So the amount column is read
  // and the rate COMPUTED from it, which is the reverse of the obvious arrangement but the
  // only one that survives duplicate rates. The OCR'd rates are still collected — not to be
  // trusted, but to be checked against, so a wrong amount cannot pass unnoticed.
  var rateMatches = text.match(/\bNOS\s+[\d,]+\.\d{1,2}\b/gi) || [];
  var ocrRates = rateMatches.map(function (s) {
    return poToNumber_(s.replace(/\bNOS\s+/i, ''));
  });

  var amounts = poCollectAmounts_(text, codes.length);

  var lines = codes.map(function (code, i) {
    var qty = qtys[i] === undefined ? null : qtys[i];
    var amount = amounts[i] === undefined ? null : amounts[i];
    var computed = (qty && amount) ? Math.round((amount / qty) * 100) / 100 : null;
    var inline = inlineRates[i];

    // Prefer the rate printed beside the code; fall back to amount / qty where OCR dropped
    // it (which happens whenever the same rate appears twice on one PO).
    var rate = inline !== null && inline !== undefined ? inline : computed;

    // Corroboration means two independent readings agree. Where both exist they must match;
    // where only the computed one exists, it has to appear somewhere in the rate column.
    var corroborated;
    if (inline !== null && inline !== undefined && computed !== null) {
      corroborated = Math.abs(inline - computed) < 0.015;
    } else if (computed !== null) {
      corroborated = ocrRates.some(function (r) { return Math.abs(r - computed) < 0.005; });
    } else {
      corroborated = false;
    }

    return {
      sr: i + 1,
      code: code,
      qty: qty,
      rate: rate,
      amount: amount,
      rateCorroborated: corroborated
    };
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
 * Amounts are the run of numbers between the last item code and the Total label.
 *
 * Two hazards in that window. Rates OCR drops from their inline position reappear here as
 * orphans — this PO's second 13.94 lands among the amounts and shifts every one of them by
 * a row. And an amount is a whole rupee figure, whereas a stray rate always carries paise.
 * Filtering to integers removes the orphans without needing to know which rate was dropped.
 */
function poCollectAmounts_(text, expected) {
  var lastCode = 0;
  var codeRe = /\bPM[A-Z]{3}\d{4}\b[^\n]*/g;
  var m;
  while ((m = codeRe.exec(text)) !== null) lastCode = m.index + m[0].length;

  var totalIdx = text.search(/\bTotal\b/i);
  var segment = text.slice(lastCode, totalIdx > lastCode ? totalIdx : text.length);

  var nums = (segment.match(/\b\d[\d,]*(?:\.\d{1,2})?\b/g) || [])
    .filter(function (s) { return s.indexOf('.') === -1; })  // drop orphaned rates
    .map(poToNumber_)
    .filter(function (n) { return n !== null && n > 0; });

  // Take the FIRST `expected` values: the amounts column comes immediately after the line
  // rows, and anything later in the window is page furniture.
  return nums.slice(0, expected);
}

/**
 * The printed totals are the outside check on everything else, so both are worth finding.
 *
 * OCR interleaves this region unpredictably — observed as
 *     Total \n 76,994 \n Packing & Forwarding Freight Charges \n 11,05,842
 * so the grand amount can sit after intervening label text rather than directly after
 * "Packing & Forwarding". Both patterns allow words in between rather than assuming
 * adjacency.
 */
function poFindPrintedTotals_(text) {
  var qty = text.match(/\bTotal\b[\s\S]{0,40}?\b([\d,]{4,})\b/i);
  var amt = text.match(/Packing[\s\S]{0,80}?\b([\d,]{6,})\b/i);
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

  // Note what is NOT checked here: qty x rate = amount. The rate is now computed as
  // amount / qty, so that test is circular and would always pass. The genuine checks are
  // that each computed rate appears in the OCR'd rate column, and that the columns sum to
  // the totals printed on the PO.
  parsed.lines.forEach(function (l) {
    sumQty += l.qty || 0;
    sumAmt += l.amount || 0;
    if (l.qty === null) problems.push('Line ' + l.sr + ' (' + l.code + '): quantity not read');
    if (l.amount === null) problems.push('Line ' + l.sr + ' (' + l.code + '): amount not read');
    if (l.rate === null) {
      problems.push('Line ' + l.sr + ' (' + l.code + '): rate could not be worked out');
    } else if (!l.rateCorroborated) {
      problems.push('Line ' + l.sr + ' (' + l.code + '): computed rate ' + l.rate +
        ' does not appear anywhere in the rate column — the quantity or amount was misread');
    }
  });

  if (parsed.printed.qty !== null && sumQty !== parsed.printed.qty) {
    problems.push('Quantities sum to ' + sumQty + ' but the PO prints ' + parsed.printed.qty);
  }
  if (parsed.printed.amount !== null && Math.abs(sumAmt - parsed.printed.amount) > 2) {
    problems.push('Amounts sum to ' + sumAmt + ' but the PO prints ' + parsed.printed.amount);
  }
  // Without the printed totals nothing proves every line was read, so this is a problem to
  // check against the PDF, not a footnote.
  if (parsed.printed.qty === null && parsed.printed.amount === null) {
    problems.push('No printed totals were found, so there is no check that every line was read. ' +
      'Count the lines against the PDF.');
  }
  if (parsed.pages > 1) {
    problems.push('This PDF has ' + parsed.pages + ' pages. Only page 1 can be read, so any ' +
      'lines on later pages are missing from this list.');
  }
  if (!parsed.poNumber) problems.push('PO number not found');
  if (!parsed.poDate) warnings.push('PO date not found — enter it by hand after importing.');

  warnings.push('Rates are computed as amount ÷ quantity. OCR drops repeated values in the ' +
    'rate column, so reading it directly would mis-assign them; each computed rate is ' +
    'checked back against the column instead.');

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
 * Write the approved lines to Dispatch. `request` is { token, srs }: the token from
 * poImportExtract and the Sr numbers left ticked. Reuses the same field logic as the manual importer:
 * Masters supplies everything the code determines, and Balance / RM Req / MB carry live
 * formulas rather than frozen constants.
 */
function poImportCommit(request) {
  // The lines come from the server's own copy of the extract, never from the page: the
  // request carries only the token and which line numbers were left ticked.
  var cache = CacheService.getScriptCache();
  var key = PO_IMPORT_CACHE_PREFIX_ + String(request && request.token);
  var stored = cache.get(key);
  if (!stored) {
    return { ok: false, error: 'This review has expired or was already imported. Upload the PDF again.' };
  }
  var extract = JSON.parse(stored);
  var keep = {};
  ((request && request.srs) || []).forEach(function (sr) { keep[Number(sr)] = true; });
  var payload = {
    poNumber: extract.poNumber, poDate: extract.poDate,
    lines: extract.lines.filter(function (l) { return keep[l.sr]; }),
  };
  if (!payload.lines.length) return { ok: false, error: 'No lines selected. Nothing imported.' };

  // Script lock, not document lock: getDocumentLock() is null when this runs from the web
  // page rather than from inside the sheet.
  var lock = LockService.getScriptLock();
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

    // Supplied sums Opening and every invoice column on the row. Without it, invoices typed
    // against a new line never reduce its Balance. Copy the row above's formula so the range
    // matches the rest of the sheet; build it only if that row has none.
    var supplied = startRow - 1 > labelRow ? dispatch.getRange(startRow - 1, cSup).getFormulaR1C1() : '';
    if (supplied) {
      dispatch.getRange(startRow, cSup, payload.lines.length, 1).setFormulaR1C1(supplied);
    } else {
      // Open-ended to the right (=SUM(BP9:9)), so invoice columns added later still count.
      var openA1 = a1(col('Opening'), startRow).replace(/[0-9]+$/, '');
      dispatch.getRange(startRow, cSup, payload.lines.length, 1).setFormulas(payload.lines.map(function (l, i) {
        var r = startRow + i;
        return ['=SUM(' + openA1 + r + ':' + r + ')'];
      }));
    }
    cache.remove(key);

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
