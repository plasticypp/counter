'use strict';

/**
 * Turn Drive's OCR text into PO line items.
 *
 * The essential thing to understand: OCR does NOT emit the table row by row. It reads the
 * page in column-major runs, so PO 00450 comes back as
 *
 *     PMCON0155 970 PMCON0105 31,248 PMCON0117 36,772 PMCON0235 8,004
 *     NOS 10.69 NOS 13.94 NOS 13.94 NOS 18.40
 *     10,369 4,35,597 5,12,602 1,47,274
 *
 * — every code+qty pair, then every unit+rate pair, then every amount. Reading line by line
 * would interleave four different products. So each column is collected independently and
 * the columns are zipped back into rows by position.
 *
 * Positional zipping is only safe because it is checked afterwards: qty x rate must equal
 * the line amount, and the column sums must equal the printed totals. A misalignment breaks
 * that arithmetic, so it surfaces as an error rather than as plausible wrong data.
 */

/** "4,35,597" and "1,05,842" are lakh-grouped; strip separators before parsing. */
function toNumber(raw) {
  if (raw === null || raw === undefined) return null;
  const cleaned = String(raw).replace(/[,\s₹]/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  return parseFloat(cleaned);
}

function findPoNumber(text) {
  // e.g. AMB/26-27/PAC/00450
  const m = text.match(/\b([A-Z]{2,5}\/\d{2}-\d{2}\/[A-Z]{2,5}\/\d{3,6})\b/);
  return m ? m[1] : null;
}

function findPoDate(text) {
  // The label and value are often split across OCR lines, so search the whole document for
  // the first dd/mm/yyyy and prefer one adjacent to a "Date" label when present.
  const labelled = text.match(/Date\s*:?\s*(\d{2}\/\d{2}\/\d{4})/i);
  if (labelled) return labelled[1];
  const any = text.match(/\b(\d{2}\/\d{2}\/\d{4})\b/);
  return any ? any[1] : null;
}

function findIndent(text) {
  const m = text.match(/REQ-IND\/(\d+)/i) || text.match(/INDENT\s*NO\s*:?\s*(\d+)/i);
  return m ? m[1] : null;
}

/** Item codes are the anchor: they are unambiguous and appear once per line. */
function collectCodes(text) {
  return [...text.matchAll(/\b(PM[A-Z]{3}\d{4})\b/g)].map((m) => m[1]);
}

/**
 * Quantities sit immediately after each code in the OCR stream. Pull the number that
 * follows each code rather than harvesting every number on the page.
 */
function collectQuantities(text, codes) {
  const qtys = [];
  for (const code of codes) {
    const re = new RegExp(code + String.raw`\s+([\d,]+)`);
    const m = text.match(re);
    qtys.push(m ? toNumber(m[1]) : null);
  }
  return qtys;
}

/**
 * Rates follow the unit token: "NOS 10.69".
 *
 * Do NOT trust position alone. OCR output is not byte-identical between runs — one pass
 * returned four NOS tokens and the next only three, which silently shifted every rate up a
 * row and produced a plausible-looking table with the wrong prices. So the count is checked
 * against the codes, and a mismatch is reported rather than papered over.
 */
function collectRates(text, expected) {
  const rates = [...text.matchAll(/\bNOS\s+([\d,]+\.\d{1,2})\b/gi)].map((m) => toNumber(m[1]));
  return { rates, aligned: rates.length === expected };
}

/**
 * Recover rates arithmetically when OCR drops a unit token: rate = amount / qty, rounded to
 * paise. The line amount and quantity are independently read and cross-checked against the
 * printed totals, so deriving the rate from them is sounder than trusting a positional
 * guess — and it is verified again afterwards by the qty x rate check.
 */
function deriveRates(lines) {
  return lines.map((l) =>
    l.qty && l.amount ? Math.round((l.amount / l.qty) * 100) / 100 : null
  );
}

/**
 * Amounts are the run of grouped numbers after the rate block. Restrict to the segment
 * between the last rate and the "Total" label so page furniture (GST figures, phone
 * numbers, the address block) cannot be mistaken for line amounts.
 */
function collectAmounts(text, expected) {
  const rateEnd = (() => {
    const all = [...text.matchAll(/\bNOS\s+[\d,]+\.\d{1,2}\b/gi)];
    return all.length ? all[all.length - 1].index + all[all.length - 1][0].length : 0;
  })();
  const totalIdx = text.search(/\bTotal\b/i);
  const segment = text.slice(rateEnd, totalIdx > rateEnd ? totalIdx : undefined);

  const nums = [...segment.matchAll(/\b(\d[\d,]*(?:\.\d{1,2})?)\b/g)]
    .map((m) => toNumber(m[1]))
    .filter((n) => n !== null && n > 0);

  // The amounts are the last `expected` values in that window; anything earlier is noise.
  return nums.slice(-expected);
}

function findPrintedTotals(text) {
  const qty = text.match(/\b(\d{1,3}(?:,\d{2,3})+)\s*[\r\n]+\s*Packing/i);
  const amount = text.match(/Packing\s*&?\s*Forwarding\s*[\r\n]+\s*([\d,]+)/i);
  const grand = text.match(/TOTAL\s*[\r\n]+\s*([\d,]+\.?\d*)/i);
  return {
    qty: qty ? toNumber(qty[1]) : null,
    amount: amount ? toNumber(amount[1]) : null,
    grand: grand ? toNumber(grand[1]) : null,
  };
}

/**
 * Descriptions are advisory only — Masters is the source of truth for what a code means —
 * but a wrong one on the review screen is worse than none, because it invites approving a
 * line by a name that belongs to a different product.
 *
 * Two rules keep them honest. Only look between the table header and the totals block, so
 * letterhead text ("LUBRICANTS") cannot be captured. And require the count to match the
 * number of codes exactly: a partial match means the layout is not what we assume, so drop
 * the descriptions entirely rather than pair them off by position and mislabel every row.
 */
function collectDescriptions(text, count) {
  const startIdx = text.search(/Product\s+Description/i);
  const endIdx = text.search(/\bTotal\b|Amount\s+In\s+Words/i);
  if (startIdx === -1) return [];

  const segment = text.slice(startIdx, endIdx > startIdx ? endIdx : undefined);
  const found = segment
    .split(/\r?\n/)
    .map((l) => l.trim())
    // A product line names a container and carries a size; the header rows and unit tokens
    // ("NOS", "HSN Code") do not.
    .filter((l) => /\b(BOTTLE|CAP|JAR|CAN|CONTAINER|DRUM)\b/i.test(l))
    .filter((l) => !/^(Sr|HSN|Item|Pack|Qty|Unit|Rate)\b/i.test(l));

  return found.length === count ? found : [];
}

function parsePoText(text) {
  const codes = collectCodes(text);
  if (!codes.length) {
    return { ok: false, error: 'No item codes (PMxxx####) found in the OCR text.', lines: [] };
  }

  const qtys = collectQuantities(text, codes);
  const { rates, aligned } = collectRates(text, codes.length);
  const amounts = collectAmounts(text, codes.length);
  const descriptions = collectDescriptions(text, codes.length);

  const lines = codes.map((code, i) => ({
    sr: i + 1,
    code,
    description: descriptions[i] || '',
    qty: qtys[i] ?? null,
    // Only use OCR'd rates when the count matches the codes exactly; otherwise every one
    // could be off by a row.
    rate: aligned ? rates[i] ?? null : null,
    amount: amounts[i] ?? null,
  }));

  // Fill any missing rate from amount / qty. Marked so the review screen can show which
  // figures were read and which were computed.
  const derived = deriveRates(lines);
  const rateSource = lines.map((l, i) => {
    if (l.rate !== null) return 'ocr';
    if (derived[i] !== null) {
      l.rate = derived[i];
      return 'derived';
    }
    return 'missing';
  });

  return {
    ok: true,
    poNumber: findPoNumber(text),
    poDate: findPoDate(text),
    indent: findIndent(text),
    printed: findPrintedTotals(text),
    lines,
    rateSource,
    ratesAligned: aligned,
  };
}

/**
 * Verify the parse arithmetically. This is what makes positional zipping trustworthy: if
 * the columns were misaligned, qty x rate stops matching the line amount and the sums stop
 * matching the printed totals.
 */
function verify(parsed) {
  const problems = [];
  const warnings = [];

  parsed.lines.forEach((l) => {
    if (l.qty === null) problems.push(`Line ${l.sr} (${l.code}): quantity not read`);
    if (l.rate === null) problems.push(`Line ${l.sr} (${l.code}): rate not read`);
    if (l.amount === null) {
      warnings.push(`Line ${l.sr} (${l.code}): amount not read — cannot cross-check`);
      return;
    }
    if (l.qty !== null && l.rate !== null) {
      const calc = l.qty * l.rate;
      // Suppliers round line amounts to the rupee; allow a rupee of drift.
      if (Math.abs(calc - l.amount) > 1.5) {
        problems.push(
          `Line ${l.sr} (${l.code}): ${l.qty} x ${l.rate} = ${calc.toFixed(2)}, ` +
            `but amount reads ${l.amount}`
        );
      }
    }
  });

  const sumQty = parsed.lines.reduce((a, l) => a + (l.qty || 0), 0);
  const sumAmt = parsed.lines.reduce((a, l) => a + (l.amount || 0), 0);

  if (parsed.printed.qty !== null && sumQty !== parsed.printed.qty) {
    problems.push(`Quantities sum to ${sumQty}, but the PO prints ${parsed.printed.qty}`);
  }
  if (parsed.printed.amount !== null && Math.abs(sumAmt - parsed.printed.amount) > 2) {
    problems.push(`Amounts sum to ${sumAmt}, but the PO prints ${parsed.printed.amount}`);
  }
  if (parsed.printed.qty === null && parsed.printed.amount === null) {
    warnings.push('No printed totals found, so the line items could not be cross-checked.');
  }
  if (!parsed.poNumber) problems.push('PO number not found');
  if (!parsed.poDate) warnings.push('PO date not found — it will need entering by hand');

  const derivedCount = (parsed.rateSource || []).filter((s) => s === 'derived').length;
  if (derivedCount) {
    warnings.push(
      `${derivedCount} rate(s) computed as amount / qty because OCR did not read them ` +
        `cleanly. They agree with the printed amounts, but confirm them against the PDF.`
    );
  }

  return { pass: problems.length === 0, problems, warnings, sumQty, sumAmt };
}

module.exports = { parsePoText, verify, toNumber };
