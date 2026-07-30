---
target: dispatch dashboard
total_score: 20
max_score: 40
na_heuristics: 
p0_count: 2
p1_count: 2
timestamp: 2026-07-30T07-46-47Z
slug: dispatch-proxy-src-dashboard-html
---
Method: dual-agent (A: a2624e608a87d9a0f · B: ade9f2dbfb4f392f4)

Note: the product record described a dark dashboard; the artifact is light (`--bg: #eef1f6`). PRODUCT.md corrected after this run.

## Design Health Score

| # | Heuristic | Score | Key Issue |
|---|-----------|-------|-----------|
| 1 | Visibility of System Status | 1 | Fetches once at load, never again — while a pulsing dot and ticking clock imply liveness |
| 2 | Match System / Real World | 3 | Sheet vocabulary honored; "Console" and bare `D` column are not plant language |
| 3 | User Control and Freedom | 2 | No sort, no collapse, no jump-to-today, no URL state to bookmark a view |
| 4 | Consistency and Standards | 2 | Same metric labelled `Sup`/`Sup`/`Supp`/`Supplied` across four renderings |
| 5 | Error Prevention | 3 | Read-only limits damage; locale-dependent dates can disagree with the sheet |
| 6 | Recognition Rather Than Recall | 2 | Machine legend scrolls away, leaving hue as the only identity signal |
| 7 | Flexibility and Efficiency | 2 | Four views is real power; no shortcuts, no deep links, no persisted preference |
| 8 | Aesthetic and Minimalist Design | 3 | Genuinely handsome and restrained; ornament occupies the status slot |
| 9 | Error Recovery | 1 | Renders raw `TypeError` text; no retry, no fallback to last-good data |
| 10 | Help and Documentation | 1 | `D`, `Sup`, `RM`, `MB` unexplained; only `title` is on a hover-only desktop element |
| **Total** | | **20/40** | **Acceptable — bottom of the band** |

## Design Specificity Verdict

Category-default with three earned exceptions. The file's header comment names its source: "ui-ux-pro-max 'Data-Dense Dashboard' pattern." Relabel the columns and ~70% ships unchanged as an unrelated SaaS product.

Product-authored: three separate DOM trees (three encodings of a time interval chosen per viewing distance); `GRID_START` anchored to the sheet's own grid; machine palette keyed to physical shop-floor assets.

Anonymous: pulsing amber "live" dot, seconds-ticking clock on a day-granularity schedule, generic enterprise blue, pill tabs. The one ratified brand commitment — the YPP / Yash Poly Plast name — appears nowhere.

Deterministic scan: exit 2, three warnings — `overused-font` (Geist, line 9), `flat-type-hierarchy` (line 49), `dark-glow` (line 83). Two are calibration artifacts: the type rule sampled only hardcoded px literals and missed the `--fs-*` token scale, and its reported 1.9:1 already exceeds its own threshold; `dark-glow` fires on a light page. The hardcoded literals it tripped over are a real defect the rule described wrongly.

Visual overlays: none — no browser automation exposed this session. All measurements are from source. Live endpoints respond (`/` 200, 38,379 bytes; `/api/data` 200, 10,919 bytes).

## Overall Impression

A well-built thing pointed slightly away from the plant. The three-rendering architecture is the real accomplishment. But the dashboard cannot tell you how old its data is while actively signalling that it is live — on a wall-mounted tablet showing 06:00's plan at 16:00, that is the difference between a mirror and a lie. The biggest opportunity is not visual: freshness, row identity under horizontal scroll, and failure states need to be as considered as the layout already is.

## What's Working

1. Three DOM trees, not one responsive table (lines 594/620/641). Desktop discrete day cells, tablet proportional bar, mobile labelled inline bar — each chosen for viewing distance and pointer type. Discharges PRODUCT.md Principle 3 structurally.
2. `GRID_START` from the sheet (lines 443–464) with timezone reasoning documented inline and an IST fallback. Cross-surface reconciliation by refusing to compute the timeline client-side.
3. Self-documenting constraints — the sticky-header comment (139–144) and anti-regression warning (335–337) exist because the fix would otherwise be silently reverted by the next editor.

## Priority Issues

**[P0] The dashboard can't tell you how old it is, while signalling that it's live.** `fetch` runs once (line 718); nothing refetches. `.brand .dot` pulses on a 2.4s loop and `#clock` ticks seconds. The shared tablet is precisely this scenario — open all shift, pulsing amber over morning's plan, on an unauthenticated surface with nobody accountable for reloading. Fix: `setInterval(load, 60000)` plus refetch on `visibilitychange`; replace the wall clock with `Updated HH:MM` from last successful fetch; tie the dot to real state with a text label; delete the pulse. → `/impeccable harden`

**[P0] Horizontal scroll detaches rows from their identity.** `table` is `width: max-content`; `th`/`td` sticky on `top` only. Scrolling right pushes Code and Item off the left edge. "Scroll right to see next week" is the planner's core motion, and at that position they face coloured rectangles with no row labels. Fix: `position: sticky; left: 0` on the first two columns with opaque background and seam shadow; add a "Today" button setting `scrollLeft`. → `/impeccable adapt`

**[P1] The smallest text and the today marker fail the plant-conditions principle.** `.date-th` is 10px rotated vertical in `--dim`; `.mg-h .mg-m` is 8px in `--faint` (2.63:1, fails AA). `--faint #9298a6` fails AA on all three grounds across 9+ usages. Today is a 2px red inset border plus 12% tint. A shared tablet is read at 600–900mm off-axis. Fix: retire `--faint` for text; `.date-th` to 11px minimum; drop `writing-mode`, show day number with month once per change (port the mobile grid's `newMonth` logic up); today as full-column wash plus persistent header cap. → `/impeccable typeset`

**[P1] Machine identity rides on hue alone, and the legend scrolls away.** Bars carry no text. `--m3` (2.97:1) and `--m7` (2.74:1) fail the 3:1 non-text threshold; the palette's claimed "matched lightness" spans 1.63×. ~8% of the male shop-floor population is red-green colour deficient, with no accommodation path and no auth to attach a preference to. Fix: print the machine number inside the first day-cell of each bar; add an `M3` chip to mobile cards; make the legend sticky. → `/impeccable colorize`

**[P2] Error and empty states are developer artifacts.** Line 727 renders `'failed to load: ' + err` — raw error text including worker strings like "GAS_URL secret not configured" — on a public URL. No `r.ok` check, so a valid-JSON 500 falls through to `ROWS = []` and renders "no rows match", telling a manager the filter is wrong when the fetch failed. Fix: skeleton loading; plain-sentence failure with Retry keeping last-good data dimmed; branch `ROWS.length === 0` before `order.length`. → `/impeccable harden`

## Persona Red Flags

**Sam (accessibility-dependent)** — zero ARIA attributes, labels, or landmarks in 735 lines. Tabs are plain `<button>`s with no `role="tab"`/`aria-selected`; active state is background colour alone. `#main` replaced via `innerHTML` with no `aria-live` — filtering silently swaps the document. `.clear-x` is a `<span>` with a click handler: ~14×14px, not focusable, not keyboard-operable. No `lang`, no `th scope`, no `:focus-visible` on tabs, no `prefers-reduced-motion` on an infinite animation. At ≤620px the sole `<h1>` is hidden, leaving zero headings.

**Riley (stress tester)** — `computeDateCols` sets `maxD` from data with no ceiling: one PO mistyped as 2027 produces ~500 day columns × every row. Filter undebounced, rebuilding all three renderings per keystroke. Every render path interpolates `r.code`, `r.name`, `r.po` unescaped into `innerHTML` — an apostrophe corrupts the render, and on an unauthenticated URL anyone who can edit the sheet can execute script in the owner's phone. Lexicographic sort puts `PO-10` before `PO-9`.

**Casey (distracted mobile — the manager)** — render-blocking Google Fonts (nine weights) before any text paints. Every mobile load builds the desktop table and tablet cards too, hidden by CSS — ~3× DOM cost for views never shown. Tabs 39px tall at the top in a scroll strip with hidden scrollbars, so no hint `By Color` exists. Mobile is the only rendering with no subtotals — `sub` computed at line 644 and never used, so the persona least able to do mental arithmetic is the only one asked to.

## Minor Observations

- Four abandoned designs ship as dead weight: `mobileGridHtml` (~65 lines, never called — `compactBar` used instead), `.card-grid`, `.card .dates`, unused mobile `sub`. The scroll-restore loop at 686 iterates an empty NodeList every render.
- `--fs-md` is declared for "the numbers people scan all day" and never applied to a data cell; `th, td` uses `--fs-sm`.
- Six hardcoded font-sizes bypass a scale whose comment claims every size references it. `--fs-xl: 20px` declared, never used.
- Two different today-tints: line 187 `rgba(194,63,31,.12)` (hardcoded, matches no token) vs line 331's correct `--today`-derived value.
- `max-height: calc(100vh - 190px)` is tied to the header's rendered height; the header `flex-wrap`s at ~1050px, growing ~40px and clipping the table below the fold.
- `overscroll-behavior: contain` is a good tablet touch — prevents pull-to-refresh nuking a mid-scroll position.

## Questions to Consider

1. If the supervisor's only question is "what runs next on this machine," why is it three interactions deep? The scope bar forbids a directive — not ordering the composition around the question the plant actually asks.
2. What is the amber dot for? If it's "makes the page feel like software," the plant-conditions principle says spend those pixels on a timestamp.
3. Is the grand-total strip anyone's number? Largest type in the file, last thing on the page, and no persona asked for it.
4. Four abandoned designs in one file — the mobile day-grid is the only mobile rendering that answers "which calendar days" precisely, and it's the one that got cut.
5. The one ratified brand commitment (the YPP name) is absent while unratified chrome churns.
