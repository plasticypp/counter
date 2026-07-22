# Counter Entry — Implementation Plan

## Intent
An operator scans a QR sticker on a machine, sees what it is running, types the
counter reading plus any rejection and runner scrap, and saves. Time is captured
by the server. Production is derived, never typed.

## Storage — ONE sheet, `Production Log`, append-only

| Timestamp | Machine | Code | Counter | Rejection | Runner Kg |
|---|---|---|---|---|---|

- One row per submission. Nothing is ever overwritten or updated in place.
- `Production` is NOT stored — it is `counter - previous counter for that machine`.
  Storing a derived value invites it disagreeing with its inputs.
- Current state of a machine = its last row.
- Run totals = rows for that machine since the last row whose `Code` differs.
  "Counters reset on product change" needs no reset mechanism; it falls out of
  the query window.

Rejected earlier: a second `Machine Runs` sheet with RunIDs, status and open/close
lifecycle. It existed only to make current-state lookup fast, which for 7 machines
and a handful of readings a day is premature. Two sheets can also disagree; one
cannot.

## Endpoints (both in the existing doGet/doPost switches)

- `doGet ?action=counter&machine=N` → serves the page (HtmlService, from file
  `Counter.html`).
- `doGet ?action=counterState&machine=N` → JSON: current code, last counter, last
  timestamp, run totals, and the product list.
- `doPost {action:'logCounter', machine, code, counter, rejection, runner}` →
  validates, appends one row, returns the new state.

## Validation (server side — the client is a convenience, not a gate)
1. `machine` required and numeric.
2. `code` required.
3. `counter` required, numeric, >= 0.
4. If the machine has a previous row with the SAME code and the new counter is
   lower → reject. A backwards counter is a replaced/rolled-over counter, not
   production. Selecting a different product re-anchors and is allowed.
5. `rejection` / `runner` optional, default 0, must be >= 0.

## QR
Each machine gets a printed sticker encoding:
`https://script.google.com/.../exec?action=counter&machine=5`
The phone's native camera opens it. No scanner library, no camera permission, no
Apps Script sandbox fight.

## UI
Single page, three states (first reading / normal / error). Bilingual EN + Hindi
shown together — no toggle, no stored preference. 3 taps to submit. Product
dropdown persists between entries, so a normal entry never touches it.
Fonts: IBM Plex Sans + Plex Sans Devanagari + Plex Mono (one family across both
scripts; the recommended Fira pairing has no Devanagari).

## Deliberately NOT built
- Operator identity column — nobody asked.
- Write-back to Dispatch (Supplied/Balance) — separate concern, and it would
  touch the columns just repaired. Add once the log is proven.
- Changeover confirmation screen — changing the dropdown IS the changeover.
- Duplicate-save guard beyond the backwards-counter check.

## Open (defaulted, say the word to change)
- Product dropdown lists ALL Masters codes, machine's scheduled jobs first.
- No confirm step on product change.
