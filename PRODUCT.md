# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Three audiences on three distinct surfaces, which is why the dispatch dashboard
renders three ways rather than one shrinking table:

- **Production planner** — desktop, in the office. Works the schedule in the
  Google Sheet; uses the dashboard's full table and Gantt grid to review.
- **Owner / manager** — phone, frequently away from the plant. Needs dispatch
  status at a glance; the mobile card view is their whole experience.
- **Shop-floor supervisor** — tablet or shared screen near the machines.
  Checking what runs next under plant conditions.

Machine operators are a separate audience served by the Counter Entry app: they
scan a per-machine QR sticker, log a counter reading plus rejection and runner
weight, and pick their name. Hindi-first.

## Product Purpose

A set of tools around Pack Masters' plastic bottle production dispatch. The
dispatch dashboard tells the shop floor what to run — it is the daily operating
instruction, not a reporting artifact. The Counter Entry app captures actual
production against that plan.

Success is that the right job runs on the right machine without anyone opening
a spreadsheet on the floor.

## Positioning

The schedule already lives in a Google Sheet the planner trusts and edits daily.
These tools do not replace that sheet or ask anyone to migrate; they project it
onto the surfaces where the sheet itself is unusable — a phone in transit, a
tablet beside a machine, a QR sticker an operator scans.

## Operating Context

- The **Dispatch Google Sheet** (`1O62QOLbsyr8Z1yNNBdx0tClDxB8CbYo7K8C5UhFkTz0`,
  tab "Dispatch") is the system of record: machine, priority, PO, product, order
  quantity, supplied, balance, plus a per-day Gantt grid and invoice columns.
- An Apps Script (`clasp-jpermits`) runs an `onEdit` trigger that re-sorts by
  machine and priority, recomputes Output/Start/End/Days, and paints the Gantt.
- The **dashboard** is a Cloudflare Worker (`dispatch-proxy`) that proxies Apps
  Script data and renders it. Read-only.
- The **Counter Entry** app is served from Apps Script, framed by a GitHub Pages
  wrapper, reached by scanning a printed per-machine QR sheet.
- Purchase orders arrive as supplier PDFs and are transcribed into the sheet.
- Product facts (shape, pack size, color, tare weight, MB ratio, output per day)
  come from a `Masters` tab keyed on product code.

## Capabilities and Constraints

- **The Google Sheet stays the source of truth.** The dashboard is a read view;
  planners work in Sheets. Nothing in the dashboard writes back.
- **The dashboard stays a schedule mirror.** It shows the plan and lets humans
  interpret it. No "run this now" directive, no acknowledgement loop, no plan-vs-
  actual reconciliation — even though the Apps Script already computes a
  current-job-per-machine value and the Counter Entry log could support it.
- **Operator surfaces are Hindi-first** with an English toggle.
- **No authentication.** Every surface is reached by QR or shared URL, so
  nothing sensitive may be displayed.
- Dispatch sheet columns are reordered by hand over time; anything reading them
  must resolve by header name, never by index.
- Three renderings at 1000px and 620px breakpoints — deliberately distinct
  layouts, not one responsive table.

## Brand Commitments

The **YPP / Yash Poly Plast** name. Nothing else is fixed — no committed logo,
palette, or typeface. The current light dashboard (`--bg: #eef1f6`, titled
"Dispatch Console") and blue Counter Entry are incumbent implementation, not a
house style anyone has ratified. Note the ratified name appears nowhere in the
dashboard itself.

## Evidence on Hand

- Live production data in the Dispatch sheet: real POs, machines, quantities.
- `Masters` tab with the real product catalogue (~83 materials).
- Real customer POs (e.g. Arabian Petroleum, ARZOL bottle range).
- Printed QR sheet and operator guide (`Counter_Guide_A4.pdf`).
- No testimonials, benchmarks, pricing, or customer quotes exist. Do not invent
  them.

## Product Principles

1. **The sheet is the truth; every surface is a projection of it.** Never
   introduce a second place where dispatch state can be edited or disagree.
2. **Resolve by name, never by position.** Both the sheet's columns and its
   layout shift under you; positional assumptions fail silently and look like
   bad data rather than bugs.
3. **Each surface earns its own layout.** A planner's table, a manager's card,
   and an operator's tile are different jobs — do not compromise one to reuse
   another.
4. **Legible under plant conditions.** Glare, distance, dirty hands, a cheap
   phone. Density is not a virtue if it cannot be read at arm's length.
5. **Show the plan; don't command.** The tools inform the people who decide.

## Accessibility & Inclusion

Hindi-first operator surfaces with an English toggle. Operators are not assumed
to be comfortable readers of English or of dense numeric tables.
