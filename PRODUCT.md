# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

A small back-office team sharing one tool. Members have
differing levels of familiarity with it, so states must explain themselves
without training or a manual.

The job: a supplier order report arrives as a PDF; they need its line items as a
usable spreadsheet, and they need to be confident nothing was lost in the
conversion before they rely on the numbers.

## Product Purpose

Converts supplier PDFs into structured Excel workbooks. It exists because these
reports arrive only as PDFs and were otherwise re-keyed by hand — slow, and
silently error-prone.

Success is not "a file was produced". Success is **a file the team can trust
without re-checking it against the PDF**. Extraction that quietly drops rows is
worse than no tool at all, because it looks like it worked.

## Positioning

Beyond generic PDF-to-Excel conversion, it holds a dedicated reader for the
supplier order report, and it reads that report from the PDF's
**physical column geometry** rather than by pattern-matching flattened text.
A value's column is a fact recorded in the file; its shape is only a guess.

That is what lets it handle wrapped cells, blank certificates and unexpected
product types without special cases — and it is why every row can be accounted
for rather than approximately captured.

## Operating Context

- Windows desktop. Launched by double-clicking `convert.bat`, which starts a
  local Streamlit server and opens the browser. Not deployed or public.
- Two surfaces: the **converter app** (`app.py`, the daily tool) and a
  **standalone reference page** (`index.html`, opened directly from disk).
- Input arrives as one PDF at a time, or a folder for batch runs.
- Reports state their own totals in a header line (`25 Items - $5970.06`),
  which is the team's fastest reconciliation check.
- Output workbooks contain real customer names, order numbers and amounts.

## Capabilities and Constraints

- Handles digital, scanned and mixed PDFs; page type is detected automatically.
- Three table engines tried in order; two are optional and often absent.
- OCR for scanned pages requires Tesseract, which may not be installed.
- Output: a formatted `.xlsx`, or one `.csv` per extracted dataset — including
  the order report, which the CSV path reads through the same cascade as Excel.
  Batch runs also produce a run-level summary workbook.
- **Rows failing validation are excluded and reported in four places** — console,
  log file, workbook metadata, and the batch summary. This reporting is the
  product's central promise and must stay prominent in any redesign.
- Counts are two distinct measures: `Tables Found` (datasets written) and
  `Rows Extracted` (data rows). Conflating them has already caused confusion.
- Streamlit constrains the UI: styling is applied over its own DOM, so the
  design must work with its widget set rather than assume arbitrary markup.
- Password-protected PDFs are unsupported.

## Brand Commitments

No formal brand guide exists; the user has
explicitly opened the visual direction, so the incumbent navy palette is
evidence, not a constraint.

**Standing preference (confirmed):** the interface follows the enterprise
console convention, executed at full fidelity rather than as a template. The
craft bar the user named is **Linear, Stripe Dashboard, Vercel/Geist and
Notion** — density and state discipline from Linear, financial-grade numeric
clarity from Stripe, hairline monochrome restraint from Geist, and enough
breathing room and in-place explanation from Notion to serve a mixed-skill
team. Future work matches that level of finish; it does not substitute a more
expressive direction.

## Evidence on Hand

- Verified against six live supplier invoices: **100 of 100 line items
  extracted, reconciling to the cent** ($91,948.90 across the six).
- Automated test suite: 59 tests, all passing.
- Real invoice PDFs exist locally but are **not** in the repository, and
  generated workbooks are git-ignored because they carry customer data.
- No public customers, testimonials, benchmarks or pricing exist. Do not invent
  any.

## Product Principles

1. **Trust is the deliverable.** The interface's job is to let someone believe
   the output. Anything that undermines confidence outranks anything that
   speeds up a click.
2. **Never fail silently.** An excluded row must be impossible to miss. Quiet
   partial success is the one unacceptable outcome.
3. **Explain in place.** The team shares the tool at mixed skill levels; states,
   defaults and warnings must teach without a manual.
4. **The document is the authority.** Read what the file records; do not infer
   what it probably means.
5. **Handle the whole range, degrade honestly.** Missing engines and odd
   documents are normal; say what is unavailable rather than pretending.

## Accessibility & Inclusion

No product-specific standard has been established. Treat WCAG AA contrast and
full keyboard operability as the working floor, and respect reduced-motion.
