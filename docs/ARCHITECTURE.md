# Architecture

How the pieces fit together and where each responsibility lives.

## Data flow

```
                    Get-Invoices.ps1   (the only script you run)
                              |
   +--------------+-----------+-----------+----------------+
   | identity     | extract               | rules          | write
   v              v                       v                v
InvIdentity   InvPdfText -> InvLayout   InvRules        InvXlsx
config\       lib\InvPdfExtract.cs      arithmetic      .xlsx + .csv
MyCompany     InvFields + InvDicts      + format        + Log sheet
  .json       InvValues                   rules
                    |                       |
                    +------ InvRecord ------+
                       one PDF, one record
```

## Files

| File | Responsibility | Tested by |
|---|---|---|
| **Get-Invoices.ps1** | Orchestration: identity wizard, prompts, the column picker, the run loop and progress, the review queue, the output prompt. No parsing logic. | (run on real folders) |
| **lib\InvPdfExtract.cs** | The PDF format: object scan, filters, page tree, fonts, content-stream interpretation. Emits words with x/y/w/h. C# 5, compiled at run time. | `Test-Layout` (through the wrapper) |
| **InvPdfText.ps1** | Wraps the engine. `Get-InvPdfWords` and `Show-InvPdfText`. | `Test-Layout` |
| **InvLayout.ps1** | Words → lines → blocks. `Find-InvLabel`, `Get-InvRightOf`, `Get-InvBelow`, `Get-InvBlock`, `Split-InvLineGroups`. | `Test-Layout` |
| **InvValues.ps1** | Text → typed value, or `$null`. Money, percentages, dates, invoice numbers. Pure functions. | `Test-Values` |
| **InvDicts.ps1** | Label synonyms with specificity scores, states, street types, company suffixes, non-name words. Pure data. | (used by all) |
| **InvIdentity.ps1** | Who we are. Name keys, fuzzy self-matching, party-block detection and classification. | `Test-Identity` |
| **InvFields.ps1** | One extractor per field, each returning findings with confidence and provenance. `Resolve-InvFinding` combines them. | `Test-Identity`, `Test-EndToEnd` |
| **InvRules.ps1** | The strict rules: format policing, reconciliation, derivation, row confidence. | `Test-Rules` |
| **InvRecord.ps1** | One PDF → one finished record. The only place that knows the whole shape of an invoice. | `Test-EndToEnd` |
| **InvColumns.ps1** | The 33-column catalogue (each with its rule), presets, the picker UI. | (picker is interactive) |
| **InvXlsx.ps1** | The OOXML writer, the style table, the append-to-existing-workbook path, the CSV. | `Test-Xlsx` |

## Why we read the PDF ourselves

Getting characters *with their positions* is the whole game — without
coordinates you cannot tell that `950.16` belongs to `TOTAL DUE:` rather than to
the line above. The .NET libraries that do this are either AGPL (iText) or a set
of DLLs to deploy (PdfPig needs four), and neither survives a rule of "nothing
downloaded, nothing installed".

So `lib\InvPdfExtract.cs` is C# **source**, compiled at run time by `Add-Type`
using the compiler already inside the .NET Framework — the same approach
`NeXlsx.ps1` uses in the address repo. It handles:

- **Object scanning rather than the xref table.** Real invoices come out of
  every printer driver and mail-merge under the sun, and a broken or lying xref
  is common. Scanning for `N G obj` does not care, and incremental updates
  resolve naturally because a later definition wins.
- **Filters**: Flate (with a fallback across likely stream starts), LZW, ASCII85,
  ASCIIHex, RunLength, and PNG/TIFF predictors.
- **Object streams** (PDF 1.5+ packs objects inside a compressed stream).
- **The inheriting page tree** — `Resources`, `MediaBox` and `Rotate` cascade.
- **Fonts**: `/Widths`, CID `/W` arrays, `/ToUnicode` CMaps, `/Differences`,
  WinAnsi, and built-in metrics for the base-14 fonts.
- **The full text operator set** plus **Form XObject recursion**, which is where
  a surprising amount of invoice text actually lives.

Output is one record per word in **top-left page coordinates**, because that is
how a human reads an invoice.

## Why lines are built by overlap

A label and its value are often set in different sizes — `TOTAL DUE:` in 11pt
bold, `950.16` in 10pt. Rounding Y to a grid splits them onto two lines and the
value is lost. Two words share a line when their vertical extents actually
overlap, which is what the eye does.

## Why labels are matched squashed

The same label arrives as `Invoice #:`, `Invoice#`, `INVOICE NO.`, `Invoice-No`
and `invoice no`. Normalising to letters and digits only — dropping spaces,
dots, dashes and colons — collapses all of them to `invoiceno` / `invoice#`.
Matches are still anchored to real **word boundaries**, so `net` cannot match
inside `netamount` and `tax` cannot match inside `taxable` (a different field).

Each label carries a **specificity score**: `invoice date` is 98, a bare `date`
is 45. The score becomes the starting confidence, so a value found by a vague
label arrives already marked as needing a look.

## Confidence

Every field is attacked from several directions — the label to its left, the
label above it, the column it sits in, and arithmetic. `Resolve-InvFinding`
takes the strongest, then:

- **agreement** between independent strategies → **+8**
- **disagreement** → **−12** and the other candidates are named in the note
- **confirmed arithmetic** → **+10** on subtotal, tax and total

Row confidence is the **lowest** of vendor, invoice number, date, total and tax.
Below 75 the row goes to the review queue.

## Known PowerShell footguns (already guarded)

- **`@()` around a call is mandatory.** PowerShell unrolls a one-element result
  into a bare scalar, and `.Count` then disappears. Every call that can return
  one item is wrapped.
- **`@($null)` is a one-element array containing `$null`**, not an empty array.
  Functions that can find nothing return `@()` explicitly, never `$null`.
- **Variables are case-insensitive** — never pair `$R` with `$r`.
- **`$script:Name(...)` is not a function call.** It reads a variable and then
  fails on the parenthesis.

## Known Excel footguns (found by round-tripping through a second reader)

- **Custom number-format ids must be contiguous from 164.** Excel treats
  `numFmtId` as an id, but several readers index the custom list by position,
  and a gap makes them throw on open.
- **Style collections can be self-closing** — `<numFmts count="0"/>`. Code that
  only inserts before `</numFmts>` silently drops the new entries and leaves
  cells pointing at formats that do not exist.
- **Style indices are per-workbook.** Appending a sheet to somebody else's file
  means merging our fonts, fills, borders and formats onto the end of *their*
  table and remapping our ids to wherever they landed.
