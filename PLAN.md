# PDF Invoice Parser — Plan

Read a folder of **wildly different vendor invoices (PDF)** and write one row per
invoice onto a tax working paper, with a **terminal column picker** so you decide
which columns come out on each run.

**Pure PowerShell.** Windows PowerShell 5.1. No Python. Same house style as the
`address` repo (`Get-CityCode.ps1`) — one script you run, prompts that guess and
let you press Enter, a real test suite.

---

## 0. Prior art — is somebody already doing this?

Searched GitHub for PowerShell PDF→Excel / invoice extraction. **Nothing to fork.**

| What I found | Verdict |
|---|---|
| `escottj/Doc2PDF`, `genericCog/Convert-Office-to-PDF`, `PrimusRising/toPDF`, `SashaBrooke/excel-p2pdf`, `AbledOcean64-shell/Convert-in-PDF` | All go the **wrong direction** — Office → PDF. Useless here. |
| `gom00n/NesahTabuParserPublic` | The only real PDF→Excel PowerShell tool (Israeli land-registry PDFs, pure PS + WinForms). **One fixed document layout.** Confirms the approach is viable; no reusable extraction logic. |
| `LuisenLou/powershell-sensor-inventory-updater` | PS reads TXT/PDF logs → updates an Excel inventory, with interactive correction. Closest *shape* to what we want (parse → review → write), tiny and domain-specific. |
| Search: "invoice parser pdf powershell" | **0 results.** |

Conclusion: there is no PowerShell invoice parser to adopt. Every mature one
(invoice2data, Docparser, Rossum, Azure Form Recognizer) is Python or a paid
cloud API. **We build it**, and we borrow only a text-extraction engine.

### The one dependency worth taking: text extraction

Getting characters *with their x/y position* out of a PDF is the hard part, and
it is not something to hand-roll well.

| Option | License | Verdict |
|---|---|---|
| **`UglyToad/PdfPig`** (2.5k★, active, port of PDFBox) | Apache-2.0 | **Recommended.** netstandard2.0 → loads in PS 5.1 with `Add-Type -Path`. Gives every **word with a bounding box**, plus page size. One ~1.5 MB DLL, vendored into `lib/`. No install, no admin. |
| `EvotecIT/PSWritePDF` | wraps iText7 (**AGPL**) | **Archived** repo, and AGPL is a licensing problem for internal firm work. No. |
| `pdftotext.exe -layout` (Poppler/Xpdf) | GPL | Good output, but an external binary per workstation. Supported as an **optional fallback**, not the default. |
| Word COM (Word 2013+ opens PDFs) | already licensed | Zero new dependencies, but it re-flows the layout and destroys column alignment. **Last-resort fallback only.** |
| Hand-written PDF parser in `Add-Type` C# | ours | Doable (xref → FlateDecode via `System.IO.Compression` → `Tj`/`TJ` operators → ToUnicode CMaps), but fonts and encodings are a swamp. **Only if a vendored DLL is not allowed** — see Open Question 1. |

Everything above the extraction layer is ours, in PowerShell.

---

## 1. What "all invoices are different" actually means

We never assume a layout. The engine works on **positioned words**, not lines of
text, and finds each field by *what is near it*:

```
Extract   →  every word on every page as { Page, Text, X, Y, W, H }
Group     →  words → lines → blocks (gap-based clustering, like the eye does)
Anchor    →  find the label ("Invoice Date", "Inv. Dt", "Date of Invoice")
Reach     →  take the value to the RIGHT of it, or BELOW it, whichever is closer
Validate  →  does it parse as a date / money / an invoice number?
Score     →  confidence 0-100 + where it came from (page, x, y)
```

Three tiers, in order — first one that answers with confidence wins:

1. **Vendor profile (memory).** Once you fix a field on an Acme Corp invoice, we
   save the anchor that worked, keyed by a vendor fingerprint. Every later Acme
   invoice is near-perfect and instant. *This is the whole payoff — invoice sets
   repeat, so the tool gets better every run.*
2. **Generic label rules.** A synonym dictionary per field (30-60 phrasings for
   "total", ~20 for "invoice number") + spatial reach.
3. **Heuristics.** Vendor = largest/boldest text block on page 1 that is not you.
   Grand total = the largest money value on the last page. Subtotal + tax = total
   is checked as **arithmetic**, and any two of the three recover the third.

Every field carries a **confidence** and a **note**. Nothing silently guesses.

---

## 2. "Never call the customer the vendor" — the identity guard

*(your follow-up requirement — and it is the single biggest accuracy win)*

Every invoice has **two** company names and **two** addresses on it: the vendor's
and yours. Naïve parsers flip them constantly.

**`config\MyCompany.json`** — created by a first-run wizard, editable after:

```jsonc
{
  "Entities": [
    { "Name": "Acme Holdings LLC",
      "Aliases": ["Acme Holdings", "Acme Hldgs", "ACME HOLDINGS, L.L.C.", "AHL"],
      "Addresses": [
        { "Line1": "123 Main St", "City": "Omaha", "State": "NE", "Zip": "68102" },
        { "Line1": "800 Old Plant Rd", "City": "Lincoln", "State": "NE", "Zip": "68508",
          "Note": "prior address, pre-2023 invoices" }
      ] }
  ],
  "ShipToNames": ["Acme Warehouse #4"],
  "AccountNumbers": ["ACCT-99812"]
}
```

Used four ways:

1. **Veto.** Any vendor / address candidate matching an entity name, alias, or
   address is struck out — it can never be reported as the vendor.
2. **Locate.** The block containing your name *is* the Bill-To / Ship-To block.
   The whole region gets blacklisted, including the address under it.
3. **Deduce.** Two address blocks on page 1, one is you → **the other is the
   vendor.** This alone fixes most vendor-address extractions.
4. **Sanity check.** If the only name we can find on the invoice is you, we do
   not guess — the row is flagged `NEEDS REVIEW: vendor not identified`.

Multiple entities, DBAs, old addresses, and misspellings-as-aliases are all
supported (fuzzy match, so `ACME HOLDINGS L.L.C.` matches `Acme Holdings LLC`).
Add an alias once in the review screen and it is remembered.

---

## 3. Columns

Your working-paper header, in order, is the **primary block** — always written,
always first:

```
VenInvoice | Source | Provided? | City | State | Vendor | Description |
Inv# | Inv Date | Tax Rate | Inv Amt | Taxable Amount | Tax Amt
```

**Defaults on** (what you said you need now): Vendor, Vendor Address, City,
State, Inv Date, Inv#, Description, Subtotal, Tax Amt, Inv Amt (grand total),
Source, Provided?.

Blank is a legitimate answer. Not every invoice has a tax rate; not every one has
a description. Empty cell + a note beats a wrong number.

### Column catalog (`InvColumns.ps1`)

Each column is a small record — header text, extractor, type, format, default
state — so **adding a column later is one entry, not a code change**:

| # | Key | Header | Type | Default | Notes |
|---|---|---|---|---|---|
| 1 | `VenInvoice` | VenInvoice | text | **locked** | key: `Vendor` + `Inv#` |
| 2 | `Source` | Source | text | **locked** | PDF file name (hyperlinked) |
| 3 | `Provided` | Provided? | Y/N | **locked** | did we get a readable PDF |
| 4 | `City` | City | text | on | vendor city → feeds `Get-CityCode.ps1` |
| 5 | `State` | State | text | on | 2-letter |
| 6 | `Vendor` | Vendor | text | on | never you (§2) |
| 7 | `Description` | Description | text | on | best line-item / memo text |
| 8 | `InvNum` | Inv# | text | on | kept as text, leading zeros safe |
| 9 | `InvDate` | Inv Date | date | on | normalised `MM/DD/YYYY` |
| 10 | `TaxRate` | Tax Rate | pct | on | stated, else Tax ÷ Taxable |
| 11 | `InvAmt` | Inv Amt | money | on | grand total |
| 12 | `Taxable` | Taxable Amount | money | on | stated, else Total − Tax |
| 13 | `TaxAmt` | Tax Amt | money | on | sales/use tax charged |
| 14 | `VendorAddress` | Vendor Address | text | on | full remit/from address |
| 15 | `Subtotal` | Subtotal | money | on | pre-tax |
| 16 | `Zip` | Zip | text | off | |
| 17 | `DueDate` | Due Date | date | off | |
| 18 | `PONum` | PO# | text | off | |
| 19 | `Freight` | Freight | money | off | shipping/handling |
| 20 | `Discount` | Discount | money | off | |
| 21 | `Currency` | Currency | text | off | USD unless stated |
| 22 | `Terms` | Terms | text | off | Net 30 etc. |
| 23 | `AcctNum` | Account# | text | off | |
| 24 | `RemitTo` | Remit To | text | off | when ≠ vendor address |
| 25 | `LineCount` | Lines | int | off | line items found |
| 26 | `Pages` | Pages | int | off | |
| 27 | `Confidence` | Confidence | int | on | lowest field score on the row |
| 28 | `NeedsReview` | Needs Review | Y/N | on | |
| 29 | `Notes` | Extraction Notes | text | on | why a cell is blank / suspect |

Columns 1-13 hold your exact header order; 14+ append to the right with a
sensible header. Order is settable in the picker.

---

## 4. The terminal experience

Console only — no WinForms, works over RDP and in the ISE-less world of PS 5.1.
Colour + box drawing in the style of the `address` repo's progress output.

**Screen 1 — Setup** (each prompt pre-filled with a guess, Enter accepts)
folder of PDFs · recurse? · output: new workbook or paste into an existing
working paper · header row · your company profile (first run → wizard).

**Screen 2 — Columns**

```
  COLUMNS                                    preset: [Use Tax - standard]
  ─────────────────────────────────────────────────────────────────────
   1 [L] VenInvoice        11 [x] Inv Amt          21 [ ] Currency
   2 [L] Source            12 [x] Taxable Amount   22 [ ] Terms
   3 [L] Provided?         13 [x] Tax Amt          23 [ ] Account#
   4 [x] City              14 [x] Vendor Address   24 [ ] Remit To
   5 [x] State             15 [x] Subtotal         25 [ ] Lines
   6 [x] Vendor            16 [ ] Zip              26 [ ] Pages
   7 [x] Description       17 [ ] Due Date         27 [x] Confidence
   8 [x] Inv#              18 [ ] PO#              28 [x] Needs Review
   9 [x] Inv Date          19 [ ] Freight          29 [x] Notes
  10 [x] Tax Rate          20 [ ] Discount
  ─────────────────────────────────────────────────────────────────────
  [L] locked (always written)          18 of 29 selected
  numbers toggle (e.g. 16,19-21)   a all   n none   d defaults
  o reorder    p load preset    s save preset    ENTER run    q quit
>
```

**Screen 3 — Run.** Live progress: `[■■■■■■□□□□] 62/104  Acme Corp  ✔ 12/12 fields`,
running counts of clean / needs-review / failed, and elapsed + ETA.

**Screen 4 — Review queue.** Only the low-confidence cells, one at a time:

```
  Acme Corp  |  invoice 4471.pdf  (page 1)          [3 of 11 to review]
  ─────────────────────────────────────────────────────────────────────
  Field: Tax Amt      guess: 84.19   confidence 41%
  Context:
      Subtotal ............ 1,203.00
      NE Sales Tax 7.0% ...    84.19        <-- taken from here
      Freight .............    35.00
  ─────────────────────────────────────────────────────────────────────
  [Enter] accept   [e] edit   [b] blank   [s] skip file   [t] teach vendor
```

`t` writes the anchor into the vendor profile → **that vendor is automatic from
now on.** Corrections also feed `MyCompany.json` when the mistake was a
you-vs-them mix-up.

**Unattended** — every prompt has a switch:

```powershell
.\Get-Invoices.ps1 -Path .\PDFs -Columns Default,VendorAddress,PONum `
    -Out .\Q3.xlsx -NonInteractive
```

---

## 5. Output

- **New workbook** (default) — one sheet `Invoices`, headers frozen, money and
  date formats applied, low-confidence cells shaded amber, `Source` hyperlinked
  to the PDF. Plus a `Log` sheet: one line per file, extractor decisions, errors.
- **Paste into an existing working paper** — same non-destructive discipline as
  `NeTaxPaste.ps1`: new columns only, never overwrite, full backup first.
- **CSV always**, next to the xlsx, so a failed Excel write never loses a run.

Writer: Excel COM if Excel is present (matches the `address` repo), else
`ImportExcel` if the module happens to be installed, else CSV only.

**Hand-off:** `City` / `State` / `Vendor Address` come out in exactly the shape
`Get-CityCode.ps1` wants — parse the invoices here, then run the city-code tool
on the same paper to fill the Nebraska tax codes.

---

## 6. Repo layout

```
Get-Invoices.ps1     <- the only script you run: prompts, loop, progress, review
InvPdfText.ps1       <- extraction engine: PdfPig via Add-Type -> positioned words
InvLayout.ps1        <- words -> lines -> blocks; "right of" / "below" reach
InvFields.ps1        <- one extractor per field + confidence scoring
InvDicts.ps1         <- label synonyms, states, currency, suffixes, stop-words
InvIdentity.ps1      <- MyCompany guard (§2) + fuzzy name/address match
InvProfiles.ps1      <- vendor fingerprint + learned anchors (the memory)
InvColumns.ps1       <- column catalog + the picker UI (§4)
InvWrite.ps1         <- xlsx / working-paper / CSV writers
lib\                 <- PdfPig.dll (vendored, Apache-2.0)
config\              <- MyCompany.json, columns.default.json
profiles\            <- learned vendor profiles (one .json per vendor)
samples\             <- redacted test invoices
tests\               <- Run-AllTests.ps1 + suites
docs\                <- ARCHITECTURE.md, RESEARCH.md
```

---

## 7. Build order

| Phase | Deliverable | Done when |
|---|---|---|
| **1** | `InvPdfText.ps1` + `InvLayout.ps1` | Any PDF → positioned words → lines/blocks; `-Dump` prints the reconstructed page so we can eyeball it. Scanned PDFs detected (0 words) → `Provided? = N`, `NEEDS OCR`. |
| **2** | `InvIdentity.ps1` + first-run wizard | You are never returned as the vendor; the two-address rule picks the right block. Tested with your name in Bill-To, Ship-To, and in a footer. |
| **3** | `InvFields.ps1` + `InvDicts.ps1` | The 5 defaults (Vendor, Inv Date, Subtotal, Tax, Total) + City/State/Vendor Address/Description, each with confidence. Arithmetic reconciliation working. |
| **4** | `InvColumns.ps1` + `Get-Invoices.ps1` | Full terminal experience end-to-end on a real folder; presets save and load. |
| **5** | `InvWrite.ps1` | xlsx + working-paper paste + CSV; formats and amber shading correct. |
| **6** | `InvProfiles.ps1` + review screen | Teach once, correct forever — second run over the same vendor needs no review. |
| **7** | `tests\` | Unit tests on parsing/normalising/identity/columns; end-to-end on `samples\` with an expected-results CSV; picker tested headless. |

Ship-ready after Phase 5; Phases 6-7 are what make it stop being a chore.

---

## 8. Honest limits

- **Scanned/photo invoices have no text.** Detected and flagged, never guessed
  at. OCR is a separate later phase (Windows 10+ `Windows.Media.Ocr` can do it
  with no install — worth a spike, not a promise).
- **Accuracy on a never-seen layout** will be roughly 70-85% per field. The
  review queue is not a fallback, it is part of the design — and the vendor
  memory is what drives the number toward 100% on your recurring vendors.
- **Line-item tables are out of scope** for v1. One row per invoice.
  `Description` is a best-effort summary line, not a parsed table.
- Anything the tool is unsure about is left **blank and flagged**, on purpose.

---

## 9. Open questions

1. **Is a vendored `PdfPig.dll` (Apache-2.0, ~1.5 MB, no install) acceptable?**
   The `address` repo's rule is "pure PowerShell, no Python" — a single managed
   DLL beside the scripts respects that, but it is your call. If not, the
   fallback is `pdftotext.exe`, and failing that a hand-written extractor in
   `Add-Type` C# (adds roughly a phase of work and will be weaker on odd fonts).
2. **Output default** — brand-new workbook per run, or always paste into the
   existing working paper?
3. **Vendor naming** — report the vendor exactly as printed, or normalise
   (strip `Inc`/`LLC`, title case) so the same vendor always keys identically
   across runs? Recommend: report as printed, key on normalised.
