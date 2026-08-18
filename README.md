# PDF Invoice Parser

A folder of vendor invoices in, one row per invoice out, straight into Excel.
Every invoice layout is different, so nothing here assumes a template.

**Pure PowerShell. Nothing is downloaded and nothing is installed** — not a
module, not a DLL, not Python, and Excel does not need to be on the machine.

---

## Quick start

**1. Put the files in one folder:**

```
<working folder>\
    Get-Invoices.ps1        <- run this
    Inv*.ps1                <- the engine (all of them)
    lib\InvPdfExtract.cs    <- the PDF reader
    config\                 <- your company profile lives here
    PDFs\                   <- your invoices
```

**2. Run it:**

```powershell
powershell -ExecutionPolicy Bypass -File .\Get-Invoices.ps1
```

The first run asks who **we** are (once — see below). Then it asks for the
folder, lets you pick columns, reads every PDF, walks you through anything it
was unsure about, and asks where to put the result.

**3. What you get:**

| Output | |
|---|---|
| **A workbook** | one row per invoice, frozen and filtered header, money/dates/percentages as real numbers, amber shading on anything uncertain, `Source` hyperlinked to the PDF |
| **A Log sheet** | every file and every rule that fired on it |
| **A CSV** | written alongside, always, so a workbook problem can never cost you the run |

**Unattended:**

```powershell
.\Get-Invoices.ps1 -Path .\PDFs -Out .\Q3.xlsx -NonInteractive
.\Get-Invoices.ps1 -Path .\PDFs -Columns Vendor,InvNum,InvAmt,PONum -NoReview
.\Get-Invoices.ps1 -Path .\PDFs -Preset "use tax standard"
```

---

## Who are we? (the bit that makes it accurate)

Every invoice names **two** companies — the vendor's and ours — and nothing on
the page says which is which. Parsers that guess report the customer as the
vendor and quietly corrupt a workpaper.

So the tool is told once, in `config\MyCompany.json`, and then:

1. **Veto** — a name or address matching us can never be written as the vendor.
2. **Locate** — the block holding our name *is* the bill-to / ship-to side.
3. **Deduce** — two address blocks, one is us, so the other is the vendor.
4. **Refuse** — if the only company on the page is us, the row is left blank
   and flagged. A blank beats a wrong vendor.

List every spelling that shows up on invoices you receive — abbreviations, the
old trading name, prior addresses, warehouse names:

```jsonc
{
  "Entities": [{
    "Name": "Acme Holdings LLC",
    "Aliases": ["Acme Holdings", "ACME HOLDINGS, L.L.C.", "Acme Hldgs"],
    "Addresses": [
      { "Line1": "123 Main St", "City": "Omaha", "State": "NE", "Zip": "68102" },
      { "Line1": "800 Old Plant Rd", "City": "Lincoln", "State": "NE", "Note": "prior address" }
    ]
  }],
  "ShipToNames": ["Acme Warehouse #4"]
}
```

Start from `config\MyCompany.example.json`, or just run the tool and answer the
questions.

---

## Ship-to

Ship-to is captured **whenever the block exists**, not only when freight is
charged — for use tax the delivery address is often the address that decides the
jurisdiction, and goods ship with no shipping line all the time.

`Shipped?` reads **Y** from a ship-to block *or* a freight charge. **N** means
only that the page shows no evidence — it is not proof nothing shipped. When
freight is charged but there is no ship-to address, the row says so, because
then the taxing jurisdiction is genuinely unknown.

---

## Columns

The first three are locked. The rest follow your working-paper header, then the
extras. **33 columns**, each with its own rule — press `?` in the picker to read
any of them.

```
  COLUMNS                                          preset: default
  ---------------------------------------------------------------------------
    1 [L] VenInvoice        12 [x] Taxable Amount   23 [ ] Due Date
    2 [L] Source            13 [x] Tax Amt          24 [ ] PO#
    3 [L] Provided?         14 [x] Vendor Address   25 [ ] Discount
    4 [x] City              15 [x] Subtotal         26 [ ] Terms
    5 [x] State             16 [x] Freight          27 [ ] Account#
    6 [x] Vendor            17 [x] Ship To City     28 [ ] Currency
    7 [x] Description       18 [x] Ship To State    29 [ ] Bill To Name
    8 [x] Inv#              19 [ ] Ship To Address  30 [ ] Pages
    9 [x] Inv Date          20 [ ] Ship To Name     31 [x] Confidence
   10 [x] Tax Rate          21 [x] Shipped?         32 [x] Needs Review
   11 [x] Inv Amt           22 [ ] Zip              33 [x] Extraction Notes
  ---------------------------------------------------------------------------
  [L] always written                             22 of 33 selected
  numbers toggle (e.g. 16,19-21)   a all   n none   d defaults   ? rule
  p load preset    s save preset    ENTER run    q quit
```

---

## How a field is found

Nothing is parsed by position. Every word comes out of the PDF **with its x/y
box**, so a value is found by what it sits next to:

```
  extract   every word with its position on the page
  group     words -> lines (by vertical overlap, not by rounding Y)
  anchor    find the label: "Invoice Date" / "Inv. Dt" / "DATE OF INVOICE"
  reach     take the value to its RIGHT, or BELOW it, whichever fits
  validate  does it parse as a date / an amount / an invoice number?
  score     confidence, plus a note saying where it came from
```

Each field is attacked from several directions. **When independent strategies
agree, confidence goes up. When they disagree, the row is flagged with the
conflict rather than a coin flip.**

---

## The rules (why a number in a cell can be trusted)

**Format rules** police one field. A bogus state code, an impossible tax rate, a
date in the future, an "invoice number" that is really an amount — the field is
**blanked** and the reason written to `Extraction Notes`.

**Reconciliation rules** police the fields against each other:

```
subtotal + freight - discount + tax  =  invoice total
tax / taxable amount                 =  tax rate
```

- All present and they agree → confidence goes **up**. This is the strongest
  evidence available.
- Exactly one missing → it is **derived** from the others and marked derived.
- They disagree → **nothing is silently changed.** The row is flagged and the
  arithmetic is spelled out:

  > `the amounts do not add up: subtotal 853.00 + freight 35.00 - discount 0.00 + tax 62.16 = 950.16, but the total reads 999.99 (out by -49.83)`

---

## Verify it

```powershell
.\tests\Run-AllTests.ps1
```

| Suite | Covers |
|---|---|
| **Test-Values** | money, percentages, dates, invoice numbers — and everything that must be *refused* |
| **Test-Layout** | label keys, line grouping, word-boundary anchoring, column splitting |
| **Test-Identity** | name keys, self-detection, bill-to vs ship-to separation, address parsing |
| **Test-Rules** | every format rule and every reconciliation rule, including that a mismatch does **not** change the value |
| **Test-EndToEnd** | the four sample invoices, field by field, against hand-checked expected values |
| **Test-Xlsx** | the workbook is a valid package, well-formed, correctly typed and escaped |

180 assertions, nothing to install.

```powershell
.\Get-Invoices.ps1 -Path .\one-invoice.pdf     # or, to see what the parser sees:
. .\InvPdfText.ps1 ; Show-InvPdfText -Path .\one-invoice.pdf
```

`Show-InvPdfText` paints the page back as text. If a number is not there, no
rule can find it — that is the first thing to check when something comes out
blank.

---

## Repository layout

```
Get-Invoices.ps1     the only script you run
InvPdfText.ps1       PDF -> positioned words   (+ Show-InvPdfText)
InvLayout.ps1        words -> lines -> blocks, and the spatial searches
InvValues.ps1        money / date / percentage / invoice-number parsing
InvDicts.ps1         label synonyms, states, street types, company suffixes
InvIdentity.ps1      the "we are never the vendor" guard
InvFields.ps1        one extractor per field, with confidence
InvRules.ps1         the strict rules and the arithmetic
InvRecord.ps1        one PDF -> one finished record
InvColumns.ps1       the column catalogue and the terminal picker
InvXlsx.ps1          the .xlsx writer (and the append-to-existing-workbook path)
lib\InvPdfExtract.cs the PDF reader, compiled at run time by Add-Type
config\              MyCompany.json + saved column presets
samples\             four synthetic invoices used by the tests
tests\               the test suite
docs\ARCHITECTURE.md how the pieces fit together
```

---

## Notes and limits

- **Scanned invoices have no text.** They are detected, marked `Provided? = N`
  and noted as needing OCR. Nothing is guessed from an image.
- **Line-item tables are out of scope.** One row per invoice; `Description` is a
  summary, not a parsed table.
- **First-run accuracy on a layout never seen before** is realistically 70–85%
  per field. The review queue is not a patch over that — it is part of the
  design, and it groups by vendor so a layout is corrected once.
- Windows PowerShell 5.1 and PowerShell 7 both work. `lib\InvPdfExtract.cs` is
  written to C# 5 because that is what 5.1's built-in compiler accepts.
- Nothing existing is overwritten: appending to a workbook copies it to
  `<name>.bak.xlsx` first.
