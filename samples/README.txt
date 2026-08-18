These four PDFs are synthetic invoices, generated to exercise the parser:

  inv_classic.pdf       vendor top-left, bill-to and ship-to side by side,
                        right-aligned totals, freight line, Flate compressed
  inv_labels_above.pdf  labels above values, longhand date, TJ kerned text,
                        taxable amount and rate stated, no ship-to
  inv_shipto.pdf        uncompressed content stream, ship-to differs from
                        bill-to, no freight line despite goods shipping
  inv_twopage.pdf       two pages, totals only on the last one

They contain no real vendor or customer data. Test-EndToEnd.ps1 checks every
extracted field on all four against hand-checked expected values, so they are
the regression baseline: if a change breaks one of these, it broke extraction.
