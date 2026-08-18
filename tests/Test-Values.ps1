Start-InvSuite 'Values - money, dates, percentages, invoice numbers'

# --- money -------------------------------------------------------------------
Assert-InvEqual 'plain'              1234.56 (Convert-InvMoney '1,234.56')
Assert-InvEqual 'dollar sign'        1234.56 (Convert-InvMoney '$1,234.56')
Assert-InvEqual 'no separators'      950.16  (Convert-InvMoney '950.16')
Assert-InvEqual 'whole number'       500     (Convert-InvMoney '500')
Assert-InvEqual 'parentheses = credit' -85.00 (Convert-InvMoney '(85.00)')
Assert-InvEqual 'trailing CR = credit' -42.00 (Convert-InvMoney '42.00 CR')
Assert-InvEqual 'european format'    1234.56 (Convert-InvMoney '1.234,56')
Assert-InvEqual 'currency suffix'    42.00   (Convert-InvMoney '42.00 USD')

# things that must NOT be read as money
Assert-InvNull 'a percentage'     (Convert-InvMoney '7.0%')
Assert-InvNull 'a date'           (Convert-InvMoney '03/14/2024')
Assert-InvNull 'a phone number'   (Convert-InvMoney '402-555-0142')
Assert-InvNull 'a long id'        (Convert-InvMoney '1234567890123')
Assert-InvNull 'empty'            (Convert-InvMoney '')
Assert-InvNull 'letters'          (Convert-InvMoney 'Net 30')

# --- percentages -------------------------------------------------------------
Assert-InvEqual 'simple pct'   7.0  (Convert-InvPercent '7.0%')
Assert-InvEqual 'integer pct'  7.0  (Convert-InvPercent '7%')
Assert-InvEqual 'precise pct'  7.25 (Convert-InvPercent '7.25 %')
Assert-InvNull  'bare number is not a rate' (Convert-InvPercent '7')
Assert-InvNull  'absurd rate rejected'      (Convert-InvPercent '95%')

# --- dates -------------------------------------------------------------------
Assert-InvEqual 'us slash'    ([datetime]'2024-03-14') (Convert-InvDate '03/14/2024').Date
Assert-InvEqual 'iso'         ([datetime]'2024-05-21') (Convert-InvDate '2024-05-21').Date
Assert-InvEqual 'long month'  ([datetime]'2024-04-02') (Convert-InvDate 'April 2, 2024').Date
Assert-InvEqual 'short month' ([datetime]'2024-04-02') (Convert-InvDate '2-Apr-2024').Date
Assert-InvEqual 'two digit yr'([datetime]'2024-03-14') (Convert-InvDate '3/14/24').Date
Assert-InvEqual 'compact'     ([datetime]'2024-03-14') (Convert-InvDate '20240314').Date
Assert-InvNull  'nonsense'    (Convert-InvDate 'Net 30')
Assert-InvNull  'impossible'  (Convert-InvDate '13/45/2024')

# day/month ambiguity is reported, never silently resolved
Assert-InvTrue  'ambiguous flagged'     (Convert-InvDate '04/03/2024').Ambiguous
Assert-InvTrue  'unambiguous not flagged' (-not (Convert-InvDate '04/25/2024').Ambiguous)
Assert-InvEqual 'day-first when forced' ([datetime]'2024-03-04') (Convert-InvDate '04/03/2024' -DayFirst).Date

# --- invoice numbers ---------------------------------------------------------
Assert-InvTrue 'normal invoice number'  (Test-InvNumberLike 'MS-88213')
Assert-InvTrue 'alphanumeric'           (Test-InvNumberLike 'PE-2024-4471')
Assert-InvTrue 'plain digits'           (Test-InvNumberLike '884213')
Assert-InvTrue 'a year is not one'      (-not (Test-InvNumberLike '2024'))
Assert-InvTrue 'an amount is not one'   (-not (Test-InvNumberLike '42.00'))
Assert-InvTrue 'a page number is not'   (-not (Test-InvNumberLike '7'))
Assert-InvTrue 'a date is not one'      (-not (Test-InvNumberLike '03/14/2024'))
Assert-InvTrue 'text with money is not' (-not (Test-InvNumberLike 'Total 684.80'))
Assert-InvTrue 'letters only is not'    (-not (Test-InvNumberLike 'INVOICE'))

# --- money tokens on a mixed line --------------------------------------------
$t = @(Get-InvMoneyTokens 'NE Sales Tax 7.0%: 62.16')
Assert-InvEqual 'one token found on the tax line' 1 $t.Count
Assert-InvEqual 'the rate is skipped, the amount kept' 62.16 $t[0].Value

$t2 = @(Get-InvMoneyTokens 'PE-2024-4471')
Assert-InvEqual 'an invoice number yields no money' 0 $t2.Count
