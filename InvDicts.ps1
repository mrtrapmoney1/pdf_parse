<#
================================================================================
 InvDicts.ps1
 The lookup lists. Pure data - no logic, no I/O.

 The label lists are the heart of "every invoice is different". Each entry has
 a SPECIFICITY score: "invoice date" (95) is a far safer anchor than a bare
 "date" (40), so when a page offers both, the specific one wins. Scores feed
 straight into field confidence, so a value found by a vague label arrives
 already marked as needing a look.

 Add a phrasing you see in the wild and every later invoice benefits.
================================================================================
#>

Set-StrictMode -Version 2.0

function New-InvLabel {
    param([string]$Phrase, [int]$Score)
    [pscustomobject]@{ Phrase = $Phrase; Score = $Score }
}

# ---------------------------------------------------------------- field labels

$InvLabels = @{

    InvNum = @(
        (New-InvLabel 'invoice number'      98), (New-InvLabel 'invoice no'         97)
        (New-InvLabel 'invoice #'           97), (New-InvLabel 'invoice num'        95)
        (New-InvLabel 'invoice id'          92), (New-InvLabel 'inv number'         92)
        (New-InvLabel 'inv no'              90), (New-InvLabel 'inv #'              90)
        (New-InvLabel 'our invoice no'      92), (New-InvLabel 'document number'    80)
        (New-InvLabel 'document no'         78), (New-InvLabel 'bill number'        80)
        (New-InvLabel 'billing number'      78), (New-InvLabel 'statement number'   70)
        (New-InvLabel 'reference number'    62), (New-InvLabel 'reference no'       60)
        (New-InvLabel 'ref no'              55), (New-InvLabel 'number'             40)
        (New-InvLabel 'invoice'             45)
    )

    InvDate = @(
        (New-InvLabel 'invoice date'        98), (New-InvLabel 'date of invoice'    97)
        (New-InvLabel 'date invoiced'       95), (New-InvLabel 'inv date'           94)
        (New-InvLabel 'billing date'        88), (New-InvLabel 'bill date'          86)
        (New-InvLabel 'issue date'          85), (New-InvLabel 'date issued'        85)
        (New-InvLabel 'issued'              70), (New-InvLabel 'invoiced on'        85)
        (New-InvLabel 'document date'       78), (New-InvLabel 'statement date'     65)
        (New-InvLabel 'dated'               62), (New-InvLabel 'date'               45)
    )

    DueDate = @(
        (New-InvLabel 'payment due date'    98), (New-InvLabel 'due date'           96)
        (New-InvLabel 'date due'            94), (New-InvLabel 'payment due'        88)
        (New-InvLabel 'pay by'              82), (New-InvLabel 'net due date'       90)
        (New-InvLabel 'due'                 50)
    )

    PONum = @(
        (New-InvLabel 'purchase order number' 98), (New-InvLabel 'purchase order no' 97)
        (New-InvLabel 'purchase order'      95), (New-InvLabel 'customer po'        94)
        (New-InvLabel 'your po'             92), (New-InvLabel 'cust po'            90)
        (New-InvLabel 'po number'           95), (New-InvLabel 'po no'              92)
        (New-InvLabel 'po #'                92), (New-InvLabel 'order number'       75)
        (New-InvLabel 'job number'          60), (New-InvLabel 'po'                 70)
    )

    Subtotal = @(
        (New-InvLabel 'subtotal'            95), (New-InvLabel 'sub total'          94)
        (New-InvLabel 'sub-total'           94), (New-InvLabel 'total before tax'   96)
        (New-InvLabel 'amount before tax'   95), (New-InvLabel 'net amount'         88)
        (New-InvLabel 'net total'           86), (New-InvLabel 'merchandise total'  90)
        (New-InvLabel 'goods total'         88), (New-InvLabel 'product total'      86)
        (New-InvLabel 'total goods'         86), (New-InvLabel 'items total'        82)
        (New-InvLabel 'total excluding tax' 95), (New-InvLabel 'total ex tax'       92)
        (New-InvLabel 'net'                 55)
    )

    TaxableAmount = @(
        (New-InvLabel 'taxable amount'      98), (New-InvLabel 'taxable sales'      96)
        (New-InvLabel 'taxable subtotal'    95), (New-InvLabel 'taxable total'      95)
        (New-InvLabel 'total taxable'       94), (New-InvLabel 'taxable'            85)
    )

    TaxAmt = @(
        (New-InvLabel 'total sales tax'     97), (New-InvLabel 'sales tax amount'   97)
        (New-InvLabel 'sales tax'           95), (New-InvLabel 'use tax'            94)
        (New-InvLabel 'sales/use tax'       95), (New-InvLabel 'state tax'          88)
        (New-InvLabel 'local tax'           86), (New-InvLabel 'city tax'           86)
        (New-InvLabel 'county tax'          86), (New-InvLabel 'tax amount'         94)
        (New-InvLabel 'total tax'           94), (New-InvLabel 'gst'                88)
        (New-InvLabel 'hst'                 88), (New-InvLabel 'vat'                88)
        (New-InvLabel 'tax'                 70)
    )

    TaxRate = @(
        (New-InvLabel 'sales tax rate'      98), (New-InvLabel 'tax rate'           96)
        (New-InvLabel 'rate of tax'         94), (New-InvLabel 'tax %'              90)
        (New-InvLabel 'tax percent'         90), (New-InvLabel 'vat rate'           90)
        (New-InvLabel 'rate'                45)
    )

    InvAmt = @(
        (New-InvLabel 'total amount due'    99), (New-InvLabel 'please pay this amount' 99)
        (New-InvLabel 'pay this amount'     98), (New-InvLabel 'invoice total'      97)
        (New-InvLabel 'total due'           97), (New-InvLabel 'amount due'         96)
        (New-InvLabel 'balance due'         95), (New-InvLabel 'grand total'        96)
        (New-InvLabel 'total invoice amount' 97),(New-InvLabel 'invoice amount'     93)
        (New-InvLabel 'amount payable'      93), (New-InvLabel 'net payable'        90)
        (New-InvLabel 'total payable'       93), (New-InvLabel 'total including tax' 95)
        (New-InvLabel 'total incl tax'      93), (New-InvLabel 'total'              70)
        (New-InvLabel 'amount'              50)
    )

    Freight = @(
        (New-InvLabel 'shipping and handling' 96), (New-InvLabel 'shipping & handling' 96)
        (New-InvLabel 'freight charge'      95), (New-InvLabel 'freight'            93)
        (New-InvLabel 'shipping cost'       93), (New-InvLabel 'shipping'           90)
        (New-InvLabel 'delivery charge'     92), (New-InvLabel 'delivery fee'       92)
        (New-InvLabel 'handling'            80), (New-InvLabel 'postage'            85)
        (New-InvLabel 's&h'                 88), (New-InvLabel 'carriage'           85)
    )

    Discount = @(
        (New-InvLabel 'less discount'       96), (New-InvLabel 'trade discount'     95)
        (New-InvLabel 'discount'            90), (New-InvLabel 'allowance'          78)
    )

    Terms = @(
        (New-InvLabel 'payment terms'       96), (New-InvLabel 'terms of payment'   96)
        (New-InvLabel 'terms'               85)
    )

    AcctNum = @(
        (New-InvLabel 'account number'      96), (New-InvLabel 'account no'         94)
        (New-InvLabel 'account #'           94), (New-InvLabel 'customer number'    90)
        (New-InvLabel 'customer no'         88), (New-InvLabel 'customer #'         88)
        (New-InvLabel 'customer id'         86), (New-InvLabel 'acct'               80)
    )

    Currency = @(
        (New-InvLabel 'currency'            95)
    )
}

# ------------------------------------------------------- address-block headers

# Blocks that identify WHO. Used by the identity guard: the block holding our
# own name is the customer side, and the vendor is whoever is not us.
$InvBlockLabels = @{

    BillTo = @(
        (New-InvLabel 'bill to'             97), (New-InvLabel 'billed to'          96)
        (New-InvLabel 'bill-to'             96), (New-InvLabel 'sold to'            95)
        (New-InvLabel 'invoice to'          94), (New-InvLabel 'invoiced to'        94)
        (New-InvLabel 'customer'            70), (New-InvLabel 'buyer'              80)
        (New-InvLabel 'purchaser'           85), (New-InvLabel 'client'             70)
        (New-InvLabel 'account of'          75)
    )

    ShipTo = @(
        (New-InvLabel 'ship to'             97), (New-InvLabel 'shipped to'         96)
        (New-InvLabel 'ship-to'             96), (New-InvLabel 'deliver to'         95)
        (New-InvLabel 'delivered to'        95), (New-InvLabel 'delivery address'   94)
        (New-InvLabel 'shipping address'    94), (New-InvLabel 'destination'        85)
        (New-InvLabel 'job site'            88), (New-InvLabel 'jobsite'            88)
        (New-InvLabel 'job location'        86), (New-InvLabel 'service address'    88)
        (New-InvLabel 'service location'    86), (New-InvLabel 'site address'       86)
        (New-InvLabel 'installed at'        84), (New-InvLabel 'work performed at'  84)
    )

    RemitTo = @(
        (New-InvLabel 'remit payment to'    97), (New-InvLabel 'remittance address' 96)
        (New-InvLabel 'make checks payable to' 96), (New-InvLabel 'remit to'        95)
        (New-InvLabel 'mail payment to'     94), (New-InvLabel 'send payment to'    94)
        (New-InvLabel 'payable to'          88)
    )

    Vendor = @(
        (New-InvLabel 'sold by'             95), (New-InvLabel 'supplier'           90)
        (New-InvLabel 'vendor'              90), (New-InvLabel 'seller'             88)
        (New-InvLabel 'from'                65)
    )
}

# ------------------------------------------------------------------- reference

$InvStateAbbr = @{
    'alabama'='AL';'alaska'='AK';'arizona'='AZ';'arkansas'='AR';'california'='CA'
    'colorado'='CO';'connecticut'='CT';'delaware'='DE';'florida'='FL';'georgia'='GA'
    'hawaii'='HI';'idaho'='ID';'illinois'='IL';'indiana'='IN';'iowa'='IA'
    'kansas'='KS';'kentucky'='KY';'louisiana'='LA';'maine'='ME';'maryland'='MD'
    'massachusetts'='MA';'michigan'='MI';'minnesota'='MN';'mississippi'='MS'
    'missouri'='MO';'montana'='MT';'nebraska'='NE';'nevada'='NV';'new hampshire'='NH'
    'new jersey'='NJ';'new mexico'='NM';'new york'='NY';'north carolina'='NC'
    'north dakota'='ND';'ohio'='OH';'oklahoma'='OK';'oregon'='OR';'pennsylvania'='PA'
    'rhode island'='RI';'south carolina'='SC';'south dakota'='SD';'tennessee'='TN'
    'texas'='TX';'utah'='UT';'vermont'='VT';'virginia'='VA';'washington'='WA'
    'west virginia'='WV';'wisconsin'='WI';'wyoming'='WY'
    'district of columbia'='DC';'puerto rico'='PR'
}

$InvStateSet = @{}
foreach ($v in $InvStateAbbr.Values) { $InvStateSet[$v] = $true }

# Suffixes that mark a line as a company name, and that are stripped when
# building the key used to recognise the same vendor across invoices.
$InvCompanySuffix = @(
    'inc','inc.','incorporated','llc','l.l.c.','llp','l.l.p.','lp','l.p.'
    'ltd','ltd.','limited','corp','corp.','corporation','co','co.','company'
    'plc','pc','p.c.','pllc','gmbh','ag','sa','nv','bv','pty'
)   # NOTE: 'holdings' and 'group' are NOT here - they are part of the name

# Words that are never a vendor name, however big and bold they are printed.
$InvNotAName = @{}
foreach ($w in @(
    'invoice','tax invoice','statement','bill','receipt','credit','credit memo'
    'debit memo','purchase order','packing slip','remittance','quote','quotation'
    'estimate','proforma','pro forma','page','original','duplicate','copy'
    'account statement','past due','reminder','delivery note','work order'
    'service invoice','sales invoice','commercial invoice','customer copy'
)) { $InvNotAName[$w] = $true }

# Month names, for dates written out longhand.
$InvMonths = @{
    'jan'=1;'january'=1;'feb'=2;'february'=2;'mar'=3;'march'=3;'apr'=4;'april'=4
    'may'=5;'jun'=6;'june'=6;'jul'=7;'july'=7;'aug'=8;'august'=8
    'sep'=9;'sept'=9;'september'=9;'oct'=10;'october'=10;'nov'=11;'november'=11
    'dec'=12;'december'=12
}

# Street types, used to spot the street line of an address block.
$InvStreetType = @{}
foreach ($w in @(
    'st','street','ave','avenue','rd','road','dr','drive','blvd','boulevard'
    'ln','lane','ct','court','cir','circle','pl','place','pkwy','parkway'
    'hwy','highway','way','ter','terrace','trl','trail','sq','square','loop'
    'route','rt','box','pobox','suite','ste','floor','fl','unit','apt','building','bldg'
)) { $InvStreetType[$w] = $true }
