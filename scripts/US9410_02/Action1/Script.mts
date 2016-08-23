'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strOpportunityId : strOpportunityId = "OPE-0002916168"
Dim strProductNumber : strProductNumber = "726722-B21"
Dim thirdPartyNumber : thirdPartyNumber = "G1S72A"

'NOTE: automation API calls only here. No raw UFT calls!

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Enter quote name and save it
Quote_EditQuoteName "Test Name"
click_save_button()

' Click on Add+
click_lineitem_add_product_search

' Set product number
set_product_number strProductNumber

' Set quantity and add to cart
set_quantity

' Set 3rd party product number
set_product_number thirdPartyNumber

' Set quantity and add to cart
set_quantity

' Add to quote and verify
add_to_quote

'Refresh Price
click_refresh_pricing()

' Lobgout and close browser
Navbar_Logout()

FinalizeTest
