'Project Number: 205713
'User Story: US9410_02
'Description: Add a product to quote
'Tags:

Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "Action1"
DataTable.Import "..\..\data\data_US9410_02.xlsx"


' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user","Global"), DataTable.Value("pass","Global"), "a")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("prodNumber","Global")
Dim thirdPartyNumber : thirdPartyNumber = DataTable.Value("thirdParty","Global")
Dim quoteName : quoteName = DataTable.Value("QuoteName","Global")

'NOTE: automation API calls only here. No raw UFT calls!

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
'NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Enter quote name and save it
Quote_EditQuoteName quoteName
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

validate_products_added_to_quote

'Refresh Price
click_refresh_pricing()

' Lobgout and close browser
Navbar_Logout()

Browser("NGQ").Close 

FinalizeTest
