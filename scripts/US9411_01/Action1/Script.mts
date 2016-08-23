'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "GC"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strOpportunityId : strOpportunityId = "OPE-0002916168"
Dim customImagingNumber : customImagingNumber = "HA842A1"
Dim assetTaggingNumber : assetTaggingNumber = "HA841A1"
Dim thirdPartyNumber : thirdPartyNumber = "HA844A1"
Dim customPackagingNumber : customPackagingNumber = "ZU706A"

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

' Set product number for custom imaging
set_product_number customImagingNumber 

' Set quantity and add to cart
set_quantity

' Set asset tagging number
set_product_number assetTaggingNumber 

' Set quantity and add to cart
set_quantity

' Set 3rd party product number
set_product_number thirdPartyNumber

' Set quantity and add to cart
set_quantity

' Set custom packaging number
set_product_number customPackagingNumber

' Set quantity and add to cart
set_quantity

' Add to quote and verify
add_to_quote

'Refresh Price
click_refresh_pricing()

'Validate numbers were added successfully 
validate_product_number_line_item customImagingNumber 
validate_product_number_line_item assetTaggingNumber
validate_product_number_line_item thirdPartyNumber
validate_product_number_line_item customPackagingNumber

' Lobgout and close browser
Navbar_Logout()

FinalizeTest

