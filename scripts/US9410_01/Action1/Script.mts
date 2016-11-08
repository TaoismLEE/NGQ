'Project Number: 205713
'User Story: US9410_01
'Description: Add a product by right clicking
'Tags:

Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "Action1"

DataTable.Import "..\..\data\data_file.xlsx"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "a")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim strProductNumber : strProductNumber = DataTable.Value("otherProductNumber", "Global")

'NOTE: automation API calls only here. No raw UFT calls!

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Click on Add+
build_ocs_bom

'add 3rd party
add_product_option strProductNumber

' Refresh Pricing
click_refresh_pricing()

validate_product_number_line_item strProductNumber

' Logout and close browser
Navbar_Logout()

browser("NGQ").close()

FinalizeTest



