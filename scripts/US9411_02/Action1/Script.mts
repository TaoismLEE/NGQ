'Project Number: 205713
'User Story: US9411_02
'Description: Add multiple items by right clicking
'Tags:

Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"

DataTable.Import "..\..\data\data_file.xlsx"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "a")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID","Global")
Dim customImagingNumber : customImagingNumber = DataTable.Value("CustomImagingNumber","Global")
Dim assetTaggingNumber : assetTaggingNumber = DataTable.Value("AssetTaggingNumber","Global")
Dim thirdPartyNumber : thirdPartyNumber = DataTable.Value("ThirdPartyNumber","Global")
Dim customPackagingNumber : customPackagingNumber = DataTable.Value("CustomPackingNumber","Global")

'NOTE: automation API calls only here. No raw UFT calls!

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Click on Add+
click_lineitem_add_ocs

'add components
add_product_option customImagingNumber

add_product_option2 assetTaggingNumber

add_product_option2 thirdPartyNumber

add_product_option2 customPackagingNumber

' Refresh Pricing
click_refresh_pricing()

validate_product_number_line_item customImagingNumber
validate_product_number_line_item assetTaggingNumber
validate_product_number_line_item thirdPartyNumber
validate_product_number_line_item customPackagingNumber


' Logout and close browser
Navbar_Logout()

FinalizeTest
