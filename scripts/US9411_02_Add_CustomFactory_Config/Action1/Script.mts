'Project Number: 205713
'User Story: US9411_02_Add_CustomFactory_Config
'Description: This test case validates that the user is able to add custom factory services within a configuration
'Tags: imaging, tagging, asset, BIOS, custom, packaging

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

DataTable.Import "..\..\data\data_file.xlsx"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "a")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID","Global")
Dim customImagingNumber : customImagingNumber = DataTable.Value("CustomImagingNumber","Global")
Dim assetTaggingNumber : assetTaggingNumber = DataTable.Value("AssetTaggingNumber","Global")
Dim thirdPartyNumber : thirdPartyNumber = DataTable.Value("ThirdPartyNumber","Global")
Dim customPackagingNumber : customPackagingNumber = DataTable.Value("CustomPackingNumber","Global")

'NOTE: automation API calls only here. No raw UFT calls!
'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74234", "CPQ_Encore Retirement_US9411_add custom factory services within a configuration _02"
' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

'Add a config
build_ocs_bom
scrollPageDown

'add components
click_lineitem_add_product_search
set_product_number customImagingNumber
set_quantity
set_product_number assetTaggingNumber
set_quantity
set_product_number thirdPartyNumber
set_quantity
set_product_number customPackagingNumber
set_quantity

' Add to quote and verify
add_to_quote
validate_products_added_to_quote

' Refresh Pricing
click_refresh_pricing()

validate_product_number_line_item customImagingNumber
validate_product_number_line_item assetTaggingNumber
validate_product_number_line_item thirdPartyNumber
validate_product_number_line_item customPackagingNumber


' Logout and close browser
Navbar_Logout()
browser("NGQ").Close

FinalizeTest
