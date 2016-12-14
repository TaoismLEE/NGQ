'Project Number: 205713
'User Story: US9411_02
'Description: Add multiple items by right clicking
'Tags:

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
dumpJenkinsOutput "US9411_02", "74234", "CPQ_Encore Retirement_US9411_add custom factory services within a configuration _02"
' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Click on Add+
build_ocs_bom
scrollPageDown

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
browser("NGQ").Close

FinalizeTest
