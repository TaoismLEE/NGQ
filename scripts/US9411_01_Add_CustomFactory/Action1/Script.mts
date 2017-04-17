'Project Number: 205713
'User Story: US9411_01_Add_CustomFactory
'Description: This testcase validates that the user is able to add custom factory  services(Custom imaging, Asset tagging, BIOS revisions, Custom packaging) via searching
'Tags: imaging, tagging, BIOS 

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

' Import data from excel sheet
DataTable.Import "..\..\data\data_file.xlsx"

' Set opportunity id and 3rd party product number
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>","a")
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID","Global")
Dim customImagingNumber : customImagingNumber = DataTable.Value("CustomImagingNumber","Global")
Dim assetTaggingNumber : assetTaggingNumber = DataTable.Value("AssetTaggingNumber","Global")
Dim thirdPartyNumber : thirdPartyNumber = DataTable.Value("ThirdPartyNumber","Global")
Dim customPackagingNumber : customPackagingNumber = DataTable.Value("CustomPackingNumber","Global")

'NOTE: automation API calls only here. No raw UFT calls!

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74232", "CPQ_Encore Retirement_US9411_add custom factory services via searching_01"

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

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
validate_products_added_to_quote

'Refresh Price
click_refresh_pricing()

'Validate numbers were added successfully 
validate_product_number_line_item customImagingNumber 
validate_product_number_line_item assetTaggingNumber
validate_product_number_line_item thirdPartyNumber
validate_product_number_line_item customPackagingNumber

' Lobgout and close browser
Navbar_Logout()
browser("NGQ").Close

FinalizeTest

