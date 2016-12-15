'Project Number: 205713
'User Story: US9412_01
'Description: Validate obsolete prducts
'Tags:

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'DataImport
DataTable.Import "..\..\data\data_file.xlsx"

' Set opportunity id and 3rd party product number
'Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "k")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim obsoleteNumber : obsoleteNumber = DataTable.Value("ObsoleteNumber", "Global")
Dim validNumber : validNumber = DataTable.Value("ValidNumber", "Global")
Dim deliverySpeed : deliverySpeed = DataTable.Value("DeliverySpeed", "Global")

'NOTE: automation API calls only here. No raw UFT calls!

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

' Set product number
set_product_number obsoleteNumber

' Set quantity and add to cart
set_quantity

' Set 3rd party product number
set_product_number validNumber

' Set quantity and add to cart
set_quantity

' Add to quote and verify
add_to_quote

validate_obsolete_message

validate_obsolete_object obsoleteNumber, 1

validate_obsolete_object validNumber, 0

CustomerData_ShipToTab

CustomerDataShipTo_SelectSameAsSoldToAddress

' Click shipping data tab
Quote_ShippingDataTab

' Set speed
ShippingData_SetDeliverySpeed deliverySpeed

' Set Delivery terms
ShippingData_SetTermsOfDelivery DataTable.Value("DeliveryTerms", "Global")

' Set receipt date
Quote_AdditionalDataTab

AdditionalData_SetReceiptDateNow

'Refresh Price
ClickRefreshPricing()

Quote_save

validate_obsolete_color()

select_preValidate_link

PreValidateQuoteOverwrite

PreValidateQuote_success

' Lobgout and close browser
Navbar_Logout()

browser("NGQ").close()

FinalizeTest


