'==================================================================================================
'Project Number: 205713
'User Story: US9412_01_Quote_Completion_Process
'Description: Validate Sales Op can complete quote with obsolete products
'Tags: obsolete, completion
'==================================================================================================

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'DataImport
DataTable.Import "..\..\data\data_file.xlsx"

Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim obsoleteNumber : obsoleteNumber = DataTable.Value("ObsoleteNumber", "Global")
Dim validNumber : validNumber = DataTable.Value("ValidNumber", "Global")
Dim deliverySpeed : deliverySpeed = DataTable.Value("DeliverySpeed", "Global")

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74259", "CPQ_Encore Retirement_US9412_Be aware of end of life products_Quote Completion Process_01"

'Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

'Enter quote name and save it
Quote_EditQuoteName "Test Name"
Quote_save

'Click on Add+
click_lineitem_add_product_search

'Set product number
set_product_number obsoleteNumber

'Set quantity and add to cart
set_quantity

'Set 3rd party product number
set_product_number validNumber

'Set quantity and add to cart
set_quantity

' Add to quote and verify
add_to_quote

validate_obsolete_message

validate_obsolete_object obsoleteNumber, 1

validate_obsolete_object validNumber, 0

'Set required data
PreValidate_FixDataCheckErrors

'Refresh Price
ClickRefreshPricing()
Quote_save

'validate_obsolete_color()
LineItemDetails_ValidateProductObsoleteFontColor 2, obsoleteNumber

'Select pre-validate from the drop down menu
SelectPreValidate

'Validate that there are no errors in Data Check, CLIC, Price, Bundle
PreValidate_DataCheckNoErrors
PreValidate_ClicNoErrors
PreValidate_PriceNoErrors
PreValidate_BundleNoErrors

'Complete the quote
PreValidate_ClickCompleteQuote
PreValidate_CloseValidationPage

'Lobgout and close browser
Navbar_Logout
Close_Browser
FinalizeTest


