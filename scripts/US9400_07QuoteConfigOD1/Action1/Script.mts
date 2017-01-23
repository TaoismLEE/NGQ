'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9400_Create Quote with Configuration and Suppress OD1_07
'Description:	1. Sales op is able to create a quote with a configuration and suppress OD1's.
'				2. There is a new option named "with OD1 suppress" below Other Option in Customize Output page. The checkbox of this option is checked by default.
'				3. The OD1 products are visible in the output file after checking "with OD1 suppress" option.
'Tags: Quote, Output, OD1, 
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"
'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_CreateQuoteWithConfigurationAndSuppressOD1_07.xlsx"

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 
Dim strDeliverySpeed : strDeliverySpeed = DataTable.Value("DeliverySpeed","Global") 
Dim strDeliveryTerms : strDeliveryTerms = DataTable.Value("DeliveryTerms","Global") 
Dim strFANNumber : strFANNumber = DataTable.Value("FANNumber","Global")
Dim strOverrideReason : strOverrideReason = DataTable.Value("OverrideReason","Global")
Dim strOutputTypeSelector: strOutputTypeSelector = DataTable.Value("OutputTypeSelector","Global")
Dim dirPath : dirPath = Environment.Value("TestDir") + "\..\..\data\pdfs\"

' For Jenkins Reporting
dumpJenkinsOutput "US9400_07QuoteConfigOD1", "74227", "CPQ_Encore Retirement_US9400_Create Quote with Configuration and Suppress OD1_07"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
'strQuotaSelection_Selector = "Save"
'QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' CPQ_Encore Retirement_US9400_Create Quote with Configuration and Suppress OD1_07
' Add product from Configuration OCS

build_ocs_bom_od1

'Validate Product Table has options with OD1
verify_prodTable_prodOpt "0D1", 7
verify_prodTable_prodOpt "0D1", 9
verify_prodTable_prodOpt "0D1", 13
'Add Customer Data
Quote_CustomerDataTab
CustomerData_ShipToTab
CustomerDataShipTo_SelectSameAsSoldToAddress
CustomerData_BillToTab
CustomerDataBillTo_SelectSameAsSoldToAddress
'Add Shipping Data
Quote_ShippingDataTab
ShippingData_SetDeliverySpeed strDeliverySpeed
ShippingData_SetTermsOfDelivery strDeliveryTerms
'Add Additional Data
Quote_AdditionalDataTab
AdditionalData_SetReceiptDateNow

' END: Core
Quote_refreshPricing

Quote_save

'Pre-validation
'Complete Quote
select_preValidate_link

PreValidateQuote

PreValidateQuote_success

PreValidate_CloseValidationPage

' Validate Quote Output
Quote_QuoteOutputTab
QuoteOutput_ValidateCustomizeOutputButtonUnavailable
QuoteOutput_ValidateOutputTypeButtonOptions
'Select Output Type
' MUST BE PORTRAIT FOR PDF PARSER TO WORK!!!!!11111
strOutputTypeSelector = "Dynamic Portrait template"
OutputQuote_SetOutputType strOutputTypeSelector
QuoteOutput_ValidateCustomizeOutputButtonAvailable
QuoteOutput_CustomizeOutputButton
QuoteOutput_VerifyOD1Suppressed
QuoteOutput_SaveCustomizeOutput
Quote_CaptureQuoteNumber


OutputQuote_ClickPreview

dim pdfPath : pdfpath = dirPath + DataTable.Value("QuoteNumber_Output", "Global") + ".pdf"
SavePdfAs pdfpath
Dim pdfObj : Set pdfObj = NewPdfParser(pdfPath)
pdfObj.verifyProductsTable_quantityProduct0D1
'Note: Incluide & complete this function when step 430 is implemented 
'QuoteOutput_OpenPDF
'QuoteOutput_ValidateOD1ProdcutIntoPDF

Navbar_Logout
Close_Browser
FinalizeTest
