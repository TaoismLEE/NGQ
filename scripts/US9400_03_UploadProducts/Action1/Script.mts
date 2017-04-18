﻿'================================================
'Project Number:205713 
'User Story: US9400_03_Upload Products
'Description: 
	'The case is to validate:
	'1. Sales op is able to upload a product file to WNGQ.
	'2. There are small icons to identify whether a product is valid in Add to Quote section:
		'1) The valid products will be displayed with a green tick icon.
		'2) The Invalid product will be displayed with a red cross icon.
		'3. Both valid and invalid products are added to the quote.
		'4. The supported file type is .xls or xlsx.
'Tags:  Upload file, Add Valid and Invalid Products to quote

'Preconditions:
' 2. No customer information is entered.
'
'Author: Guillermo Soria
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load test data
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_UploadProducts_03.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
Dim strUploadFileName : strUploadFileName = Environment.Value("TestDir") & "\" & DataTable.Value("UploadFilename","Global")

'Jenkins plugin
dumpJenkinsOutput Environment.Value("TestName"), "74220", "CPQ_Encore Retirement_US9400_Upload Products_03"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
Quote_EditQuoteName strQuoteName
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector


'Upload product
uploadProduct
setUploadProductPath strUploadFileName
UploadProducts_changeDataColumns "B", "C", "D", "1"
UploadProducts_ProceedWithImport
UploadProducts_VerifyAddToQuoteTabDisplayed
UploadProducts_AddValidProducts
UploadProducts_ProductsAddedMsg 

'Refreshing price
strQuotaSelection_Selector = "Refresh Pricing"
QuoteServices_SelectOption strQuotaSelection_Selector

'Fill neccessary data
Quote_CustomerDatatab
CustomerData_ShipToTab
CustomerDataShipTo_SelectSameAsSoldToAddress
Quote_ShippingDataTab
ShippingData_SetDeliverySpeed DataTable.Value("DeliverySpeed", "Global")
ShippingData_SetTermsOfDelivery DataTable.Value("DeliveryTerms", "Global")
Quote_AdditionalDataTab
AdditionalData_SetReceiptDateNow
SetValueForPaymentTerm

'Refreshing price and save quote
Quote_refreshPricing
Quote_save

'PreValidate Quote
SelectPreValidate
PreValidate_DataCheckNoErrors
'PreValidate_ProductCheckNoErrors
PreValidate_ClicNoErrors
PreValidate_PriceNoErrors
PreValidate_BundleNoErrors
PreValidate_ClickCompleteQuote

Navbar_Logout
Close_Browser
FinalizeTest
