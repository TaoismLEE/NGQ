'================================================
'Test Case: CPQ_Encore Retirement_US9400_Upload Products_03
'
'Preconditions:
' 2. No customer information is entered.
'
'Recommended: Use programing descriptive not objects repository
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
'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
'Test Data
'Fill path and file with its extension (C:\ngq-demo-develop\data\fileName.xlsx)
'ImportTestData strTestDataFile
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_UploadProducts_03.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""
 
'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<encrypted DigitalBadge>")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
Dim strUploadFileName : strUploadFileName = Environment.Value("TestDir") & "\" & DataTable.Value("UploadFilename","Global")

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
Quote_EditQuoteName strQuoteName
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector


' CPQ_Encore Retirement_US9400_Upload Products_03
' Upload Product
'Quote_UploadProduct
'LineItemDetails_UploadProduct_BrowseClick
'LineItemDetails_UploadProduct_FileName strUploadFileName
uploadProduct
setUploadProductPath strUploadFileName
'UploadProducts_VerifyProducts strUploadFileName
UploadProducts_changeDataColumns "B", "C", "D", "1"
UploadProducts_ProceedWithImport
UploadProducts_VerifyAddToQuoteTabDisplayed
UploadProducts_AddValidProducts
UploadProducts_ProductsAddedMsg 


'Quote_SearchBundleByID strBundleID
'Quote_SearchBundleIncludeGlobalBundles
'Quote_SearchBundleAction
'Quote_SearchBundleValidateBundleID

' END: Core
strQuotaSelection_Selector = "Refresh Pricing"
QuoteServices_SelectOption strQuotaSelection_Selector
'Quote_AddBundleValidation
'Quote_ValidateAddButtonOptions

' REQUIRED FOR PRE-VALIDATION TO PASS
Quote_CustomerDatatab

CustomerData_ShipToTab

CustomerDataShipTo_SelectSameAsSoldToAddress

Quote_ShippingDataTab

ShippingData_SetDeliverySpeed DataTable.Value("DeliverySpeed", "Global")

ShippingData_SetTermsOfDelivery DataTable.Value("DeliveryTerms", "Global")

Quote_AdditionalDataTab

AdditionalData_SetReceiptDateNow
' END REQUIREMENTS FOR PRE-VALIDATION TO PASS

Quote_refreshPricing

Quote_save

select_preValidate_link

'PreValidateQuote
PreValidateQuoteOverwrite

PreValidateQuote_success

Navbar_Logout

Close_Browser
FinalizeTest
