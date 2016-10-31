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

InitializeTest

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
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
Dim strUploadFileName : strUploadFileName = DataTable.Value("UploadFilename","Global")

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
Quote_EditQuoteName strQuoteName
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector


' CPQ_Encore Retirement_US9400_Upload Products_03
' Upload Product
Quote_UploadProduct
LineItemDetails_UploadProduct_BrowseClick
LineItemDetails_UploadProduct_FileName strUploadFileName



'Quote_SearchBundleByID strBundleID
'Quote_SearchBundleIncludeGlobalBundles
'Quote_SearchBundleAction
'Quote_SearchBundleValidateBundleID

' END: Core
strQuotaSelection_Selector = "Refresh Pricing"
Quote_AddBundleValidation
Quote_ValideAddButtonOptions

CloseBrowser
FinalizeTest
