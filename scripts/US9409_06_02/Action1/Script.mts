'================================================
'Test Case: CPQ_Encore Retirement_US9400_Include Third Party Parts in the Configuration from OCS_01
'
'Preconditions:
'1. Sales op has access to NGQ.
'2. An Opportunity ID is ready.
'3. A third party product number is ready.
'
'Recommended: Use programing descriptive not objects repository
'Author: Latha Venkataram
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9409_06_02.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strBundleProductNumber : strBundleProductNumber = DataTable.Value("BundleProductNumber","Global") 
Dim intBundleProductQuantity : intBundleProductQuantity = DataTable.Value("BundleProductQuantity","Global") 
'Dim strIsBundleProduct : strIsBundleProduct = DataTable.Value("IsBundleProduct","Global")
Dim strStandaloneProductNumber : strStandaloneProductNumber = DataTable.Value("StandaloneProductNumber","Global") 
Dim intStandaloneProductQuantity : intStandaloneProductQuantity = DataTable.Value("StandaloneProductQuantity","Global") 

Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 

Dim strReceiptDate : strReceiptDate = DataTable.Value("ReceiptDate","Global") 
Dim strDeliverySpeed : strDeliverySpeed = DataTable.Value("DeliverySpeed","Global") 
Dim strTermsOfDelivery : strTermsOfDelivery = DataTable.Value("TermsOfDelivery","Global") 

''START: Core

OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector  

Quote_ValideAddButtonOptions

LineItemDetails_AddProductOrOption strStandaloneProductNumber,"1"

Quote_CustomerDataTab

Quote_ShiptoTab
Quote_ShiptoAddress
	
Quote_ShippingDataTab	
Quote_DeliverySpeed strDeliverySpeed
Quote_TermsofDelivery strTermsOfDelivery
'	
Quote_AdditionalDataTab
AdditionalData_SetReceiptDate strReceiptDate
Quote_ClickFooter
Quote_RefreshPricng

Quote_PricingTermsTab
Quote_PricingBand

strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector  

Quote_CaptureQuoteNumber

Navbar_Logout
FinalizeTest

'LV END

