'================================================
'Test Case: CPQ_Encore Retirement_US9400_Item Actions_08
'
'Preconditions:
'1. Sales op has access to NGQ.
'2. An Opportunity ID is ready.
'
'Author: Reese Childers
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_ItemActions_08.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "a")

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 
Dim strDeliverySpeed : strDeliverySpeed = DataTable.Value("DeliverySpeed","Global") 
Dim strDeliveryTerms : strDeliveryTerms = DataTable.Value("DeliveryTerms","Global") 
Dim strLineItemSelector: strLineItemSelector = DataTable.Value("LineItemSelector","Global")

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
click_save_button()

' Build ocs configuration
build_ocs_bom @@ hightlight id_;_Browser("Home").Page("Home 2").Link("NI00155377")_;_script infofile_;_ZIP::ssf2.xml_;_

CustomerData_ShipTo

CustomerDataShipTo_SelectSameAsSoldToAddress

' Click shipping data tab
Quote_ShippingDataTab

' Set speed
ShippingData_SetDeliverySpeed strDeliverySpeed

' Set Delivery terms
ShippingData_SetTermsOfDelivery strDeliveryTerms

' Set receipt date
Quote_AdditionalDataTab

AdditionalData_SetReceiptDateNow

'Refresh Pricing
click_refresh_pricing()

rightClickAddPageBreak()

rightClickAddComment()

selectMultipleLines

removeItemInSubTotal

deleteSubTotalLine

removeItem

click_refresh_pricing()

editProductNum strProductNumber

rightClickPromoteItem

rightClickDemoteItem

' Logout and close browser
Navbar_Logout()
browser("NGQ").Close

FinalizeTest
