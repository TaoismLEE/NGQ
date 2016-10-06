'================================================
'Test Case: CPQ_Encore Retirement_US9400_Item Actions_08
'
'Preconditions:
'1. Sales op has access to NGQ.
'2. An Opportunity ID is ready.
'
'Recommended: Use programing descriptive not objects repository
'Author: Ana Karina Orduña
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"
DataTable.Import "C:\ngq-demo-develop\data\TD_NGQ_CPQ_EncoreRetirement_US9400_ItemActions_08.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")
Dim strQuoteNumber : strQuoteNumber = "NI00156228"

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
Dim strLineItemSelector: strLineItemSelector = DataTable.Value("LineItemSelector","Global")

'START: Core
OpenNgq objUser
'Navbar_CreateNewQuote
'NewQuote_ValideEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
'OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
'Quote_EditQuoteName strQuoteName
'strQuotaSelection_Selector = "Save"
'QuoteServices_SelectOption strQuotaSelection_Selector
'Quote_ValideAddButtonOptions

' CPQ_Encore Retirement_US9400_Item Actions_08
'Validate Line Item by default and options
'Quote_LineItem
'LineItemDetails_ValidateDefaultLineNumber
'LineItemDetails_ValidateOptionsLineNumber
'LineItemRemove 1

' Add product from Configuration OCS
'Quote_AddConfigOCS
'Quote_SelectConfigOCS
'Quote_SaveAndConvertToQuote

' Add line item: Page Break
'function needs: prodcut row, line item row, option
'strLineItemSelector = "Page Break"
'LineItemDetails_SelectItemNumber 2
'LineItemDetails_SelectItemOption 8, strLineItemSelector
'strQuotaSelection_Selector = "Save"
'QuoteServices_SelectOption strQuotaSelection_Selector

' Add line item: Comment
'function needs: prodcut row, line item row, option
'strLineItemSelector = "Comment"
'LineItemDetails_SelectItemNumber 2
'LineItemDetails_SelectItemOption 9, strLineItemSelector
'strQuotaSelection_Selector = "Save"
'QuoteServices_SelectOption strQuotaSelection_Selector

'****************
'Alternative path
'****************
'QuickSearch strQuoteNumber
'QuickSearch_Search
'SelectResult_Search strQuoteNumber
'**************
'Add subtotal 
LineItemsSelectMultiple 6,8
Quote_AddSubtotal
'Try to remove Item into Subtotal
LineItem_RightClick_RemoveItem 6


'Replace Item

'Promote Item

'Demote Item

' END: Core
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector

Navbar_Logout
CloseBrowser
FinalizeTest @@ hightlight id_;_Browser("Home").Page("Home 2").Link("NI00155377")_;_script infofile_;_ZIP::ssf2.xml_;_
