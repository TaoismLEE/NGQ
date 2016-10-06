'================================================
'Test Case: CPQ_Encore Retirement_US9400_Add Products via Product Search_05
'
'Preconditions:
'1. An Opportunity ID is ready.
'2. The product number A and B are ready.
'3. A product number does not exist in Corona is ready.
'4. Sales op has access to NGQ.
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

DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_AddProductsViaProductSearch_05.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global")
Dim strProductNumberA : strProductNumberA = DataTable.Value("ProductNumberA","Global")
Dim strProductNumberNE : strProductNumberNE = DataTable.Value("ProductNumberNE","Global")
Dim strProductNumberB : strProductNumberB = DataTable.Value("ProductNumberB","Global")
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global")


'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
'strQuotaSelection_Selector = "Save"
'QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions


' CPQ_Encore Retirement_US9400_Add Products via Product Search_05
Quote_SearchProduct
'Enter the product number A 
set_product_number strProductNumberA
set_quantity

'Enter the product number that does not exist in Corona 
set_product_number strProductNumberNE
Quote_SearchProductNoQualifiedDataValidation

'Enter the product number B 
set_product_number strProductNumberB
set_quantity

'Remove product B
Quote_SearchProductRemoveItem 2

'Add Products to Quote
add_to_quote
Quote_SearchProductAddProductsToQuoteValidation

'Validate Products Added
verify_prodTable_prodNum strProductNumberA, 2


' END: Core
Quote_refreshPricing

Quote_save

Navbar_Logout

Close_Browser
FinalizeTest
