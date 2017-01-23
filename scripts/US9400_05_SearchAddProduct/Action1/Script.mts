'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9400_Add Products via Product Search_05
'Description:	1. Sales op is able to add products via Product Search option to WNGQ.
'				2. Sales op is able to search product by product number.
'				3. Sales op is able to delete a previously added product via clicking delete icon in Add to Quote section.
'Tags: Add, Search, Product, 
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"
'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_AddProductsViaProductSearch_05.xlsx"

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")

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

' For Jenkins Reporting
dumpJenkinsOutput "US9409_10_SearchAddProduct", "74225", "CPQ_Encore Retirement_US9400_Add Products via Product Search_05"

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
