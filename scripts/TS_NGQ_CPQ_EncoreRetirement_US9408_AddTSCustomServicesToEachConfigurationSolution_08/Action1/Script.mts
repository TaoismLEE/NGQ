'================================================
'Test Case: CPQ_Encore Retirement_US9408_Add TS custom services to each configuration solution_08
'
'Preconditions:
'1. An Opportunity ID is ready.
'2. TS custom services product number is ready..
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

InitializeTest "IE"

'Test Data
'Fill path and file with its extension (C:\ngq-demo-develop\data\fileName.xlsx)
'ImportTestData strTestDataFile
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9408_AddTSCustomServicesToEachConfigurationSolution_08.xlsx" 'added 20jul

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<encrypted digitalbadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
Dim strProductNumberB : strProductNumberB = DataTable.Value("ProductNumberB","Global")

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' Add third party product number
'Quote_AddProductOrOption
'Quote_SetBundleID strProductNumber
'LineItemDetails_SetProductQuantityByIndex 0, intProductQuantity
LineItemDetails_AddProductByNumber strProductNumber, intProductQuantity

' Add product from Configuration OCS
'Quote_AddConfigOCS
'Quote_SelectConfigOCS 
'Quote_ServiceAndSupportCenter 
'Quote_SaveAndConvertToQuote
build_ocs_bom_serviceSupport
'Search a product and add quantity
'Quote_SearchProduct
'Quote_SearchProductByProductNumber strProductNumberB
'Quote_SearchProductSelectResult intProductQuantity
'Quote_SearchProductAddProductsToCart
'Quote_SearchProductAddProdcutsToQuote
Quote_SearchProduct
set_product_number strProductNumberB
set_quantity
add_to_quote

' END: Core
Quote_refreshPricing

Quote_save

verify_prodTable_prodNum "H0JT1A1", 20
verify_prodTable_prodNum "H9P11A1", 24
verify_prodTable_prodNum "H0JD4A1", 23

Navbar_Logout

Close_Browser
FinalizeTest
