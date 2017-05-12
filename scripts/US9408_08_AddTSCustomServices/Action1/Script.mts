'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_08: Add TS custom services to each configuration solution
'Description:	"The case is to validate:
'               1. Add TS custom services in BOM page
'               2. Add TS custom services within configuration
'               3. Add TS custom services as standalone
'Tags: TS custom services
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'ImportTestData strTestDataFile
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9408_AddTSCustomServicesToEachConfigurationSolution_08.xlsx" 'added 20jul

'Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
Dim strProductNumberB : strProductNumberB = DataTable.Value("ProductNumberB","Global")

'For Jenkins reporting
dumpJenkinsOutput Environment.Value("TestName"), "74257", "CPQ_Encore Retirement_US9408_Add TS custom services to each configuration solution_08"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions


LineItemDetails_AddProductByNumber strProductNumber, intProductQuantity

' Add product from Configuration OCS
build_ocs_bom

'Search a product and add quantity
Quote_SearchProduct
set_product_number strProductNumberB
set_quantity
add_to_quote

' END: Core
'Quote_refreshPricing
ClickRefreshPricing
Quote_save

verify_prodTable_prodNum "H1K92A3", 5
verify_prodTable_prodNum "HA114A1", 7
verify_prodTable_prodNum strProductNumberB, 9

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
