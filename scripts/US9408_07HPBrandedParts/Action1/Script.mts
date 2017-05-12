'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_07: Add HP branded third party parts in any configuration
'Author: Joshua Hunter
'Description: This test deals with testing adding OCS item with HP Branded Parts
'Tags: Quote, OCS
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'ImportTestData strTestDataFile
DataTable.Import "..\..\data\US9408_07HPBrandedParts.xlsx"

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

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74256", "CPQ_Encore Retirement_US9408_Add HP branded third party parts in any configuration_07"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

'Add a HP branded part
LineItemDetails_AddProductByNumber strProductNumber, 1

'Add a configuration from OCS
build_ocs_bom

Quote_Refresh_Pricing
Quote_Save

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
