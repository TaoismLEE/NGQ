'================================================
'Summary: Quick Search
'Description: Check quick search works fine with all qury crierials
'Creator: yu.li9@hpe.com
'Last Modified: 5/8/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
Dim strLoginUser : strLoginUser = objUser.Username

'Fetch data
DataTable.Import "..\..\data\Func_QuickSearch.xlsx"
Dim strOppID : strOppID = DataTable.Value("OpportunityID",1)
Dim strMDCPOrgID : strMDCPOrgID = DataTable.Value("MDCPOrgID",1)
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber",1)

'Dump report for jenkins
dumpJenkinsOutput Environment.Value("TestName"), "000006", "Check quick search works fine with all qury crierials"

'Open browser
OpenNgq objUser

'Create a new quote
Navbar_CreateNewQuote
OpportunityAndQuoteInfo_SetOpportunityId strOppID
OpportunityAndQuoteInfo_Import

'Add a product
LineItemDetails_AddProductByNumber strProductNumber,1

'Fill neccessary data
PreValidate_FixDataCheckErrors

'Save quote
Quote_Save

'Store quote number
Quote_CaptureQuoteNumber
Dim strQuoteNumber : strQuoteNumber = DataTable.Value("QuoteNumber_Output",1)

'Input search criterias and search
ChangeToHomePage
InputSearchCriterias strQuoteNumber,strMDCPOrgID,strProductNumber,strLoginUser,strOppID
ClickSearch_QuickSearch

'validate the search result
verify_advSearch_quoteID strQuoteNumber,2

Navbar_Logout
Close_Browser
FinalizeTest

