'================================================
'Summary: Simple search
'Description: Test the functionality of simple search with Quote ID, Company Name, and Opportunity ID
'Creator: yu.li9@hpe.com
'Last Modified: 4/18/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data
DataTable.Import "..\..\data\Func_SimpleSearch.xlsx"
Dim strOppID : strOppID = DataTable.Value("OpportunityID",1)
Dim strCompanyName : strCompanyName = DataTable.Value("CompanyName",1)

'Dump the jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "000005", "Test the functionality of simple search with Quote ID, Company Name, and Opportunity ID"

'Open browser.
OpenNgq objUser

'Create a new quote for search
Navbar_CreateNewQuote
OpportunityAndQuoteInfo_SetOpportunityId strOppID
OpportunityAndQuoteInfo_Import

'Fill neccessary data
PreValidate_FixDataCheckErrors

'Save quote
Quote_Save

'Store quote number into global sheet
Quote_CaptureQuoteNumber

'Check searching with Quote Number
ChangeToHomePage
Dim strQuoteNumber : strQuoteNumber = DataTable.Value("QuoteNumber_Output",1)
SearchWithQuoteNumber(strQuoteNumber)
verify_advSearch_quoteID strQuoteNumber,2

'Check searching with Company Name
ChangeToHomePage
SearchWithOppID(strOppID)
CheckOppSearchReslut(strOppID)

'Check searching with Opportunity ID
ChangeToHomePage
SearchWithCompanyName(strCompanyName)
CheckNameSearchReslut(strCompanyName)

Navbar_Logout
Close_Browser
FinalizeTest

