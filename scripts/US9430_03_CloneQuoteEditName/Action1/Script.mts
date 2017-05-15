'================================================
'Product Number:205713
'User Story: CPQ_Encore Retirement_US9430_03: Without being part of the sales team, NGQ user access and edit others quote after clone the quote
'Description: This case is to validate:
'			1. After cloning a quote, Sales Op is able to access and edit others' quote without being part of the sales team.
'Tags: Quote, Company, Name, Clone
'Last Modified: 5/15/2017 by yu.li9@hpe.com
'================================================

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\US9430_03.xlsx"
Dim strCompanyName : strCompanyName = DataTable("strCompanyName",1)

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74248", "CPQ_Encore Retirement_US9430_Without being part of the sales team NGQ user access and edit others quote after clone the quote_03"

'Open brower and go to My Dashboard
OpenNgq(objUser)
ClickMyDashboard

'Go to Group Quote Tab
ClickMyGroupQuoteTab

'Click in the first row number
ClickMyGroupStatusCount

'To get the first quote
Dim strQuote : strQuote = GetFirstQuoteNumberofMyGroupQuote(2)

'Click the Auto filter Btn and enter the value
ClickAutoFilter
FillFilterQuoteNumber(strQuote)

'click the quoete number value
ClickQuoteNumber(2)

'Clone the Quote and save it
Click_Clone
Quote_save

'Go to Customer Data Tab and change the Company Name and save it
Quote_CustomerDataTab
ClickCompanyPencilBtn
EditCompanyName(strCompanyName)
Quote_save

'logout and close the browser
Navbar_Logout()
Close_Browser
FinalizeTest

