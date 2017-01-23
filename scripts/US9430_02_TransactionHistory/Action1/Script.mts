'================================================
'Product Number:205713
'User Story: CPQ_Encore Retirement_US9430_NGQ capture the transaction history_02
'Description: This case is to validate:
'			1. NGQ is able to capture the email ID and timestamping of the last person having edited the quote.
'			2. NGQ is able to capture the transaction history of the quote.
'Tags: Advanced, Search, Filter, Table,
'================================================

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'DataTable.Import "..\..\data\US9430_02.xlsx"
'Dim strQuoteNumber : strQuoteNumber = DataTable("strQuoteNumber",1)

' For Jenkins Reporting
dumpJenkinsOutput "US9430_02_TransactionHistory", "74247", "CPQ_Encore Retirement_US9430_NGQ capture the transaction history_02"

InitializeTest "Action1"

'Open browser and go to My Dashboard
OpenNgq(objUser)

'Go to my Dashboard
ClickMyDashboard()

'Validate If Quote Tab is selected
ValidateQuoteTab()

'Click the autofilter button, set the Quote Number
ClickAutoFilter()

'Set Quote Number
'NI00159565
FillFilterQuoteNumber("NI00162806")
'FillFilterQuoteNumber(strQuoteNumber)

'validate if the colums 'Last Modify by' 'Las modified Ts' and 'Owner History' are active
ValidateLastModifyBy_TS_and_OwnerHistory()

'Scroll SideBar to validate fields
MoveScrollBarToRight()

'Validate if the colums 'Last Modify by' 'Las modified Ts' and 'Owner History' has value different from NULL
ValidateFieldsByTsOwner(2)

'Goto Owner History and Validate if the table has deployed
ClickOwnerHistory(2)
ValidateOwnerHistoryTable()

'go to My Dashboard
ClickMyDashboard()

ValidateQuoteTab()
ClickAutoFilter()
FillFilterQuoteNumber("NI00162806")
'FillFilterQuoteNumber(strQuoteNumber)

'Click the quote Number
ClickQuoteNumber(2)

Quote_AdditionalDataTab()
ClickLogHistoryButton()
ValidateLogHistory()

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest @@ hightlight id_;_Browser("Home").Page("Home 2").WebElement("History")_;_script infofile_;_ZIP::ssf41.xml_;_

