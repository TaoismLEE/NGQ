'================================================
'Product Number:205713
'User Story: CPQ_Encore Retirement_US9430_Transfer a quote to another NGQ user in the same group_04
'Description: This case is to validate:
'			1.Sales Op is able to transfer a quote to another NGQ user in the same group.
'Tags: Quote, Transfer, Group, 
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'DataTable.Import "..\..\data\US9430_04.xlsx"
'Dim strQuoteNumber : strQuotenumber = DataTable("strQuotenumber",1)

InitializeTest "Action1"

' For Jenkins Reporting
dumpJenkinsOutput "US9430_04_TransferQuote", "74249", "CPQ_Encore Retirement_US9430_Transfer a quote to another NGQ user in the same group_04"

'open web browser and go to NGQ/My Dashboard
OpenNgq(objUser)
ClickMyDashboard()

'validate if QuoteTab is selected
ValidateQuoteTab()
'To get the first quote
Dim strQuote: strQuote = GetFirstQuoteNumberofMyQuote(2)

'Click auto Filter Button
ClickAutoFilter()

'set and submit Quote Number
FillFilterQuoteNumber(strQuote)  'NI00159743
'FillFilterQuoteNumber(strQuotenumber)
ClickQuoteNumber(2)
'Validate the submit value match with the value in table
ValidateQuoteNumberValue(strQuote)
'ValidateQuoteNumberValue(strQuotenumber)

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest

