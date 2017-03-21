'================================================
'Product Number:205713
'User Story: CPQ_Encore Retirement_US9430_04: Transfer a quote to another NGQ user in the same group
'Description: This case is to validate:
'			1.Sales Op is able to transfer a quote to another NGQ user in the same group.
'Tags: Quote, Transfer, Group
'Author: yu.li9@hpe.com
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Load test data
DataTable.Import "..\..\data\US9430_04.xlsx"
Dim strTargetUser : strTargetUser = DataTable.Value("TargetUser",1)
Dim strTransferReason : strTransferReason = DataTable.Value("TransferReason",1)
Dim strUserGroup : strUserGroup = DataTable.Value("UserGroup",1)

InitializeTest "Action1"

' For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74249", "CPQ_Encore Retirement_US9430_Transfer a quote to another NGQ user in the same group_04"

'open web browser and go to NGQ/My Dashboard
OpenNgq objUser
ClickMyDashboard

'validate if QuoteTab is selected
ValidateQuoteTab
'To get the first quote
Dim strQuote: strQuote = GetFirstQuoteNumberofMyQuote(2)

'Click auto Filter Button
ClickAutoFilter

'Set the first row quote
FillFilterQuoteNumber(strQuote)

'Check the check item
CheckFirstRowQuote(2)

'Click transfer button
Click_TransferOwnership

'Set values for transfer
SelectTransferOwnership_TransferReason strTransferReason
SelectTransferOwnershipGroup strUserGroup
SelectTransferEmail strTargetUser
Click_TransferContinue

'Inform message pops up
VerifyQuoteTransfer strTargetUser

'Click home, then click my dashboard
ClickHome
ClickMyDashboard

'Click My Group Quote tab and first status number
ClickMyGroupQuoteTab
ClickMyGroupStatusCount

'Click auto Filter Button
ClickAutoFilter

'Set the filter quote the same with the tansfered quote
FillFilterQuoteNumber(strQuote)

'Get the result
Dim strQuoteNumber : strQuoteNumber = GetFirstQuoteNumberofMyGroupQuote(2)

'Make sure they are the same
CompareTwoQuote strQuoteNumber, strQuote

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest

