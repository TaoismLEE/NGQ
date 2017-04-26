'================================================
'Summary: Assign Group Name
'Description: Validate that user can assign group name for products
'Creator: yu.li9@hpe.com
'Last Modified: 4/26/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data
DataTable.Import "..\..\data\Func_AssignGroupName.xlsx"
Dim strProductNum1 : strProductNum1 = DataTable.Value("Product1", "Global")
Dim strProductNum2 : strProductNum2 = DataTable.Value("Product2", "Global")
Dim strGroupName : strGroupName = DataTable.Value("GroupName", "Global")

'Dump jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "000009", "Validate that user can assign group name for products"

'Open browser
OpenNgq objUser
Navbar_CreateNewQuote

'Input two products
LineItemDetails_AddProductByNumber2 strProductNum1
LineItemDetails_AddProductByNumber2 strProductNum2

'Setting group name for first line by righg-clicking
PopUpGroupNameDialog 2
SetGroupName strGroupName

'Check the group name of first product has been changed to prefered value
DisplayGroupName
verify_prodTable_GroupName strGroupName, 1

'Manual set group name and check it
ManualSelectGroupName strGroupName, 2
verify_prodTable_GroupName strGroupName, 2

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest

