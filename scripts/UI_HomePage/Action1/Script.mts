'================================================
'Summary: Check the UI of Home Page
'Description: Check all the elements which should be displayed are displayed in the home page
'Creator: yu.li9@hpe.com
'Last Modified Time: 2/28/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
Dim strLoginUser : strLoginUser = DataTable.Value("user", "Global")

'Fetch data.
DataTable.Import "..\..\data\UI_HomePage.xlsx"
Dim strSystemName : strSystemName = DataTable.Value("SystemName","Global")

'Open browser
OpenNgq objUser

'Check the displayed elements
CheckHeaderOfHomePage strSystemName, strLoginUser
CheckQuickSearchOfHomePage
CheckDashboardDocumentPhotoOfHomePage

'Compare all the lists of home page
CompareListOptionsOfHomePage
Navbar_Logout
Close_Browser
FinalizeTest
