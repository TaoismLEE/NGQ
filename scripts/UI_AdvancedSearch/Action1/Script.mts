'================================================
'Summary: Check the UI of Advanced Search Page
'Description: Check all the elements which should be displayed are displayed in the advanced search page and the options of list
'Creator: yu.li9@hpe.com
'Last Modified Time: 4/17/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data.
DataTable.Import "..\..\data\UI_AdvancedSearch.xlsx"

'Dump jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "000001", "Check all the elements which should be displayed are displayed in the advanced page and the options of list"

'Open browser
OpenNgq objUser
ChangeToAdvancedSearch
CheckDefaultPageOfAdvancedSearch
CompareListOptionsOfAdvancedSearch

Navbar_Logout
Close_Browser
FinalizeTest
