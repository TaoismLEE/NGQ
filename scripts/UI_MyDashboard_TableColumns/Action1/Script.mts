﻿'================================================
'Summary: Check the listed columns of My Dashboard page
'Description: Check all the columns which should be displayed are displayed in My Dashboard page
'Creator: yu.li9@hpe.com
'Last Modified Time: 5/8/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data.
DataTable.Import "..\..\data\UI_MyDashboard_TableColumns.xlsx"

dumpJenkinsOutput Environment.Value("TestName"), "000003", "Check all the columns which should be displayed are displayed in My Dashboard page"

'Open browser
OpenNgq objUser
'Check the listed columns of status table
ChangeToMyDashboard
CheckStatusColumnsExist

Dim arrSource : arrSource = GetSourceDataFromExcel
'Check the listed columns of result table
CheckEachExists arrSource

'Check the listed columns of status table of My group page
ClickMyGroupQuoteTab
CheckStatusColumnsExist

'Check the listed columns of result table
CheckEachExists arrSource

Navbar_Logout
Close_Browser
FinalizeTest
