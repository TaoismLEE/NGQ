﻿'================================================
'Summary: Check the UI of Advanced Search Page
'Description: Check all the elements which should be displayed are displayed in the advanced page and the options of list
'Creator: yu.li9@hpe.com
'Last Modified Time: 3/6/2017
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
DataTable.Import "..\..\data\UI_AdvancedSearch.xlsx"
'Open browser
OpenNgq objUser
ChangeToAdvancedSearch
CheckDefaultPageOfAdvancedSearch
CompareListOptionsOfAdvancedSearch

FinalizeTest