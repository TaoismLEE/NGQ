﻿'================================================
'Summary: Check the UI of My Preference page
'Description: Check all the elements which should be displayed are displayed in My Preference page
'Creator: yu.li9@hpe.com
'Last Modified Time: 3/7/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data.
DataTable.Import "..\..\data\UI_MyPreference.xlsx"
'Open browser
OpenNgq objUser

ClickMyPreferenceUnderAdminTools
CheckDefaultValueOfSelectDefaultCountry
CompareCountryList
CompareAvailableColumnList
CompareAssignedColumnList

FinalizeTest
