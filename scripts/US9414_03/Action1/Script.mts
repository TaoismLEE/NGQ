'================================================
'Project Number: 205713
'User Story: US9414_03
'Description: Capture comments when Transfer a quote
'Author : Pramesh Bhandari
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

Dim strQuote
Dim strEmail
Dim strReason
Dim strGroup
Dim strOutputSheet
Dim strOutputFilePath

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Import Data
DataTable.Import "..\..\data\US9414_03.xlsx"
'strQuote = DataTable("QuoteNumber", dtGlobalSheet)	
strEmail = DataTable("Email", dtGlobalSheet)	
strReason = DataTable("TransferReason", dtGlobalSheet)
strGroup = DataTable("TransferGroup", dtGlobalSheet)

'Creating output sheet
strOutputSheet = "US9414_03_Output"
DataTable.AddSheet strOutputSheet
DataTable.GetSheet(strOutputSheet).AddParameter "QuoteNumber", ""

InitializeTest "Action1"
'Opens browser and ngq website
OpenNgq objUser

'Clicks My Dashboard tab
Click_MyDashboard

'Makes sure the Quote tab is opened
ValidateQuoteTab

'Click on filte under my recent Quote
ClickAutoFilter

'Fetch the first quote in my quotes
strQuote = GetFirstQuoteNumberofMyQuote(2)

'FillFilterQuoteNumber strQuote

'====================================
SetAutoFilterQuoteNumber strQuote
'=================================
'Selects the whole row
Check_RadioButton strQuote

DataTable("QuoteNumber", strOutputSheet) = strQuote
'Clicks the transfer ownership Button
Click_TransferOwnership

'Selects the transfer reason
SelectTransferOwnership_TransferReason strReason

'Set text to the transfer ownership editbox
SelectTransferOwnershipGroup strGroup

'Selects the transfer email
SelectTransferEmail strEmail

'Clicks continue button of transfer ownership window
Click_TransferContinue
'No Ok button to proceed==========================

'Click Advanced Serach
AdvancedSearchClick

''Set Quote Number in a Quote Number Text field under Advanced Search Tab
SetQuoteNumber_AdvancedSearch strQuote 

'Clicks Search button under Advanced search tab
ClickSearch_advancedSearch

VerifyEmailQuote strEmail, strQuote

'Export output data to excel sheet in test script dir
strOutputFilePath = Environment.Value("TestDir") & "\Output" & strOutputSheet & ".xls"
DataTable.ExportSheet strOutputFilePath, strOutputSheet

'Logoff NGQ
Navbar_Logout

'Close the browser
Close_Browser

'Skipped all the steps from 22 as per the aggrement with business people

FinalizeTest


