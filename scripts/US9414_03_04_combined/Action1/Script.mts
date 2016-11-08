'=================================================================
'Project Number: 205713
'User Story: US9414_03
'Description: Transfer a quote and sends confrimation email to both transferer and transferee; Validates the transfer of quote by 
'searching same quote in advanced search; User B have to validate 
'the specific comments manually by clicking the link in User B email
'=================================================================
Option Explicit
Dim al : Set al = NewActionLifetime
Dim strQuote
Dim strEmail
Dim strReason
Dim strGroup
Dim strOutputSheet
Dim strOutputFilePath
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")

'Import Data
DataTable.Import "..\..\data\US9414_03\InputFiles\US9414_03.xlsx"
strQuote = DataTable("QuoteNumber", dtGlobalSheet)	
strEmail = DataTable("Email", dtGlobalSheet)	
strReason = DataTable("TransferReason", dtGlobalSheet)
strGroup = DataTable("TransferGroup", dtGlobalSheet)

'Creating output sheet
strOutputSheet = "US9414_03_Output"
DataTable.AddSheet strOutputSheet
DataTable.GetSheet(strOutputSheet).AddParameter "QuoteNumber", ""

InitializeTest "US9414_03"
'Opens browser and ngq website
OpenNgq objUser

'Clicks My Dashboard tab
Click_MyDashboard

'Makes sure the Quote tab is opened
ValidateQuoteTab

'Click on filte under my recent Quote
ClickAutoFilter

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
'TODO FIX THIS JOSH
VerifyEmailQuote strEmail, strQuote

'Export output data to excel sheet in test script dir
strOutputFilePath = Environment.Value("TestDir") & "\..\..\data\US9414_01\OutputFiles\" & strOutputSheet & ".xlsx"
DataTable.ExportSheet strOutputFilePath, strOutputSheet

Navbar_Home
'Logoff NGQ
Navbar_Logout

'Close the browser
Close_Browser

'Skipped all the steps from 22 as per the aggrement with business people

'FinalizeTest



