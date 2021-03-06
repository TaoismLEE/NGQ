﻿'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9400_04: Upload Config
'Description: This test uploads a spreadsheet and verifies product table contains all products
'Tags: Quote, Upload
'================================================

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'DataImport
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'set var from data sheets
Dim emptyQuoteNumber : emptyQuoteNumber = DataTable.Value("quoteNumber", "Global")
Dim emptyQuoteVersion : emptyQuoteVersion = DataTable.Value("quoteVersion", "Global")
Dim emptyQuoteStatus : emptyQuoteStatus = DataTable.Value("quoteStatus", "Global")
Dim emptyQuoteEndDate : emptyQuoteEndDate = DataTable.Value("quoteEndDate", "Global")
Dim emptyQuoteSelectedTab : emptyQuoteSelectedTab = DataTable.Value("selectedTab", "Global")

DataTable.Import "..\..\data\NGQ_US9400_04_data.xlsx"

Dim opportunityID : opportunityID = DataTable.Value("Opportunity_ID", "Global")
Dim quoteName : quoteName = DataTable.Value("quoteName", "Global")
Dim uploadConfigPath : uploadConfigPath = Environment.Value("TestDir") & "\..\..\data\" & DataTable.Value("uploadConfigPath", "Global")
dumpJenkinsOutput Environment.Value("TestName"), "74224", "CPQ_Encore Retirement_US9400_Upload Config_04"

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote

NewQuote_ValidateEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

UFT.BrowserNavigationTimeout = 180000
Quote_UploadConfig uploadConfigPath
UFT.BrowserNavigationTimeout = 60000

verify_product_table

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
'file to uploado: C:\Users\rosaljah\OneDrive - Hewlett Packard Enterprise\TAO\ngq-demo-develop\ngq-demo-develop - Rosales\NGQ\data\
'file name: US9400-04-Configuration.xlsx
'. OPE-0005373487
