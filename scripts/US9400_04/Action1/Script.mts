'================================================
'Summary: US9400_04
'
'Description:Written by Joshua Hunter
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

'InitializeTest "CH"
InitializeTest ""

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "57ac9e99c73adf9d179db3384e132872a103691a0efd")

'DataImport
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"

'set var from data sheets
Dim emptyQuoteNumber : emptyQuoteNumber = DataTable.Value("quoteNumber", "Global")
Dim emptyQuoteVersion : emptyQuoteVersion = DataTable.Value("quoteVersion", "Global")
Dim emptyQuoteStatus : emptyQuoteStatus = DataTable.Value("quoteStatus", "Global")
Dim emptyQuoteEndDate : emptyQuoteEndDate = DataTable.Value("quoteEndDate", "Global")
Dim emptyQuoteSelectedTab : emptyQuoteSelectedTab = DataTable.Value("selectedTab", "Global")

DataTable.Import "..\..\data\NGQ_US9400_04_data.xlsx"

Dim opportunityID : opportunityID = DataTable.Value("Opportunity_ID", "Global")
Dim quoteName : quoteName = DataTable.Value("quoteName", "Global")

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote

NewQuote_ValideEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

Quote_UploadConfig Environment.Value("TestDir") & "\..\..\data\DESSTEPS_335038_CPQ Encore_US9400_04_config.xls"

'Navbar_Logout

'FinalizeTest

