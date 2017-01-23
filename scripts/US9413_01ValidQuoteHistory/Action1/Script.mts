'================================================ @@ hightlight id_;_2426184_;_script infofile_;_ZIP::ssf4.xml_;_
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9413_Version Quote and View History_Valid Quote_01
'Author: Joshua Hunter
'Description: This test deals with testing the ability to create new versions of a quote
'Tags: Quote, NewVersion
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'InitializeTest "CH"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'DataImport
'DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"

'set var from data sheets
Dim emptyQuoteNumber : emptyQuoteNumber = DataTable.Value("quoteNumber", "Global")
Dim emptyQuoteVersion : emptyQuoteVersion = DataTable.Value("quoteVersion", "Global")
Dim emptyQuoteStatus : emptyQuoteStatus = DataTable.Value("quoteStatus", "Global")
Dim emptyQuoteEndDate : emptyQuoteEndDate = DataTable.Value("quoteEndDate", "Global")
Dim emptyQuoteSelectedTab : emptyQuoteSelectedTab = DataTable.Value("selectedTab", "Global")

DataTable.Import "..\..\data\NGQ_US9413_01_data.xlsx"

Dim opportunityID : opportunityID = DataTable.Value("Opportunity_ID", "Global")
Dim quoteName : quoteName = DataTable.Value("quoteName", "Global")
dumpJenkinsOutput "US9413_01", "74231", "CPQ_Encore Retirement_US9413_Version Quote and View History_Valid Quote_01"
'Open browser.
OpenNgq objUser
'browser("NGQ").FullScreen
Navbar_CreateNewQuote

NewQuote_ValidateEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

Dim quoteID : QuoteID = Quote_get_quoteNumber

build_ocs_bom

Quote_refreshPricing

Dim grand_total : grand_total = get_grand_total

quote_newVersion quoteID

LineItemDetails_AddProductByNumber_manualQtyRow DataTable.Value("productID", "Global"), 1, 8 'Workaround for NGQ not locking their product quantity rows like normal

'msgbox Err.Number

DataTable.SetNextRow
Quote_refreshPricing
Quote_GrandTotal_Changed grand_total

quote_newVersion quoteID
'Quote_save


Navbar_Home

Navbar_QuickSearch quoteID

verify_advSearch_table quoteID

Navbar_Logout

FinalizeTest

'Browser("NGQ").Close
