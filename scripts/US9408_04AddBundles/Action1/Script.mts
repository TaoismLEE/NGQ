'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_Add bundles from line item grid with customer information enterd_04
'Author: Joshua Hunter
'Description: This test deals with testing adding a bundle to a quote
'Tags: Quote, Bundle
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Test Data
'Fill path and file with its extension (C:\ngq-demo-develop\data\fileName.xlsx)
'ImportTestData strTestDataFile
DataTable.Import "..\..\data\US9408_04AddBundles.xlsx" 'added 20jul

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global") 'modified 20jul
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strBundleID : strBundleID = DataTable.Value("BundleID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
' For Jenkins Reporting
dumpJenkinsOutput "US9408_04", "74253", "CPQ_Encore Retirement_US9408_Add bundles from line item grid with customer information enterd_04" 
'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' CPQ_Encore Retirement_US9408_Add bundles from line item grid with customer information enterd_04
' Add Bundle ID
'Quote_AddBundleAction
'Quote_SetBundleID strBundleID
Quote_SelectBundle
Quote_AddBundleToQuote strBundleID
' END: Core
Quote_refreshPricing

Quote_save

Close_Browser
FinalizeTest
