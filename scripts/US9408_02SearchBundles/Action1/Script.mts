'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_02: Search bundles with customer information entered
'Author: Joshua Hunter
'Description: This test deals with testing adding a specific bundle id to a quote
'Tags: Quote, Bundles
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
DataTable.Import "..\..\data\US9408_02SearchBundles.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<encrypted digitalbadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strBundleID : strBundleID = DataTable.Value("BundleID","Global")
Dim strCustomerID : strCustomerID = DataTable.Value("CustomerID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""

dumpJenkinsOutput Environment.Value("TestName"), "74251", "CPQ_Encore Retirement_US9408_Search bundles with customer information entered_02"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' Validate Customer ID
Quote_CustomerDataTab
'This functionality is not implemented, test will fail until implemented
'CustomerData_ValidateCustomerID strCustomerID

' Search Bundle ID
'Quote_SearchProduct
'Quote_SearchProductSelectBundle
'Quote_SearchBundleByID strBundleID
'Quote_SearchBundleIncludeGlobalBundles
'Quote_SearchBundleAction
'Quote_SearchBundleValidateBundleID
'Quote_SearchBundleSelectRecord
'Quote_SearchBundleAddBundleToQuote
Quote_SelectBundle
Quote_AddBundleToQuote strBundleID

Quote_Refresh_Pricing
Quote_Save
Navbar_Logout

Close_Browser
FinalizeTest
