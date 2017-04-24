'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_02: Search bundles with customer information entered
'Author: Joshua Hunter
'Description: This test deals with testing adding a specific bundle id to a quote
'Tags: Quote, Bundles
'Last Modified: 4/21/2017 by yu.li9@hpe.com
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'ImportTestData strTestDataFile
DataTable.Import "..\..\data\US9408_02SearchBundles.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strBundleID : strBundleID = DataTable.Value("BundleID","Global")
Dim strCustomerID : strCustomerID = DataTable.Value("CustomerID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""

'Dump jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "74251", "CPQ_Encore Retirement_US9408_Search bundles with customer information entered_02"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector

' Validate Customer ID generated
Quote_CustomerDataTab
CustomerData_ValidateCustomerID strCustomerID

' Search Bundle ID and add to quote
Quote_SelectBundle
Quote_AddBundleToQuote strBundleID

'Validate bundle products
Dim arrProducts : arrProducts = GetSourceDataFromExcel
ValidateValidProducts arrProducts

'Refereshing price
Quote_Refresh_Pricing
Quote_Save

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
