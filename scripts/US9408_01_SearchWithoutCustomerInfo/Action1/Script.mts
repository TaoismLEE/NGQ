﻿'==============================================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_01: Search bundles without customer information entered
'Description:	This case is to validate: When not importing opportunity ID or
'             entering customer ID, sales op is able to search out bundles
'Tags: Search
'==============================================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Test Data
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9408_SearchBundlesWithoutCustomerInformationEntered_01.xlsx" 'added 20jul

'Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""

'Hard-coded data.
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strBundleID : strBundleID = DataTable.Value("BundleID","Global")
Dim intProductQuantity : intProductQuantity = 1
Dim strOpportunityID : strOpportunityID = DataTable.Value("OpportunityID","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74250", "CPQ_Encore Retirement_US9408_Search bundles without customer information entered_01"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityID
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector

'CPQ_Encore Retirement_US9408_Search bundles without customer information entered_01
Quote_SelectBundle
Quote_AddBundleToQuote strBundleID

'END: Core
Quote_Refresh_Pricing

'Validate bundle added
verify_product_table

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
