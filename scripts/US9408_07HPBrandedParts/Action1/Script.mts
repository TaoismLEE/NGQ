'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_Add HP branded third party parts in any configuration_07
'Author: Joshua Hunter
'Description: This test deals with testing adding OCS item with HP Branded Parts
'Tags: Quote, OCS
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
DataTable.Import "..\..\data\US9408_07HPBrandedParts.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = ""
' For Jenkins Reporting
dumpJenkinsOutput "US9408_07", "74256", "CPQ_Encore Retirement_US9408_Add HP branded third party parts in any configuration_07"
'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' Add third party product number
'Quote_AddProductOrOption
'LineItemDetails_SetProductNumberByIndex 1, strProductNumber
'Quote_SetBundleID strProductNumber
'LineItemDetails_SetProductQuantityByIndex 0, intProductQuantity
LineItemDetails_AddProductByNumber strProductNumber, 1

' Add product from Configuration OCS
'Quote_AddConfigOCS
'Quote_SelectConfigOCS
'Quote_SaveAndConvertToQuote
build_ocs_bom
' END: Core
Quote_Refresh_Pricing
Quote_Save

Close_Browser
FinalizeTest
