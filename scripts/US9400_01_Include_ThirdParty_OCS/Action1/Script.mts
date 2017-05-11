'================================================
'Project Number: 205713
'User Story:  US9400_01_Include Third Party Parts in the Configuration from OCS
'Description:
' The case is to validate:
'	1. Sales op is able to add a configuration from OCS to WNGQ.
'	2. A third party product can be included in this configuration.
'Tags OCS, Third Party Product
'Last Modified: 5/11/2017 by yu.li9@hpe.vom
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_Include ThirdPartyPartsInTheConfigurationFromOCS_01.xlsx"
'Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global") 
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global") 
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 

'Jenkins information
dumpJenkinsOutput Environment.Value("TestName"), "74218", "CPQ_Encore Retirement_US9400_Include Third Party Parts in the Configuration from OCS_01 "

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

'Add third party product number
LineItemDetails_AddProductByNumber strProductNumber, 1

'Add product from Configuration OCS
build_ocs_bom

'END: Core
Quote_Refresh_Pricing
Quote_Save

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
