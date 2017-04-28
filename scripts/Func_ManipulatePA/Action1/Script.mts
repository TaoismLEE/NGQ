'================================================
'Summary: PA Manipulation
'Description: Apply PA, Apply header level PA to override it and delete it finally
'Creator: yu.li9@hpe.com
'Last Modified: 4/28/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data
DataTable.Import "..\..\data\Func_ManipulatePA.xlsx"
Dim strOppID : strOppID = DataTable.Value("OpportunityID", "Global")
Dim strProductNum : strProductNum = DataTable.Value("Product", "Global")
Dim strManualPA : strManualPA = DataTable.Value("ManualPA", "Global")
Dim strHeaderPA : strHeaderPA = DataTable.Value("HeaderPA", "Global")

dumpJenkinsOutput Environment.Value("TestName"), "0000010", "Apply PA, Apply header level PA to override it and delete it finally"

'Open browser
OpenNgq objUser
Navbar_CreateNewQuote

'Import Opportunity id
OpportunityAndQuoteInfo_SetOpportunityId strOppID
OpportunityAndQuoteInfo_Import

'Add the product
LineItemDetails_AddProductByNumber2 strProductNum
click_refresh_pricing

'Add a PA manually
SelectFirstLineProduct 2
AddLintItemPA strManualPA, 2
click_refresh_pricing
DisplayPANumberColumn
CheckPAWhetherInLineItemDetailPage2 strManualPA

'Add a header level PA to override the previously added PA
Quote_PricingTermsTab
AddHeaderLevelPA strHeaderPA
click_refresh_pricing
CheckPAWhetherInLineItemDetailPage2 strHeaderPA

'Remove the PA
SelectFirstLineProduct 2
RemoveManualAddedPA 2
click_refresh_pricing
CheckPAWhetherInLineItemDetailPage strHeaderPA

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest

