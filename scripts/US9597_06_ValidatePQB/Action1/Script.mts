'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9597_06: show Price Quality Band
'Description:	This case is to validate:
'				1. Sales op is able to view Price Quality Band (PQB).
'				2. The Current Band color at the header level and the traffic lights of all products are changed according to the PQB when changing the discount percentage.
'Tags: PQB, BOM, Product, Pricing
'Last Modified: 5/15/2017 by yu.li9@hpe.com
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\US9547_06.xlsx"
Dim strOportunityId, strQuoteName, strMCCDisc, strAmount, strMCCNum, strTargReqDiscPercentage
strOportunityId = DataTable.Value("strOportunityId",1)
strQuoteName = DataTable.Value("strQuoteName",1)
strMCCDisc = DataTable.Value("strMCCDisc",1)
strAmount = DataTable.Value("strAmount",1)
strMCCNum = DataTable.Value("strMCCNum",1)
strTargReqDiscPercentage = DataTable.Value("strTargReqDiscPercentage",1)

' For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74273", "CPQ_Encore Retirement_US9597_ show Price Quality Band_06"

'Open browser
OpenNgq(objUser)

'go to new Quote navbar
Navbar_CreateNewQuote()

'Validate information in New Quote 
NewQuote_ValidateEmptyQuote null,null,null,null

'set opportunity ID
OpportunityAndQuoteInfo_SetOpportunityId(strOportunityId)

'Click Import button
OpportunityAndQuoteInfo_Import()

'Set Quote Name
Quote_EditQuoteName(strQuoteName)

'Save Import
Quote_save

'Scroll down
pageDownNewQuotePage()

'Add config from ocs
build_ocs_bom

'Click refresh pricing
ClickRefreshPricing()

'Click in pricing and term tab
ClickPricingTermsTab()

'Add a MCC disacount
requestOPDisc_MCC(strMCCDisc)
RequestOPDisc_amount(strAmount)
RequestOPDisc_Submit()

'validate the discount
MCC_success_message(strMCCNum)

'Click refresh pricing
ClickRefreshPricing

'Save Import
Quote_save

'Auto Allocation total disc and apply
SetAutoAllocTargReqDiscPercentage(strTargReqDiscPercentage)

'Click refresh pricing
ClickRefreshPricing()

'logout and close the browser
Navbar_Logout
Close_Browser
FinalizeTest
