'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9597_ show Price Quality Band_06
'Description:	This case is to validate:
'				1. Sales op is able to view Price Quality Band (PQB).
'				2. The Current Band color at the header level and the traffic lights of all products are changed according to the PQB when changing the discount percentage.
'Tags: PQB, BOM, Product, Pricing 
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

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
dumpJenkinsOutput "US9597_06_ValidatePQB", "74273", "CPQ_Encore Retirement_US9597_ show Price Quality Band_06"

InitializeTest "Action1"

'Open browser and go to My Dashboard
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
Quote_save()

'Scroll down
pageDownNewQuotePage()

'Click add btn and add config from ocs
'ClickAddConfigOcs() - DOESNT WORK
'click_lineitem_add_ocs

'Process inside OCS
'Ocs_SelectAndConfigureProduct()
'Ocs_SaveBom()
'Ocs_SaveBomValid()
'Ocs_ClickConvertToQuote()
build_ocs_bom
'Click refresh pricing
ClickRefreshPricing()

'Click in pricing and term tab
ClickPricingTermsTab()

requestOPDisc_MCC(strMCCDisc)
 
RequestOPDisc_amount(strAmount)

'submint discount
RequestOPDisc_Submit()

'validate the discount
MCC_success_message(strMCCNum)

'Click refresh pricing
ClickRefreshPricing()

'Save Import
Quote_save()

'Auto Allocation total disc and apply
SetAutoAllocTargReqDiscPercentage(strTargReqDiscPercentage)

'Click refresh pricing
ClickRefreshPricing()

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()
FinalizeTest

