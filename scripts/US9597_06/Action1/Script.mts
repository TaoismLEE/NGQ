'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<encrypted DigitalBadge>")

DataTable.Import "..\..\data\US9547_06.xlsx"

Dim strOportunityId, strQuoteName, strMCCDisc, strAmount, strMCCNum, strTargReqDiscPercentage

	strOportunityId = DataTable.Value("strOportunityId",1)
	strQuoteName = DataTable.Value("strQuoteName",1)
	strMCCDisc = DataTable.Value("strMCCDisc",1)
	strAmount = DataTable.Value("strAmount",1)
	strMCCNum = DataTable.Value("strMCCNum",1)
	strTargReqDiscPercentage = DataTable.Value("strTargReqDiscPercentage",1)

InitializeTest "IE"

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

