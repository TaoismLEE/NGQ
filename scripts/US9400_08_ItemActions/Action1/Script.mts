'================================================
'Project Number: 205713
'User Story: US9400_08_Item Actions
'Description:
	'The case is to validate:
		'1. Sales op is able to add an Page break and Comment item.
		'2. Sales op is able to remove the selected item only when Sub-total has not been applied to the line being removed.
		'3. There is a message popup indicating removing the sub-total when Sales op is trying to remove a part of subtotal.
		'4. Sales op is able to replace an item.
		'5. Sales op is able to promote an item.
		'6. Sales op is able to demote an item.
		'7. Sales op is able to add config from OCS.
'Tags: Remove item, Replace item, Promote item, Demote item
'Author: Reese Childers
'Last Modified: 4/21/2017 by yu.li9@hpe.com
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Load test data
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_ItemActions_08.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim strBaseProduct : strBaseProduct = DataTable.Value("BaseProduct","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 
Dim strDeliverySpeed : strDeliverySpeed = DataTable.Value("DeliverySpeed","Global") 
Dim strDeliveryTerms : strDeliveryTerms = DataTable.Value("DeliveryTerms","Global") 
Dim strLineItemSelector: strLineItemSelector = DataTable.Value("LineItemSelector","Global")

'Jenkins plugin
dumpJenkinsOutput Environment.Value("TestName"), "74228", "CPQ_Encore Retirement_US9400_Item Actions_08"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
click_save_button()

' Build ocs configuration
build_ocs_bom @@ hightlight id_;_Browser("Home").Page("Home 2").Link("NI00155377")_;_script infofile_;_ZIP::ssf2.xml_;_

'Add neccessary data
PreValidate_FixDataCheckErrors

'Refresh Pricing
click_refresh_pricing()

rightClickAddPageBreak()

rightClickAddComment()

selectMultipleLines

AddSubtotal

removeItemInSubTotal

deleteSubTotalLine

removeItem

click_refresh_pricing()

editProductNum strProductNumber

rightClickPromoteItem

rightClickDemoteItem


'Logout and close browser
Navbar_Logout
Close_Browser
FinalizeTest
