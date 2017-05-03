'================================================
'Summary: Manipulate MCC
'Description: Validate user can add line item MCC, Header MCC
'Creator: yu.li9@hpe.com
'Last Modified: 5/03/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data
DataTable.Import "..\..\data\Func_ManipulateMCC.xlsx"
Dim strOpportunityID : strOpportunityID = DataTable.Value("OpportunityID",1)
Dim strProductNum1 : strProductNum1 = DataTable.Value("Product1",1)
Dim strProductNum2 : strProductNum2 = DataTable.Value("Product2",1)

Dim StrMCC : StrMCC = DataTable.Value("MCC",1)
Dim strMCCNum : strMCCNum = DataTable.Value("MCCNum",1)
Dim strDiscountType : strDiscountType = DataTable.Value("DiscountType",1)
Dim strPercent : strPercent = DataTable.Value("Percent",1)
Dim strHeaderValueType : strHeaderValueType = DataTable.Value("HeaderValueType",1)
Dim strLineItemValue : strLineItemValue = DataTable.Value("LineValue",1)
Dim strHeaderValue : strHeaderValue = DataTable.Value("HeaderValue",1)

'Dump jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "000011", "Validate user can add line item MCC, Header MCC"

'Open browser
OpenNgq objUser

'Start a new quote and import opportunity
Navbar_CreateNewQuote
OpportunityAndQuoteInfo_SetOpportunityId strOpportunityID
OpportunityAndQuoteInfo_Import

'Fill neccessary data
PreValidate_FixDataCheckErrors

'Add two products
LineItemDetails_AddProductByNumber2 strProductNum1
LineItemDetails_AddProductByNumber2 strProductNum2

'Apply a line item MCC for the first line product
SelectFirstLineProduct 2
PopUpLineItemMCCDialog 2
ApplyLineItemMCC StrMCC, strDiscountType, strPercent, strLineItemValue
DisplayMCC
ValidateLineItemMCC 2, strLineItemValue

'Apply a header level MCC for the quote
Quote_PricingTermsTab
Utils_scrollToBottom_lineItemAddColumn
requestOPDisc_MCC StrMCC
requestOPDisc_discType strDiscountType
requestOPDisc_valueType strHeaderValueType
RequestOPDisc_percentage strPercent
RequestOPDisc_amount strHeaderValue
RequestOPDisc_Submit
MCC_success_message strMCCNum

'Validate Header MCC for all line items
ValidateLineItemMCC 2, strHeaderValue
ValidateLineItemMCC 3, strHeaderValue

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest

