'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9400_06: Add Product or Option
'Description:	1. When adding an invalid product, the font color of product number and description is blue.
'				2. When adding an obsolete product, the font color of product number and description is red.
'				3. The new product line item is added at the bottom of the line items before entering the product number.
'				4. The new product line item moves right below the line item right-clicked previously when the product number is entered.
'Tags: Quote, Validate, Color, Text 
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Load test data
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_AddProductOrOption_06.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")

Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strProductNumberInv : strProductNumberInv = DataTable.Value("ProductNumberInv","Global")
Dim strProductNumberObs : strProductNumberObs = DataTable.Value("ProductNumberObs","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global") 
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 

' For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74226", "CPQ_Encore Retirement_US9400_Add Product or Option_06"

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' CPQ_Encore Retirement_US9400_Add Product or Option_06
' Try to add an invalid product number that doesn’t exist in Corona
LineItemDetails_AddProductByNumber strProductNumberInv, 1
'Validate invalid product
LineItemDetails_ValidateProductNonExist strProductNumberInv
'LineItemDetails_ValidateProductNonExistFontColor 2, strProductNumberInv
' Try to add an obsolete product number 
LineItemDetails_AddProductByNumber2 strProductNumberObs
'Validate obsolete product
LineItemDetails_ValidateProductObsolete strProductNumberObs
'LineItemDetails_ValidateProductObsoleteFontColor 3, strProductNumberObs
'Add a valid product number
LineItemDetails_AddProductByNumber2 strProductNumber
'Validate class value for the invalid and obsoleted products
LineItemDetails_ValidateProductNonExistFontColor 2, strProductNumberInv
LineItemDetails_ValidateProductObsoleteFontColor 3, strProductNumberObs
' Add product from Configuration OCS
UFT.ReplayType = 1
build_ocs_bom

' END: Core
Quote_Refresh_Pricing
CloseInformMessage
QuoteServices_SelectOption strQuotaSelection_Selector
VerifySaveButtonColor
Navbar_Logout

Close_Browser
FinalizeTest
