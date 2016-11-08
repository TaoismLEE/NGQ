'================================================
'Test Case: CPQ_Encore Retirement_US9400_Add Product or Option_06
'
'Preconditions:
'1. Sales op has access to NGQ.
'2. An Opportunity ID is ready.
'3. An invalid product number is ready.
'4. An obsolete product number is ready.
'5. A valid product number is ready.
'
'Recommended: Use programing descriptive not objects repository
'Author: Ana Karina Orduña
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "Action1"
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_AddProductOrOption_06.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")

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
' Try to add an obsolete product number 
LineItemDetails_AddProductByNumber strProductNumberObs, 1
'Validate obsolete product
LineItemDetails_ValidateProductObsolete strProductNumberObs
'Add a valid product number
LineItemDetails_AddProductByNumber strProductNumber, 1
'Validate font color
LineItemDetails_ValidateProductNonExistFontColor 2, strProductNumberInv
LineItemDetails_ValidateProductObsoleteFontColor 3, strProductNumberObs

' Add product from Configuration OCS
build_ocs_bom

' END: Core
Quote_Refresh_Pricing
Quote_Save
Navbar_Logout

Close_Browser
FinalizeTest
