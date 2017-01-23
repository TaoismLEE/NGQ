'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_Identify loose parts within the bundle_09
'Description:	"The case is to validate:
'               1. Validate user can recognize the loose parts within the bundle.
'               2. Just add bundle, not add standalone product, user can not see PQB.
'               3. Add standalone product, user can see PQB."
'Tags:
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
DataTable.Import "..\..\data\TS_NGQ_CPQ_EncoreRetirement_US9408_IdentifyLoosePartsWithinTheBundle_09.xlsx"

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = ""
Dim strQuoteVersion : strQuoteVersion = ""
Dim strQuoteStatus : strQuoteStatus = ""
Dim strQuoteEndDate : strQuoteEndDate = ""

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global") 'modified 20jul
Dim strQuoteName : strQuoteName = DataTable.Value("QuotaName","Global")
Dim strBundleID : strBundleID = DataTable.Value("BundleID","Global")
Dim strProductNumberA : strProductNumberA = DataTable.Value("ProductNumberA","Global")
Dim intProductQuantity : intProductQuantity = DataTable.Value("ProductQuantity","Global")
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = "" 'Valid values: Refresh Pricing, Custom Group, Save
Dim strShowHideColumns_Selector : strShowHideColumns_Selector = "" 'Valid values: Solution ID, Total Requested Discount, My Empowerment Disc %
Dim intIndex : intIndex = 0
' Loose item table items - JH
Dim strLoosePartNum : strLoosePartNum = DataTable.Value("LoosePartNumber","Global")
Dim strLoosePartDesc : strLoosePartDesc = DataTable.Value("LoosePartDesc","Global")
Dim intLoosePartRow : intLoosePartRow = DataTable.Value("LoosePartRow","Global")

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector
Quote_ValidateAddButtonOptions

' Search Bundle ID
'Quote_SearchProduct
'Quote_SearchProductSelectBundle
'Quote_SearchBundleByID strBundleID
'Quote_SearchBundleIncludeGlobalBundles
'Quote_SearchBundleAction
'Quote_SearchBundleValidateBundleID
'Quote_SearchBundleSelectRecord 'new
'Quote_SearchBundleAddBundleToQuote 'new
Quote_SelectBundle
Quote_AddBundleToQuote strBundleID

' Core - Refresh Pricing
'strQuotaSelection_Selector = "Refresh Pricing"
'QuoteServices_SelectOption strQuotaSelection_Selector
Quote_Refresh_Pricing
' hide/show columns functionality removed
' Show Hide Columns popup & option selection 
'LineItemDetails_ShowHideColumns


'WORKAROUND TO GET SOLUTION ID START
lineItemDetails_changeView "SolutionID"
wait 5
'strShowHideColumns_Selector = "Solution ID"
'LineItemDetails_ShowHideColumnsSelection strShowHideColumns_Selector


' Remove columns to display last one
'strShowHideColumns_Selector = "Total Requested Discount"
'LineItemDetails_ShowHideColumnsSelection strShowHideColumns_Selector
'strShowHideColumns_Selector = "My Empowerment Disc %"
'LineItemDetails_ShowHideColumnsSelection strShowHideColumns_Selector
'LineItemDetails_ShowHideColumns
'LineItemDetails_MaximizeRestore

' Validate Loose Parts in Bundle
'Quote_AddBundleLooosePartValidation 
'Quote_AddBundleValidation
lineItemDetails_SolutionIDAsc
lineItemDetails_LooseItemVerification strLoosePartNum, strLoosePartDesc, intLoosePartRow

' Add Products to Quote via Search Product(s)
'Quote_SearchProduct
'Quote_SearchProductByProductNumber strProductNumberA
'Quote_SearchProductSelectResult intProductQuantity
'Quote_SearchProductAddProductsToCart
'Quote_SearchProductAddProdcutsToQuote
'Quote_SearchProductAddProductsToQuoteValidation
lineItemDetails_resetGrid
Quote_SearchProduct
'Enter the product number A 
set_product_number strProductNumberA
set_quantity
add_to_quote

'Validate Products Added
lineItemDetails_SolutionIDAsc
verify_prodTable_prodNum strProductNumberA, 4

Quote_Refresh_Pricing
Quote_Save
Navbar_Logout

Close_Browser
'Navbar_Logout
'CloseBrowser

FinalizeTest
