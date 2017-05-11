'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9400_06: Add Product or Option
'Description:	1. When adding an invalid product, the font color of entire line is red
'				2. When adding an obsolete product, the font color of product number is red
'				3. When adding an valid product, the font color of entire line is blue
'				4. The font color of entire line is black when the product belongs to a configurationg or a bundle
'Tags: Quote, Validate, Color, Text
'Last modified: 5/11/2017 by yu.li9@hpe.com
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
Dim strBundleID : strBundleID = DataTable.Value("BundleID","Global")
Dim strBundleProductNumber : strBundleProductNumber = DataTable.Value("BundleProductNumber","Global")
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

' Try to add an invalid product number that doesn’t exist in Corona
LineItemDetails_AddProductByNumber strProductNumberInv,intProductQuantity
LineItemDetails_ValidateProductNonExist strProductNumberInv

'Try to add an obsoleted product
LineItemDetails_AddProductByNumber2 strProductNumberObs
LineItemDetails_ValidateProductObsolete strProductNumberObs

'Try to add a valid product
LineItemDetails_AddProductByNumber2 strProductNumber

'Try to add a bundle
Quote_ManualAddBundle strBundleID,intProductQuantity

'Refreshing price
Quote_Refresh_Pricing
CloseInformMessage

'Validate font color for all kinds of products
LineItemDetails_ValidateProductNonExistFontColor 2, strProductNumberInv
LineItemDetails_ValidateProductObsoleteFontColor 3, strProductNumberObs
LineItemDetails_ValidateValidProductFontColor 4, strProductNumber
LineItemDetails_ValidateBundleProductFontColor 6, strBundleProductNumber

' END: Core
QuoteServices_SelectOption strQuotaSelection_Selector
VerifySaveButtonColor

'Finishing test
Navbar_Logout
Close_Browser
FinalizeTest
