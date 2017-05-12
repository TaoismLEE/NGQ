'================================================
'Project Number: 205713
'User Story: CPQEncoreRetirement_US9408_06: Sbundleusebestpricingshoppinglogic
'Description: This test deals with best pricing shopping logic for S Bundle
'Tags: Quote, SBundle
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\Tbundleusebestpricingshoppinglogic_US9408_06.xlsx"
Dim oEdit
'START: CoreLV
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strBundleId : strBundleId = DataTable.Value("BundleId","Global") 
Dim intBundleQty : intBundleQty = DataTable.Value("BundleQty","Global") 
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 
Dim strReceiptDate : strReceiptDate = DataTable.Value("ReceiptDate","Global") 
Dim strOutputType : strOutputType = DataTable.Value("OutputType","Global") 
Dim dirPath : dirPath = Environment.Value("TestDir") + "\..\.."

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74255", "CPQEncoreRetirement_US9408_06_Sbundleusebestpricingshoppinglogic"

OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption "Save"  
Quote_CaptureQuoteNumber

'Quote_SearchProduct 
Quote_SelectBundle 
Quote_AddBundleToQuote strBundleId

'Quote_refreshPricing
ClickRefreshPricing
Quote_QtyUpdate 1, intBundleQty
Quote_ClickFooter

'Quote_refreshPricing
ClickRefreshPricing

'Generate PDF file
Quote_OutputTab
Quote_SelectIncludeCoverPage
Quote_DeSelectIncludeCoverPage
Quote_SelectCustomGroupView
Quote_DeSelectCustomGroupView
OutputQuote_SetOutputType "Extended Net Price by item with estimated delivery"
OutputQuote_ClickPreview

'SavePdfAs pdfpath
OutputQuote_SaveQuotePdf DataTable.Value("QuoteNumber_Output", "Global")
Dim strSavePath : strSavePath = CreateSavePath(DataTable.Value("QuoteNumber_Output", "Global")) 
Dim pdfObj : Set pdfObj = NewPdfParser(strSavePath)
pdfObj.verifyProductsTable_bundlePricing

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest



