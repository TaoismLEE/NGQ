'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9408_T bundle use best pricing shopping logic_05
'Author: Latha Venkataram
'Description: This test deals with best pricing shopping logic for T Bundle
'Tags: Quote, TBundle
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

InitializeTest "Action1"
DataTable.Import "..\..\data\Tbundleusebestpricingshoppinglogic_US9408_05.xlsx"

'Hard-coded data.
'Dim objUser : Set objUser = NewRealUser("Yuudachi@nightmare.sb", "poipoipoipoi","<Encrypted DigitalBadge>")


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
Dim strDownloadDirectory : strDownloadDirectory = DataTable.Value("DownloadDirectory","Global")
Dim strDownloadFileName : strDownloadFileName = DataTable.Value("DownloadFileName","Global")
Dim dirPath : dirPath = Environment.Value("TestDir") + "\..\.."

''START: Core
dumpJenkinsOutput "US9408_05", "74254", "CPQ_Encore Retirement_US9408_T bundle use best pricing shopping logic_05"
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector  
Quote_CaptureQuoteNumber


Quote_SelectBundle 
Quote_AddBundleToQuote strBundleId

Quote_refreshPricing

'Doesnt work
'Quote_DealId
'Quote_BundleIdCheck
'Quote_CaptureDealId
'To make sure the DealID colomn is displayed
DisplayDealId
'new capture function for DealId
Dim dealId : dealId = get_prodTable_dealId(2)

Quote_OutputTab

Quote_SelectIncludeCoverPage
Quote_SelectCustomGroupView

Quote_DeSelectIncludeCoverPage
Quote_DeSelectCustomGroupView

OutputQuote_SetOutputType "Extended Net Price by item with estimated delivery"

Quote_CaptureQuoteNumber


OutputQuote_ClickPreview

'dim pdfPath : pdfpath = dirPath + "\data\depends\" + DataTable.Value("QuoteNumber_Output", "Global") + ".pdf"
'SavePdfAs pdfpath
OutputQuote_SaveQuotePdf DataTable.Value("QuoteNumber_Output", "Global") 'strQuoteNumber
Dim strSavePath : strSavePath = CreateSavePath(DataTable.Value("QuoteNumber_Output", "Global"))
Dim pdfObj : Set pdfObj = NewPdfParser(strSavePath)
pdfObj.verifyProductsTable_bundlePricing
'Doesnt work
'Quote_PreProcessDownload strDownloadDirectory, strDownloadFileName
'Quote_SaveFromDownloadBar
'Quote_ProcessDownload strDownloadDirectory, strDownloadFileName
Quote_QtyUpdate 1, intBundleQty
' Doesnt work
'Quote_ClickFooter @@ hightlight id_;_2083603328_;_script infofile_;_ZIP::ssf3.xml_;_
'Fetch data.
'Open browser.
'Take over world.

'NOTE: automation API calls only here. No raw UFT calls!
Navbar_Logout
FinalizeTest

