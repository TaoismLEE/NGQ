'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"
DataTable.Import "..\..\data\Tbundleusebestpricingshoppinglogic_US9408_05.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")


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


''START: Core

OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
strQuotaSelection_Selector = "Save"
QuoteServices_SelectOption strQuotaSelection_Selector  
Quote_CaptureQuoteNumber

Quote_ValideAddButtonOptions

Quote_SelectBundle 
Quote_AddBundleToQuote strBundleId

Quote_refreshPricing


Quote_DealId
Quote_BundleIdCheck
Quote_CaptureDealId
Quote_OutputTab


Quote_SelectIncludeCoverPage
Quote_SelectCustomGroupView

Quote_DeSelectIncludeCoverPage
Quote_DeSelectCustomGroupView

Quote_PreviewButton
Quote_PreProcessDownload strDownloadDirectory, strDownloadFileName
Quote_SaveFromDownloadBar
Quote_ProcessDownload strDownloadDirectory, strDownloadFileName

Quote_QtyUpdate intBundleQty
Quote_ClickFooter @@ hightlight id_;_2083603328_;_script infofile_;_ZIP::ssf3.xml_;_
'Fetch data.
'Open browser.
'Take over world.

'NOTE: automation API calls only here. No raw UFT calls!
Navbar_Logout
FinalizeTest

