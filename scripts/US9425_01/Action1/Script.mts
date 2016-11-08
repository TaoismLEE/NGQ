'================================================
'Test case: US9425_01
'
'Summary: Send an email with quotes details and PDF of non-complete Quote
'
'Description: This case is to validate:
'               1. The user is able to Preview the PDF generated and send 
'                  the PDF as attachment in the email before quote completed.
'
'Pre-condition:
'    1. Sales ops have access to NGQ.
'    2. Sales ops have a valid opportunity ID
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "Action1"

'Fetch data.
DataTable.Import("..\..\data\NGQ_US9425_01_data.xlsx")
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable("OpportunityId")
Dim strQuoteName : strQuoteName = DataTable("QuoteName")
Dim strPdfOutputType : strPdfOutputType = DataTable("pdfOutputType")

'Open browser.
OpenNgq objUser

'Navigate to create new quote
Dim strQuoteNumberID
Dim strQuoteVersion
Dim strQuoteStatus
Dim strQuoteEndDate
Dim strQuoteTabSelected : strQuoteTabSelected = "Opportunity and Quote Info"

'Navigate to Create quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
Quote_currentlySelectedTab(strQuoteTabSelected)

'Import opportunity ID
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
OpportunityIdIsValid

'Edit quote name
Dim strQuoteNumber
Quote_EditQuoteName strQuoteName

'Click save and verify that save notification and quote number appear
Quote_save
strQuoteNumber = Quote_get_quoteNumber

build_ocs_bom

'Fix errors/warnings that will be present in steps 21, 23
Reporter.ReportEvent micWarning, "Step not in script", "This step is not in the test case, but it is required to remove the errors from pre-validation."
PreValidate_FixDataCheckErrors

'Click refresh pricing
Dim strQuoteSelector : strQuoteSelector = "Refresh Pricing"
QuoteServices_SelectOption strQuoteSelector
ValidatePriceRefreshed

'Select pre-validate from the drop down menu
SelectPreValidate

'Validate that there are no errors in Data Check, CLIC, Price, Bundle
PreValidate_DataCheckNoErrors
PreValidate_ClicNoErrors 'Overrides error to get rid of it
PreValidate_PriceNoErrors
PreValidate_BundleNoErrors
PreValidate_CloseValidationPage

'Navigate to quote output tab
Quote_QuoteOutputTab

'Select the pdf type that will be examined
Dim strPdfType : strPdfType = strPdfOutputType
OutputQuote_SetOutputType strPdfType

'Click on the preview button to generate the pdf
OutputQuote_ClickPreview @@ hightlight id_;_Browser("Home#/selfservicequote/createq").Page("Home#/selfservicequote/createq 5").WebElement("QT OP")_;_script infofile_;_ZIP::ssf1.xml_;_
OutputQuote_SaveQuotePdf strQuoteNumber

PdfVerification DataTable("OutputFilePath"), DataTable("PdfCheckSoldTo"), DataTable("PdfCheckShipTo"), DataTable("PdfCheckSalesContact"), _
                DataTable("PdfCheckLineItems"), DataTable("PdfCheckGrandTotal"), DataTable("PdfCheckExtComment"), DataTable("PdfCheckHeader")

Navbar_Logout
Browser("NGQ").Close()

FinalizeTest
