'================================================
'Test case: US9425_02
'
'Summary: Send an email with quotes details and PDF_Complete Quote
'
'Description: This case is to validate:
'               1. The user is able to Preview the PDF generated and send 
'                  the PDF as attachment in the email after quote completed.
'
'Pre-condition:
'    1. Sales ops have access to NGQ.
'    2. Sales ops have a valid opportunity ID
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"

'Fetch data.
DataTable.Import("..\..\data\NGQ_US9425_02_data.xlsx")
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>",  "<Encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable("OpportunityId")
Dim strQuoteName : strQuoteName = DataTable("QuoteName")
Dim strDeliverySpeed : strDeliverySpeed = DataTable("DeliverySpeed")
Dim strPdfOutputType : strPdfOutputType = DataTable("pdfOutputType")

'Open browser.
OpenNgq objUser

'Variables for create new quote validation
Dim strQuoteNumberID
Dim strQuoteVersion
Dim strQuoteStatus
Dim strQuoteEndDate
Dim strQuoteTabSelected : strQuoteTabSelected = "Opportunity and Quote Info"

'Navigate to Create New Quote
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

'Go to "Shipping Data" tab
Quote_ShippingDataTab

'Choose a delivery speed from the drop down list menu
ShippingData_SetDeliverySpeed strDeliverySpeed

'This step is not in the script, but it removes the error in step 27 "Errors in Price section"
Reporter.ReportEvent micWarning, "Step not in script", "This step is not in the test case, but it is required to remove the error from step 27. This step adds a ship to address by selecting ""Same as Sold to Address"" and selects the ""Terms Of Delivery"""
ShippingData_SetTermsOfDelivery "Carriage Paid To"
Quote_CustomerDataTab
CustomerData_ShipToTab
If Not CustomerDataShipTo_SameAsSoldToAddressIsSelected Then
	CustomerDataShipTo_SelectSameAsSoldToAddress
End If
' End of additional steps that were not in the script

'Go to "Additional Data" tab
Quote_AdditionalDataTab
AdditionalData_SetReceiptDateNow

'Click refresh pricing
Dim strQuoteSelector : strQuoteSelector = "Refresh Pricing"
QuoteServices_SelectOption strQuoteSelector
ValidatePriceRefreshed

'Save the quote
Quote_save

'Select pre-validate from the drop down menu
SelectPreValidate

'Validate that there are no errors in Data Check, CLIC, Price, Bundle
PreValidate_DataCheckNoErrors
PreValidate_ClicNoErrors 'Overrides error to get rid of it
PreValidate_PriceNoErrors
PreValidate_BundleNoErrors

'Complete the quote
PreValidate_ClickCompleteQuote
PreValidate_CloseValidationPage

'Navigate to quote output tab
Quote_QuoteOutputTab

'Select the pdf type that will be examined
Dim strPdfType : strPdfType = strPdfOutputType
OutputQuote_SetOutputType strPdfType

'Click on the preview button to generate the pdf
OutputQuote_ClickPreview
 @@ hightlight id_;_Browser("Home#/selfservicequote/createq").Page("Home#/selfservicequote/createq 5").WebElement("QT OP")_;_script infofile_;_ZIP::ssf1.xml_;_
OutputQuote_SaveQuotePdf strQuoteNumber

PdfVerification DataTable("OutputFilePath"), DataTable("PdfCheckSoldTo"), DataTable("PdfCheckShipTo"), DataTable("PdfCheckSalesContact"), _
                DataTable("PdfCheckLineItems"), DataTable("PdfCheckGrandTotal"), DataTable("PdfCheckExtComment"), DataTable("PdfCheckHeader")

Navbar_Logout
Browser("NGQ").Close()

FinalizeTest
