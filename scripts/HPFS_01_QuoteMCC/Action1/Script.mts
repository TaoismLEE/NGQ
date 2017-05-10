'==============================================================================
'Project Number: 205713
'User Story: CPQEncoreRetirement_HPFS_01_QuoteMCC
'Description:	1. Sales op is able to create an HPEFS quote in NGQ.
'				2. MCC code 72 can be used to adjust quote total to HPFS quote price.
'				3. NGQ is able to generate the quote output for this HPEFS quote.
'Tags: HPEFS Quote, MCC, Output
'Last Modified: 5/10/2017 by yu.li9@hpe.com
'==============================================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'DataImport
DataTable.Import "..\..\data\NGQ_HPFS_01_data.xlsx"
DataTable_ImportDataSheet "..\..\data\NGQ_empty_quote_data.xlsx", "Sheet1"

'Intialize a new user
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Sheet1"), DataTable.Value("pass", "Sheet1"), "<Encrypted DigitalBadge>")

'Fetch and store test data
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim strQuotename : strQuoteName = DataTable("QuoteName")
Dim deliverySpeed : deliverySpeed = DataTable.Value("DeliverySpeed", "Global")
Dim strDeliveryTerms : strDeliveryTerms = DataTable("DeliveryTerms")
Dim strExternalComments : strExternalComments = DataTable("ExternalComments")
Dim strPdfOutputType : strPdfOutputType = DataTable("pdfOutputType")
Dim strProductFilePath : strProductFilePath = getProductFilePath(DataTable.Value("ProductFileName", "Global"))

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74467", "CPQEncoreRetirement_HPFS_01_QuoteMCC"

' Open the NGQ
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Check the HPEFS checkbox
CheckHPEFS

'Enter an opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

'Enter quote name and save it
Quote_EditQuoteName strQuotename
Quote_save

'Upload products
uploadProduct 
setUploadProductPath strProductFilePath
UploadProducts_ProceedWithImport
UploadProducts_VerifyAddToQuoteTabDisplayed
UploadProducts_AddValidProducts
UploadProducts_ProductsAddedMsg

'Fill neccesary data
Quote_CustomerDataTab
CustomerData_ShipToTab
CustomerDataShipTo_SelectSameAsSoldToAddress
CustomerData_ShipToTab
Quote_ShippingDataTab
ShippingData_SetDeliverySpeed deliverySpeed
ShippingData_SetTermsOfDelivery strDeliveryTerms
Quote_AdditionalDataTab
AdditionalData_SetReceiptDateNow

'Click refresh pricing
Dim strQuoteSelector : strQuoteSelector = "Refresh Pricing"
QuoteServices_SelectOption strQuoteSelector

'Apply MCC
Quote_PricingTermsTab
Dim MCCType : MCCType = DataTable.Value("MCCType", "Global")
Dim MCCOffApp : MCCOffApp = DataTable.Value("MCCOffApp", "Global")
Dim MCCDiscType : MCCDiscType = DataTable.Value("MCCDiscType", "Global")
Dim MCCValueType : MCCValueType = DataTable.Value("MCCValueType", "Global")
Dim MCCPercentage : MCCPercentage = DataTable.Value("MCCPercentage", "Global")
Dim MCCmsg : MCCmsg = DataTable.Value("MCC_msg", "Global")
Dim MCCAmount : MCCAmount = DataTable.Value("MCCDiscAmt", "Global")
RequestOPDisc MCCType, MCCOffApp, MCCDiscType, MCCValueType, MCCPercentage, MCCAmount, MCCmsg
applyEmpowerment "MCC"

Dim intGrandTotal : intGrandTotal = get_grand_total
Quote_save

'Generate PDF and verify the data
Quote_QuoteOutputTab
QuoteOutput_ExternalComments strExternalComments
OutputQuote_SetOutputType strPdfOutputType
Dim strQuoteNumber : strQuoteNumber = Quote_get_quoteNumber
OutputQuote_ClickPreview
OutputQuote_SaveQuotePdf strQuoteNumber

PdfVerification DataTable("OutputFilePath"), DataTable("PdfCheckSoldTo"), DataTable("PdfCheckShipTo"), DataTable("PdfCheckSalesContact"), _
                DataTable("PdfCheckLineItems"), DataTable("PdfCheckGrandTotal"), DataTable("PdfCheckExtComment"), DataTable("PdfCheckHeader")

'Exit test
Navbar_Logout
Close_Browser
FinalizeTest
