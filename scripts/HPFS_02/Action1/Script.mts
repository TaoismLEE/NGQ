'Project Number: 205713
'User Story: HPFS_02
'Description: Validates that auto allocation can be used to adjust quote total to HPFS quote price.
'Tags:

Option Explicit
Dim al : Set al = NewActionLifetime

'Initialize test
InitializeTest "IE"

'DataImport
DataTable.Import "..\..\data\HPFS_02.xlsx"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "a")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim obsoleteNumber : obsoleteNumber = DataTable.Value("ObsoleteNumber", "Global")
Dim validNumber : validNumber = DataTable.Value("ValidNumber", "Global")
Dim deliverySpeed : deliverySpeed = DataTable.Value("DeliverySpeed", "Global")
Dim targPrice : targPrice = DataTable.Value("TargetPrice", "Global")
Dim comment : comment = DataTable.Value("ExternalComment", "Global")
Dim outputType : outputType = DataTable.Value("OutputType", "Global")
Dim quoteName : quoteName = DataTable.Value("QuoteName", "Global")
Dim dirPath : dirPath = Environment.Value("TestDir") + "\..\.."
Dim excelName : excelName = DataTable.Value("ExcelSheet", "Global")
Dim quotationName : quotationName = DataTable.Value("quotationName", "Global")
Dim numberOfRows : numberOfRows = DataTable.Value("NumberOfValidProducts","Global")
Dim pdfPath

'NOTE: automation API calls only here. No raw UFT calls!

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Enter quote name and save it
Quote_EditQuoteName quoteName
quote_editCutomerSpecQouteID quoteName

click_save_button()

'Upload the product
uploadProduct 

importProductExcelSheet dirPath + "\data\depends\" + excelName

Quote_ShiptoTab

CustomerDataShipTo_SelectSameAsSoldToAddress

' Click shipping data tab
Quote_ShippingDataTab

' Set speed
ShippingData_SetDeliverySpeed deliverySpeed

' Set Delivery terms
ShippingData_SetTermsOfDelivery DataTable.Value("DeliveryTerms", "Global")

' Set receipt date
Quote_AdditionalDataTab

AdditionalData_SetReceiptDateNow

'Refresh Pricing
click_refresh_pricing()

' Add auto allocation target requested net price
Quote_SetTargReqPrice targPrice

verifyGrandTotal

Quote_OutputTab

EditExternalComment comment

Quote_CaptureQuoteNumber

pdfPath = dirPath + "\data\depends\" + DataTable.Value("QuoteNumber_Output", "Global") + ".pdf"
Quote_SelectOutputType outputType, pdfPath

Dim pdfObj : Set pdfObj = NewPdfParser(pdfPath)

verifyHeaderInPDF quotationName, pdfObj

verifyCommentInPDF comment, pdfObj

verifyGrandTotalInPDF DataTable.Value("GrandTotal", "Global"), pdfObj

verifyProductInPDF pdfObj, numberOfRows

FinalizeTest




