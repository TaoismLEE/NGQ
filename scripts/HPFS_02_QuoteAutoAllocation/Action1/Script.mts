'==============================================================================
'Project Number: 205713
'User Story: HPEFS_02
'Description: Validates that auto allocation can be used to adjust quote total to HPEFS quote price.
'Tags:Quote, Allocation, HPEFS
'Last Modified: yu.li9@hpe.com
'==============================================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Initialize test
InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'DataImport
DataTable.Import "..\..\data\HPFS_02.xlsx"

' Set opportunity id and 3rd party product number
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

'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74468", "CPQEncoreRetirement_HPFS_02_QuoteAutoAllocation"

'Open the NGQ
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Check HPEFS checkbox
CheckHPEFS

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

'Enter quote name and save it
Quote_EditQuoteName quoteName
quote_editCutomerSpecQouteID quoteName
click_save_button()

'Upload the product
uploadProduct 
importProductExcelSheet dirPath + "\data\depends\" + excelName

'Input nessary data
PreValidate_FixDataCheckErrors

'Refresh Pricing
click_refresh_pricing()

'Add auto allocation target requested net price
Quote_SetTargReqPrice2 targPrice
verifyGrandTotal

'Generate PDF and validate data
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

Navbar_Logout
Close_Browser
FinalizeTest




