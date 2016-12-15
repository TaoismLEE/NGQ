﻿'================================================
'Project Number: 205713
'User Story: HPFS_01
'Description:	1. Sales op is able to create an HPFS quote in NGQ.
'				2. MCC code 60B can be used to adjust quote total to HPFS quote price.
'				3. NGQ is able to generate the budgetary quote output for this HPFS quote.
'Tags: Quote, MCC, Output
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'DataImport
DataTable.Import "..\..\data\NGQ_HPFS_01_data.xlsx"
DataTable_ImportDataSheet "..\..\data\NGQ_empty_quote_data.xlsx", "Sheet1"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Sheet1"), DataTable.Value("pass", "Sheet1"), "<Encrypted DigitalBadge>")
'Dim objUser : Set objUser = NewRealUser("yu.li9@hpe.com", "", "<Encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim strQuotename : strQuoteName = DataTable("QuoteName")
Dim deliverySpeed : deliverySpeed = DataTable.Value("DeliverySpeed", "Global")
Dim strDeliveryTerms : strDeliveryTerms = DataTable("DeliveryTerms")
Dim strExternalComments : strExternalComments = DataTable("ExternalComments")
Dim strPdfOutputType : strPdfOutputType = DataTable("pdfOutputType")
Dim strProductFilePath : strProductFilePath = getProductFilePath(DataTable.Value("ProductFileName", "Global"))

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Enter quote name and save it
Quote_EditQuoteName strQuotename
Quote_save

uploadProduct 

setUploadProductPath strProductFilePath
UploadProducts_VerifyProducts strProductFilePath

UploadProducts_ProceedWithImport

UploadProducts_VerifyAddToQuoteTabDisplayed

UploadProducts_AddValidProducts

UploadProducts_ProductsAddedMsg

Quote_CustomerDataTab
Quote_currentlySelectedTab "Customer Data"

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
ValidatePriceRefreshed

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

Quote_QuoteOutputTab

QuoteOutput_ExternalComments strExternalComments

OutputQuote_SetOutputType strPdfOutputType

Dim strQuoteNumber : strQuoteNumber = Quote_get_quoteNumber
OutputQuote_ClickPreview
OutputQuote_SaveQuotePdf strQuoteNumber
'print DataTable("OutputFilePath")
PdfVerification DataTable("OutputFilePath"), DataTable("PdfCheckSoldTo"), DataTable("PdfCheckShipTo"), DataTable("PdfCheckSalesContact"), _
                DataTable("PdfCheckLineItems"), DataTable("PdfCheckGrandTotal"), DataTable("PdfCheckExtComment"), DataTable("PdfCheckHeader")

Navbar_Logout
Browser("NGQ").Close()

FinalizeTest
