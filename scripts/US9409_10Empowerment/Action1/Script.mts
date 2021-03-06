﻿'================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9409_10: Deal Generated After Quote Complete
'Author: Joshua Hunter
'Description: This test deals with testing Operator Discounts / Empowerments
'Tags: Quote, Empowerments
'================================================

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"
'InitializeTest ""

'DataImport
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<encrypted digitalbadge>")


'set var from data sheets
Dim emptyQuoteNumber : emptyQuoteNumber = DataTable.Value("quoteNumber", "Global")
Dim emptyQuoteVersion : emptyQuoteVersion = DataTable.Value("quoteVersion", "Global")
Dim emptyQuoteStatus : emptyQuoteStatus = DataTable.Value("quoteStatus", "Global")
Dim emptyQuoteEndDate : emptyQuoteEndDate = DataTable.Value("quoteEndDate", "Global")
Dim emptyQuoteSelectedTab : emptyQuoteSelectedTab = DataTable.Value("selectedTab", "Global")

DataTable.Import "..\..\data\NGQ_US9409_10_data.xlsx"

Dim opportunityID : opportunityID = DataTable.Value("Opportunity_ID", "Global")
Dim quoteName : quoteName = DataTable.Value("quoteName", "Global")

' For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74322", "CPQ_Encore Retirement_US9409_Deal Generated After Quote Complete_10"

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote

NewQuote_ValidateEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

' REQUIRED FOR PRE-VALIDATION TO PASS
Quote_CustomerDatatab

CustomerData_ShipToTab

CustomerDataShipTo_SelectSameAsSoldToAddress

Quote_ShippingDataTab

ShippingData_SetDeliverySpeed DataTable.Value("DeliverySpeed", "Global")

ShippingData_SetTermsOfDelivery DataTable.Value("DeliveryTerms", "Global")

Quote_AdditionalDataTab

AdditionalData_SetReceiptDateNow
' END REQUIREMENTS FOR PRE-VALIDATION TO PASS

'Dim quoteID : quoteID = Quote_get_quoteNumber
Quote_AddProductOrOption 
AddProductsFromTable

reset_DataTable

Quote_refreshPricing

Quote_save

Quote_PricingTermsTab

Dim MCCType : MCCType = DataTable.Value("MCCType", "Global")
Dim MCCOffApp : MCCOffApp = DataTable.Value("MCCOffApp", "Global")
Dim MCCDiscType : MCCDiscType = DataTable.Value("MCCDiscType", "Global")
Dim MCCValueType : MCCValueType = DataTable.Value("MCCValueType", "Global")
Dim MCCPercentage : MCCPercentage = DataTable.Value("MCCPercentage", "Global")
Dim MCCmsg : MCCmsg = DataTable.Value("MCC_msg", "Global")
Dim MCCAmount : MCCAmount = DataTable.Value("MCCDiscAmt", "Global")

RequestOPDisc MCCType, MCCOffApp, MCCDiscType, MCCValueType, MCCPercentage, MCCAmount, MCCmsg

'after this part it brakes in 12.52
applyEmpowerment "MCC"

Quote_refreshPricing

Quote_save

'Pre-validate config
SelectPreValidate
PreValidate_DataCheckNoErrors
PreValidate_ClicNoErrors
PreValidate_PriceNoErrors
PreValidate_BundleNoErrors

'Check the Deal is generated
CheckDealGenerateInfoDisplay
Dim strDealId
strDealId = FetchDealId
PreValidate_CloseValidationPage
DisplayDealId
CheckDealIdDisplayed strDealId

Navbar_Logout
Close_Browser
FinalizeTest

