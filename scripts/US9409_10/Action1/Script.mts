'================================================
'Project Number: 205713
'User Story: US9409_10
'Author: Joshua Hunter
'Description: This test deals with testing Operator Discounts / Empowerments
'Tags: Quote, Empowerments
'================================================

Option Explicit
Dim al : Set al = NewActionLifetime

'InitializeTest "CH"
InitializeTest ""

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

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote

NewQuote_ValidateEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

'Dim quoteID : quoteID = Quote_get_quoteNumber
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

applyEmpowerment "MCC"


' REQUIRED FOR PRE-VALIDATION TO PASS
Quote_CustomerDatatab

Quote_ShiptoTab

CustomerDataShipTo_SelectSameAsSoldToAddress

Quote_ShippingDataTab

ShippingData_SetDeliverySpeed DataTable.Value("DeliverySpeed", "Global")

ShippingData_SetTermsOfDelivery DataTable.Value("DeliveryTerms", "Global")

Quote_AdditionalDataTab

AdditionalData_SetReceiptDateNow
' END REQUIREMENTS FOR PRE-VALIDATION TO PASS

Quote_refreshPricing

Quote_save

select_preValidate_link

PreValidateQuote

PreValidateQuote_success

Navbar_Logout

FinalizeTest

browser("NGQ").Close

