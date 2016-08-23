'================================================
'Summary: A demo for NGQ
'
'Description:
'Demo is as demo does.
'
'Preconditions:
'Recommended: Use programing descriptive not objects repository
'Author: Ana Karina Orduna
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strOpportunityId : strOpportunityId = "OPE-0002916168"
Dim strDeliverySpeed : strDeliverySpeed = "Express"
Dim strDeliveryTerms : strDeliveryTerms = "Delivery Duty Paid"
Dim strProductNumber : strProductNumber = "BW904A"
Dim intProductQuantity : intProductQuantity = 1
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = "Refresh Pricing"

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

Quote_CustomerDataTab
CustomerData_ShipToTab
CustomerDataShipTo_SelectSameAsSoldToAddress
CustomerData_BillToTab
CustomerDataBillTo_SelectSameAsSoldToAddress

Quote_ShippingDataTab
ShippingData_SetDeliverySpeed strDeliverySpeed
ShippingData_SetTermsOfDelivery strDeliveryTerms

Quote_AdditionalDataTab
AdditionalData_SetReceiptDateNow

Quote_AddProductOrOption
LineItemDetails_SetProductNumberByIndex 0, strProductNumber
LineItemDetails_SetProductQuantityByIndex 0, intProductQuantity
strQuotaSelection_Selector = "Refresh Pricing"
QuoteServices_SelectOption strQuotaSelection_Selector
strQuotaSelection_Selector = "Save"


FinalizeTest

