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


'InitializeTest "CH"
InitializeTest ""
'Hard-coded data.
Dim objUser : Set objUser = NewRealUser(Parameter("username"), Parameter("password"))
Dim emptyQuoteNumber : emptyQuoteNumber = "New Quote"
Dim emptyQuoteVersion : emptyQuoteVersion = "01"
Dim emptyQuoteStatus : emptyQuoteStatus = "Quote/Configuration Created"
Dim emptyQuoteEndDate : emptyQuoteEndDate = "Need Pricing Call"
Dim emptyQuoteSelectedTab : emptyQuoteSelectedTab = "Opportunity and Quote Info"
Dim opportunityID : opportunityID = "OPE-0002935252"
Dim quoteName : quoteName = "Test Quote"

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote

NewQuote_ValideEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

Dim quoteID : quoteID = Quote_get_quoteNumber

LineItemDetails_AddProductByNumber "AJ762B", "1"
'LineItemDetails_AddProductByNumber "AF556A", "1"
'LineItemDetails_AddProductByNumber "BW946A", "1"
'LineItemDetails_AddProductByNumber "G1S72A", "1"

Quote_refreshPricing
Quote_PricingTermsTab
verify_price_quality_band

lineItemDetails_addColumn "Pricing_Source"
lineItemDetails_addColumn "Source_ID"

lineItemDetails_verifyPricingSource
'Quote_save
'quoteID = Quote_get_quoteNumber

'applyEmpowerment "Preferred"
'applyEmpowerment "Personal"
'applyEmpowerment "Manager"


FinalizeTest
