'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest
Dim optId : optId = "OPE-0002935249"
Dim quoteNumber1 : quoteNumber1 = "NI00156297" 
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")

'Opens browser and ngq website
OpenNgq objUser

'Create New Quote
Navbar_CreateNewQuote

' Validates Quote number, version number, Quote status, start date and end date
NewQuote_ValideEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Makes sure Opportunity and quote tab is displayed as default
OpportunityandQuoteInfoTabExistence


OpportunityAndQuoteInfo_ImportOpportunityId optId

'Click the pencil icon next to "Quote Name" and enter your quote name
Quote_EditQuoteName quoteName1

'Clicks the "Save" button on the right of the page
Quote_save

Quote_QuoteStatus "Quote/Configuration Created"

pageDownNewQuotePage
'Mouse over the "+Add" button.
Click_Add



FinalizeTest

