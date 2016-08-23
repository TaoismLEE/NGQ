'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest

'Fetch data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strOpportunityId : strOpportunityId = "OPE-0005373487"

'Open browser.
OpenNgq objUser

'Navigate to create new quote
Dim strQuoteNumberID
Dim strQuoteVersion
Dim strQuoteStatus
Dim strQuoteEndDate
Dim strQuoteTabSelected : strQuoteTabSelected = "Opportunity and Quote Info"

Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote strQuoteNumberID, strQuoteVersion, strQuoteStatus, strQuoteEndDate
Quote_currentlySelectedTab(strQuoteTabSelected)
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

'TODO Jesus: Add validation for proper opportunity ID



'NOTE: automation API calls only here. No raw UFT calls!

FinalizeTest

