'================================================
'Summary: A demo for NGQ
'
'Author: Guillermo Soria & Ana Karina Orduña
'Demo is as demo does.
'Preconditions: Have strQuoteNumber as owner
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest

'Hard-coded data
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strQuoteNumber : strQuoteNumber = "NI00147432"
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = "Custom Group"
'Dim strQuotaSelection_Selector : strQuotaSelection_Selector = "Claim"
Dim strGroupLabel : strGroupLabel ="Automated Test Label"
Dim strGroupSummary : strGroupSummary ="Automated Test Summary"
Dim strGroupLabelEdited : strGroupLabelEdited ="Automated Test Label - Edited"
Dim strGroupSummaryEdited : strGroupSummaryEdited ="Automated Test Summary - Edited"

'Open browser.
OpenNgq objUser
QuickSearch strQuoteNumber
QuickSearch_Search
SelectResult_Search strQuoteNumber
'QuoteServices_Claim

QuoteServices_SelectOption strQuotaSelection_Selector
QuoteServices_AddCustomGroup strGroupLabel,strgroupSummary
QuoteServices_EditCustomGroup strGroupLabelEdited,strgroupSummaryEdited
QuoteServices_RemoveCustomGroup

FinalizeTest
