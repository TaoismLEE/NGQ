'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest
'Initialize Data


Dim optId : optId = "OPE-0005373487"
Dim productNumber : productNumber = "AF556A"
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim quoteNumber

'Open Browser
OpenNgq objUser


'Create New Quote
Navbar_CreateNewQuote

'Validates Quote Number,Quote Version,Quote Name,Quote Status,Quote Start Date and Quote End Date
NewQuote_ValideEmptyQuote "New Quote", "01", "Quote/Configuration Created", "Quote/Configuration Created"
OpportunityandQuoteInfoTabExistence

'Display Quote Output Page
QuoteOutput

'validates Internal Comments and he Max number of chars that can be stored for internal comments is limited to 4000 characters
QuoteOutput_Internal_Comments_Validation "","( 4000 characters remaining )"

'Set the comments
Set_InternalComments "first comment for the quote"

'Save the comments
Quote_Save

'Validates the quote number generated
Quote_get_quoteNumber
quoteNumber = Quote_get_quoteNumber
'Validate Comments
QuoteOutput_ValidateInternalComments

'Validate the second comment box is empty
QuoteOutput_Internal_Comments_Empty ""

Set_InternalComments "second comment is the quote"

Quote_Save

QuoteOutput_ValidateInternalComments

QuoteOutput_ExternalComments "External comment for testing"

Quote_Save

QuoteOutput_ExternalCommentCheckBox

'AddBtnHover
'ClickAddproductOption 'Step 28 -33 Check with josh 


'34 Started
AdvancedSearchClick
SetQuoteNumber_AdvancedSearch quoteNumber
ClickSearch_advancedSearch

ClickResult_QuoteNumber

QuoteOutput
validate_SavedInternalComments


