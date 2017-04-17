'================================================
'Project Number: 205713
'User Story : CPQ_Encore Retirement_US9414_02: Capture Comments When Clone a Quote with Comments
'Description: This case is to validate:
'			1. After cloning another user's quote, the existing comments of the quote don't remain in the cloned quote.
'			2. Sales Op is able to add new internal comments, NGQ is able to captures all the new internal comments. 
'			3. NGQ is able to capture when the internal comments have been created and by whom.
'Tags: Quote, Comment, Comments, Clone
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

Dim strQuote
Dim strInternalComments
'Import Data
DataTable.Import "..\..\data\US9414_02\US9414_02.xlsx"
strQuote = DataTable("QuoteNumber", dtGlobalSheet)
strInternalComments = DataTable("InternalComments", dtGlobalSheet)

' For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74236", "CPQ_Encore Retirement_US9414_Capture Comments When Clone a Quote with Comments_02"

InitializeTest "Action1"
'Opens browser and ngq website
OpenNgq objUser

'Set Quote number under Quick Search
QuickSearch strQuote

'Clicks Search Button under Quick search
QuickSearch_Search

'Clicks the quote number that is displayed Under Result
'AdvancedSearch_Result_OpenQuoteNumber strQuote
ClickQuoteNumberResult(2)

'Clicks Output tab
Quote_OutputTab

'Validates two Internal Comment box are visible
Validate_TwoInternalCommentsBox

'Validates the internal comment display box is readonly mode
Validate_ReadOnlyInternal_DisplayCommentsBox ' Need to validate

'validates Internal comment box is empty and read only mode
Internal_Commentbox_Empty

'Clicks Clone Button
Click_Clone

'Clicks the save button on the top right of the page
Quote_save

'Clicks on Quote output tab
Quote_OutputTab

'Validates the internal comment box is empty
QuoteOutput_Internal_Comments_Empty 

'Set the comment in the internal box
Set_InternalComments strInternalComments

'Clicks the save button on the top right of the page
Quote_save
'Validates saved comments remain in the interanal comments
validate_SavedInternalComments strInternalComments

'Validates ther displayed comments are in correct format
QuoteOutput_ValidateInternalComments

'Makes sure the second internal comment box is empty
QuoteOutput_Internal_Comments_Empty

'Log off NGQ
Navbar_Logout

'Close the browser
Close_Browser

FinalizeTest


