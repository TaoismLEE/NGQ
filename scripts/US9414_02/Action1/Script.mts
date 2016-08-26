'================================================
'Project Number: 205713
'User Story : US9414_02
'Description: Capture comments when clone a quote with comments
'Author: Pramesh Bhandari
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strQuote
Dim strInternalComments
'Import Data
DataTable.Import "..\..\data\US9414_02.xlsx"
strQuote = DataTable("QuoteNumber", dtGlobalSheet)
strInternalComments = DataTable("InternalComments", dtGlobalSheet)

InitializeTest
'Opens browser and ngq website
OpenNgq objUser

'Set Quote number under Quick Search
QuickSearch strQuote

'Clicks Search Button under Quick search
QuickSearch_Search

'Clicks the quote number that is displayed Under Result
'==============================================
AdvancedSearch_Result_OpenQuoteNumber strQuote
'==============================================
'Clicks Output tab
QuoteOutput

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
QuoteOutput

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

