'==========================================================================
'Project Number: 205713
'User Story:US9414_01 
'Description:Capture comments when create new a Quote and version the quote
'==========================================================================
Option Explicit
Dim al : Set al = NewActionLifetime
Dim strExternalCommemnt : strExternalCommemnt = "External comment for testing"
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<encrypted digitalbadge>")
Dim strQuote,strProductNumber,strQutoteNum1,strVersion1,strQuoteStatus1,strEndDate1
Dim strVersion2,strInternalComments1,strInternalComments2, strGenOutputType,strNewVersionSource, strNewversionReason
Dim strExternalComments,strNewVersionComments,strVersion3,strQuoteName,strNewVersionIntComments
Dim strOutputFilePath, strSavePath

'Import Data
DataTable.Import "..\..\data\US9414_01.xlsx"
strProductNumber = DataTable("Product_Number", dtGlobalSheet)
strQutoteNum1 = DataTable("QuoteNumber_1",dtGlobalSheet)
strVersion1 =  DataTable("Version_1",dtGlobalSheet)
strQuoteStatus1 =  DataTable("QuoteStatus_1",dtGlobalSheet)
strEndDate1 = DataTable("EndDate_1",dtGlobalSheet)
strVersion2 = DataTable("Version_2", dtGlobalSheet)
strInternalComments1 = DataTable("InternalComments_1",dtGlobalSheet)
strInternalComments2 = DataTable("InternalComments_2",dtGlobalSheet)
strExternalComments = DataTable("External_Comments", dtGlobalSheet)
strGenOutputType = DataTable("General_OutputType", dtGlobalSheet)
strNewVersionSource = DataTable("NewVersion_Source", dtGlobalSheet)
strNewversionReason = DataTable("NewVersion_Reason", dtGlobalSheet)
strNewVersionComments = DataTable("NewVersionComment", dtGlobalSheet)
strVersion3 = DataTable("Version_3", dtGlobalSheet)
strNewVersionIntComments = DataTable("NewVersionIntComments", dtGlobalSheet)

'Creating output sheet
Dim strOutputSheet
strOutputSheet = "US9414_01_Output"
DataTable.AddSheet strOutputSheet
DataTable.GetSheet(strOutputSheet).AddParameter "QuoteNumber", ""

InitializeTest ""
'Opens the browser and opens ngq website
OpenNgq objUser

'Create New Quote
Navbar_CreateNewQuote

'Validates Quote Number,Quote Version,Quote Name,Quote Status,Quote Start Date and Quote End Date
NewQuote_ValideEmptyQuote strQutoteNum1, strVersion1, strQuoteStatus1, strEndDate1

'Makes sure the opportunity tab is opened
OpportunityandQuoteInfoTabExistence

'Display Quote Output tab
QuoteOutput

'Validates Internal Comments and he Max number of chars that can be stored for internal comments is limited to 4000 characters
QuoteOutput_Internal_Comments_Validation 

'Set the comments in the internal comment box
Set_InternalComments strInternalComments1

'Save the comments in the internal comment box and the page is refreshed
Quote_Save

'Validates quote number is generated and saves the Quote number
strQuote = Quote_get_quoteNumber
'Set data in output sheet
DataTable("QuoteNumber", strOutputSheet) = strQuote
'Validates The newly entered comments is saved in read only mode with the created user email ID, timestamp in the first text box of the "Internal Comments".
QuoteOutput_ValidateInternalComments

'Validates the second comment box is empty
QuoteOutput_Internal_Comments_Empty

'Set the comments in the internal comment box
Set_InternalComments strInternalComments2
'Clicks the "Save" button on the top right of the page and the page is refreshed
Quote_Save

'Validate tehcomments is displayed with the created user email ID, timestamp and comment information in read only mode and The existing comments is displayed 
QuoteOutput_ValidateInternalComments

'Clicks the pencil icon of the "External Comments" text box and set the comment in the external comments box
'==============old url=========================
'QuoteOutput_ExternalComments strExternalComments
'=============New URL============================
QuoteOutput_SetExternalComments strExternalComments

'Clicks the "Save" button on the top right of the page
Quote_Save

'Checks the checkbox of "Include the comment in quote Output" under the "External Comments" text box.
QuoteOutput_ExternalCommentCheckBox

'Selects the "Add Product or Option" option
'Click_AddProdAndOption deleted as duplicate - JH

'Set the product number and press enter key in the keyboard
'SetProductNumber strProductNumber deleted as duplicate - JH

'Replaced duplicated functions to add product -JH
LineItemDetails_AddProductByNumber strProductNumber, "1"

'Selects the "Output Type" as "Extended Net Price by item with estimated delivery" under General information section
Select_Quote_GeneraL_OutputType strGenOutputType

'Clicks the "Preview" button
Click_Preview_Quote_OutputType 

'Saves the pdf in a specific  path
OutputQuote_SaveQuotePdf strQuote

'Validates the pdf contains the external comments
validateCommentInPdf(strExternalComments)

'Closing pdf file
pdfClose

'Clicks the "Advanced Search" tab on the top of this page 
AdvancedSearchClick

'Sets the quotenumber saved before in the quote number edit box
SetQuoteNumber_AdvancedSearch strQuote

'Clicks the search button
ClickSearch_advancedSearch

'Clicks the hyperlink quote number under result section
AdvancedSearch_Result_OpenQuoteNumber strQuote

'Clicks the Quote Output tab
QuoteOutput

'Makes sure the saved internal comments remain in the comment box
validate_SavedInternalComments strInternalComments1
validate_SavedInternalComments strInternalComments2  

'Makes sure the saved external comment remain in the comment box
validate_SavedExternalComments strExternalComments

'Clicks the "New Version" button on the top right of the page
Click_QuoteNewVersionButton

'Selects New version source from the dropdown list
Choose_NewVersionSource strNewVersionSource

''Selects New version reason from the dropdown list
Choose_NewVersionReason strNewversionReason

'Set the comments in the comment section
SetComment_NewVersionWindow strNewVersionComments

'Presses OK button 
NewVersion_OkButton

strQuoteName = "New Version of " & strQuote &"-01"
'Validates Quote number, version number, Quote name, Quote Status, Start Date and End Date
Validated_Quote strQuote, strVersion3,strQuoteName, strQuoteStatus1, strEndDate1 

'Clicks Save button on the top right of the page
Quote_Save

'Validates the version number is shown as "02"
validate_VersionNumber strVersion2

'Clicks the Quote Output tab
QuoteOutput

'Validates the saved internal comments in original version remain in the new version of this quote
validate_SavedInternalComments strInternalComments1
validate_SavedInternalComments strInternalComments2

'Makes sure the second internal comment box is empty
QuoteOutput_Internal_Comments_Empty

'Validates the saved external comments in original version remain in the new version of this quote
validate_SavedExternalComments strExternalComments

'Sets comment in the internal comment box
Set_InternalComments strNewVersionIntComments

'Clicks save button on the top right of the page
Quote_Save

'Validates the existing internal comments is displayed with the created user email ID, timestamp and comment information in read only mode
validate_SavedInternalComments strInternalComments1
validate_SavedInternalComments strInternalComments2

'Makes sure that The "External Comments" text box is displayed with the external comments
validate_SavedExternalComments strExternalComments

'Export output data to excel sheet in test script dir

strOutputFilePath = Environment.Value("TestDir") & "\Output\" & strOutputSheet & ".xls"
DataTable.ExportSheet strOutputFilePath, strOutputSheet

'Log off NGQ 
Navbar_Logout

'Closes the browser
Close_Browser

FinalizeTest

