'=========================================================================================================
'Project Number: 205713
'User Story: CPQ_Encore Retirement_US9414_Capture Comments When Claim a Quote_04
'DescriptioN: This user case claims a quote owned by User B and verifies the claim process by searching the 
'             same quote in advanced search. The comments presence will be validated manually by clicking the link in the email.
'Tags:
'=========================================================================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

Dim strQuote
Dim strEmail : strEmail = DataTable.Value("user", "Global")

'Import Data
'DataTable.Import "..\..\data\US9414_04\US9414_04.xlsx"
'strQuote = DataTable("QuoteNumber", dtGlobalSheet)	
'strEmail = DataTable("Email", dtGlobalSheet)

'Doesnt work - Ask Richard how to do this
'NGQBrowserReInit

InitializeTest "US9414_04"

'For Jenkins reporting
dumpJenkinsOutput "US9414_04", "74238", "CPQ_Encore Retirement_US9414_Capture Comments When Claim a Quote_04"

'Opens browser and ngq website
OpenNgq objUser

'Navigate my dashboard tab
Click_MyDashboard

'Validates if the quote tab is active/selected
ValidateQuoteTab

'Clicks the "My Group Quote" tab next to the "My Quote" tab
Click_MyGroupQuote

'Clicks the "Count" number associated to "Quote Status"- Quote/Configuration Created
Click_QuoteConfiguration_Count

'Locates and clicks the "Auto Filter" button in the "Result'' section 
ClickAutoFilter

'Fetch the first quote in my group quotes
strQuote = GetFirstQuoteNumberofMyGroupQuote(2)

'Enter and submit th equote number in My Dashboard
SetAutoFilterQuoteNumber strQuote

'Selects the Quotenumber row
Check_RadioButton strQuote

'click on Claim button on the top right of result section
Click_Claim

'Click ok button to confirm
Click_Ok_Claim

'Click the "OK" button to close the window
Quote_Claim_Success_Ok

'Click on Advanced search
AdvancedSearchClick

'Set the Quote number in Quote number field
SetQuoteNumber_AdvancedSearch strQuote

'Click the search button
ClickSearch_advancedSearch

'TODO FIX THIS JOSH
'Verify the email to check the quote has been claimed
VerifyEmailQuote strEmail, strQuote

'Logout the NGQ
Navbar_Logout

'close the browser
Close_Browser

FinalizeTest




