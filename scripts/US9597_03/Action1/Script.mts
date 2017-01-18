'================================================
'Project Number:205713
'User Story: US9597_03
'Description: This case is to validate:
'				1. Sales op is able to create,edit,delete custom view and set custom view as default.
'				2. NGQ is able to  display the default custom view configured by Sales op.
'Tags: Create, Edit, Delete, Custom, View
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\US9547_03.xlsx"
Dim strChooseViewName : strChooseViewName = DataTable("strChooseViewName",1)
Dim strQuoteName : strQuoteName = DataTable("strQuoteName",1)
Dim strOportunityId : strOportunityId = DataTable("strOportunityId",1)

Dim strColumnLabel
	strColumnLabel = ExellToArray()
	
'array with Values in the assign column
Dim LabelsAssignedColumn

InitializeTest "Action1"
'Open browser and go to NGQ
OpenNgq(objUser)

'go to My Preferences in Admin Tools navbar
ClickAdminTools_MyPreferences()

'set the Choose view field
EditChooseView(strChooseViewName)

'Select a item from Available column and send to assigned column
AvailableColumn(strColumnLabel(5))
AvailableColumn(strColumnLabel(9))

'Select a item from Assigned Column and send to Available column
AssignedColumn(strColumnLabel(5))

'Move a item in the assigned column up
MoveUpAssignedLabel(strColumnLabel(9))

'Move a item in the assigned column down
MoveDownAssigendLAbel(strColumnLabel(9))

'Click the "Set as Default" checkbox
CheckSetAsDefault()

'Save all the items in the array "LabelsAssignedColumn" to compare later 
LabelsAssignedColumn = NoteDownAssignedColumn()

'save the new choose view
ClickSaveBtnMyPrecerences()

'go to new Quote navbar
Navbar_CreateNewQuote()

'Validate information in New Quote 
NewQuote_ValidateEmptyQuote null,null,null,null

'set opportunity ID
OpportunityAndQuoteInfo_SetOpportunityId(strOportunityId)

'Click Import button
OpportunityAndQuoteInfo_Import()

'Save Import
Quote_save()

'Set Quote Name
Quote_EditQuoteName(strQuoteName)

'Save Import
Quote_save()

'Scroll down
pageDownNewQuotePage()

'Validate Choose view match with the new one
ValidateChooseView(strChooseViewName)

'Validate assigned column in new quote
ValidateAssignedList_NewQuote LabelsAssignedColumn

'go to My Preferences in Admin Tools navbar
ClickAdminTools_MyPreferences()

'Set Choose Viewd
EditChooseView(strChooseViewName)

'UnClick the check box
CheckSetAsDefault()

'Save
ClickSaveBtnMyPrecerences()

'go to new Quote navbar
Navbar_CreateNewQuote()

'Scroll down
pageDownNewQuotePage()

'Validate Choose view match with the default
ValidateChooseView("DEFAULT_VIEW")

'go to My Preferences in Admin Tools navbar
ClickAdminTools_MyPreferences()

'Set Choose View
EditChooseView(strChooseViewName)

'Click to Delete Choose view
ClickDeleteBtnChooseView()

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest
