'================================================
'Project Number:205713
'User Story: CPQ_Encore Retirement_US9597_03: Sales Op Create, Edit and Delete Custom View
'Description: This case is to validate:
'				1. Sales op is able to create,edit,delete custom view and set custom view as default.
'				2. NGQ is able to  display the default custom view configured by Sales op.
'Tags: Create, Edit, Delete, CustomView
'Last Modified: 4/20/2017 by yu.li9@hpe.com
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\US9547_03.xlsx"
Dim strChooseViewName : strChooseViewName = DataTable("strChooseViewName",1)
Dim strQuoteName : strQuoteName = DataTable("strQuoteName",1)
Dim strOportunityId : strOportunityId = DataTable("strOportunityId",1)

Dim arrColumnLabel
	arrColumnLabel = ExellToArray()
	
'array with Values in the assign column
Dim arrLabelsAssignedColumn

' For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74269", "CPQ_Encore Retirement_US9597_03: Sales Op Create, Edit and Delete Custom View"

'Open browser and go to NGQ
OpenNgq(objUser)

'go to My Preferences in Admin Tools navbar
ClickMyPreferenceUnderAdminTools

'set the Choose view field
EditChooseView(strChooseViewName)

'Select a item from Available column and send to assigned column
AvailableColumn(arrColumnLabel(5))
AvailableColumn(arrColumnLabel(9))

'Select a item from Assigned Column and send to Available column
AssignedColumn(arrColumnLabel(5))

'Move a item in the assigned column up
MoveUpAssignedLabel(arrColumnLabel(9))

'Move a item in the assigned column down
MoveDownAssigendLAbel(arrColumnLabel(9))

'Click the "Set as Default" checkbox
CheckSetAsDefault()

'Save all the items in the array "LabelsAssignedColumn" to compare later 
arrLabelsAssignedColumn = NoteDownAssignedColumn()

'save the new choose view
ClickSaveBtnMyPrecerences()

'go to new Quote navbar
Navbar_CreateNewQuote()

'Scroll down
pageDownNewQuotePage()

'Validate Choose view match with the new one
ValidateChooseView(strChooseViewName)

'Validate assigned column in new quote
ValidateAssignedList_NewQuote arrLabelsAssignedColumn

'go to My Preferences in Admin Tools navbar
ClickMyPreferenceUnderAdminTools

'Set Choose Viewd
EditChooseView(strChooseViewName)

'Click the check box to uncheck the default setting
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
ClickMyPreferenceUnderAdminTools

'Set Choose View
EditChooseView(strChooseViewName)

'Click to Delete Choose view
ClickDeleteBtnChooseView()

'logout and close the browser
Navbar_Logout
Close_Browser
FinalizeTest
