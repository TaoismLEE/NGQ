'================================================
'Summary: Import Config from OCA
'Description: System support importing config from OCA
'Creator: yu.li9@hpe.com
'Last Modified: 4/25/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data
DataTable.Import "..\..\data\Func_ImportConfigFromOCA.xlsx"
Dim strConfiID : strConfiID = DataTable.Value("ConfigID",1)

'Dump jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "000008", "System support importing config from OCA"

'Open browser
OpenNgq objUser
Navbar_CreateNewQuote
ClickImportConfigFromOCA
UploadConfigFromOCA strConfiID

'Check the products are imported
Dim arrProducts : arrProducts = GetSourceDataFromExcel
ValidateValidProducts arrProducts

'Exist test
Navbar_Logout
Close_Browser
FinalizeTest

