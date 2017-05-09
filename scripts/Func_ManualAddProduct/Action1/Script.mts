'================================================
'Summary: Manually add products
'Description: Validate adding product, adding option,and quick add manually
'Creator: yu.li9@hpe.com
'Last Modified: 5/9/2017
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

'Fetch data
DataTable.Import "..\..\data\Func_ManualAddProduct.xlsx"
Dim strProduct1 : strProduct1 = DataTable.Value("ProductNum1",1)
Dim strProduct2 : strProduct2 = DataTable.Value("ProductNum2",1)
Dim strProduct3 : strProduct3 = DataTable.Value("ProductNum3",1)

'Dump jenkins report
dumpJenkinsOutput Environment.Value("TestName"), "000008", "Validate adding product, adding option,and quick add manually"

'Open browser
OpenNgq objUser

'Start a new quote
Navbar_CreateNewQuote

'Add a product by clicking Product or Option
LineItemDetails_AddProductByNumber2 strProduct1
ValidateAddedProduct strProduct1,1,2

'Add a product option by clicking product or option
SelectFirstLineProduct 2
Quote_AddProductOrOption
SelectOption
SetOptionValue "0D1",3
ValidateAddedProduct strProduct1,1,3

'Add product and option by quick add
Quote_QuickAdd
QuickAddProducts strProduct2,strProduct3

'Validate the quick add products
ValidateAddedProduct strProduct2, 5, 4
ValidateAddedProduct strProduct2, 5, 5
ValidateAddedProduct strProduct3, 1, 6


Navbar_Logout
Close_Browser
FinalizeTest

