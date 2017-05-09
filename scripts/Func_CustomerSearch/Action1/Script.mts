'================================================
'Summary: Customer Search
'Description: Validate populating customer info by customer searching
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
DataTable.Import "..\..\data\Func_CustomerSearch.xlsx"

'Dump the report for jenkins
dumpJenkinsOutput Environment.Value("TestName"), "000007", "Validate populating customer info by customer searching"

'Open browser.
OpenNgq objUser

'Start a new quote
Navbar_CreateNewQuote

'Open customer search dialog
Quote_CustomerDataTab
PopUpCustomerSearch

'Customer search
Dim strMDCDOrgID : strMDCDOrgID = DataTable.Value("MDCPOrgID",1)
Dim strCompanyName : strCompanyName = DataTable.Value("CompanyName",1)
Dim strStreet1 : strStreet1 = DataTable.Value("Street1",1)
Dim strCity : strCity = DataTable.Value("City",1)
Dim strZipCode : strZipCode = DataTable.Value("ZipCode",1)
Dim strCountry : strCountry = DataTable.Value("Country",1)

EnterCustomerSearchCriterias strMDCDOrgID,strCompanyName,strStreet1,strCity,strZipCode,strCountry
ClickSearchButton
CheckSoldTo
CheckBillTo
CheckShipTo
CheckReseller
CheckEndCustomer
ValidateDataPopulated strMDCDOrgID
UnCheckReseller
UnCheckEndCustomer
SubmitCustomerData

'Validate the customer info being populated correctly
Quote_CustomerDataTab
CustomerData_SoldToTab
ValidateSoldToData strMDCDOrgID
CustomerData_ShipToTab
ValidateShipToData strMDCDOrgID
CustomerData_BillToTab
ValidateBillToData strMDCDOrgID
CustomerData_ResellerTab
ValidateResellerData strMDCDOrgID
CustomerData_EndCustomerTab
ValidateEndCustomerData strMDCDOrgID

Navbar_Logout
Close_Browser
FinalizeTest

