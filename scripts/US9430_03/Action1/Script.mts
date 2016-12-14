﻿'================================================
'Product Number:205713
'User Story: WithoutbeingpartofthesalesteamNGQuseraccessandeditothersquoteafterclonethequote
'Author: Rosales, Jahaziel Alejandro
'Description: Select one QuoteNumber in MyDashboard, then clone it and change the company name
'Tags:
'================================================

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

'Load the xls file for the user information
DataTable.Import "..\..\data\NGQ_empty_quote_data.xlsx"
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")

DataTable.Import "..\..\data\US9430_03.xlsx"
'Dim strQuoteNumber : strQuoteNumber = DataTable("strQuoteNumber",1)
Dim strCompanyName : strCompanyName = DataTable("strCompanyName",1)

InitializeTest "Action1"

'Open brower and go to My Dashboard
OpenNgq(objUser)
ClickMyDashboard()

'Go to Group Quote Tab
ClickMyGroupQuoteTab()

'Click in the first row number
ClickMyGroupStatusCount()

'Click the Auto filter Btn and enter the value
ClickAutoFilter()
FillFilterQuoteNumber("NI00161552") 'NI00159734
'FillFilterQuoteNumber("NI00159591")
'FillFilterQuoteNumber(strQuoteNumber)

'click the quoete number value
ClickQuoteNumber(2)

'Clone the Quote and save it
Click_Clone()
Quote_save()

'Go to Customer Data Tab and change the Company Name and save it
Quote_CustomerDataTab()
ClickCompanyPencilBtn()
EditCompanyName(strCompanyName)
Quote_save()

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest

