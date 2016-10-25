
'================================================
'Product Number:205713
'User Story: TransferaquotetoanotherNGQuserinthesamegroup
'Author: Rosales, Jahaziel Alejandro
'Description: Validate a quote number in MyDashboard
'Tags:
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "<Encrypted Digitalbadge>")
DataTable.Import "..\..\data\US9430_04.xlsx"
Dim strQuoteNumber : strQuotenumber = DataTable("strQuotenumber",1)

InitializeTest "IE"

'open web browser and go to NGQ/My Dashboard
OpenNgq(objUser)
ClickMyDashboard()

'validate if QuoteTab is selected
ValidateQuoteTab()

'Click auto Filter Button
ClickAutoFilter()

'set and submit Quote Number
FillFilterQuoteNumber("NI00159743")
'FillFilterQuoteNumber(strQuotenumber)
ClickQuoteNumber(2)
'Validate the submit value match with the value in table
ValidateQuoteNumberValue("NI00159743")
'ValidateQuoteNumberValue(strQuotenumber)

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest

