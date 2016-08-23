'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "IE"

' Set opportunity id and 3rd party product number
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strOpportunityId : strOpportunityId = "OPE-0002916168"
Dim strProductNumber : strProductNumber = "M8S07A"

'NOTE: automation API calls only here. No raw UFT calls!

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValideEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

' Click on Add+
click_lineitem_add_ocs

'add 3rd party
add_product_option

' Refresh Pricing
click_refresh_pricing()

validate_product_number_line_item "HA844A1"

' Logout and close browser
Navbar_Logout()

FinalizeTest


'Add to ngq-quote
Sub click_lineitem_add_ocs()
	browser("NGQ").Page("Quote - Line Item Details").WebElement("addButton").Click
	browser("NGQ").Page("Quote - Line Item Details").WebElement("OCSConfig").Click
	browser("NGQ").Page("OCS Config").WebElement("ProductList").Click
	Browser("NGQ").Page("OCS Config").WebEdit("InputField").Set "752426-B21"
	Browser("NGQ").Page("OCS Config").WebElement("SearchButton").Click
	Browser("NGQ").Page("OCS Config").WebEdit("SetQuantity").Set "1"
	Browser("NGQ").Page("OCS Config").WebElement("Go2Bom").Click
	Browser("NGQ").Page("OCS Config").WebElement("SaveOCSConfig").Click
	Browser("NGQ").Page("OCS Config").WebElement("ConvertQuote").Click
End Sub

Sub add_product_option()
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu-1').setAttribute('style', 'left: 421.04px;');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_5').setAttribute('style', 'left: -75px;');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu-1').setAttribute('class', 'dropdown ng-scope open');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_7').setAttribute('class', 'submenu');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('item_28').getElementsByTagName('a')[0].click()"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu-1').setAttribute('class', 'dropdown ng-scope');"
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//div[@class='ngCell col5 colt5 row6']//input").Set "HA844A1"
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//div[@class='ngCell col5 colt5 row6']//input").Click
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//div[@class='ngCell col5 colt5 row6']//input").SendKeys "~"
End Sub

Sub select_product()
	browser("NGQ").Page("New Configuration").WebElement("ProductList").Click
End Sub
