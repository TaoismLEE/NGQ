'================================================
'Test Case: CPQ_Encore Retirement_US9400_Item Actions_08
'
'Preconditions:
'1. Sales op has access to NGQ.
'2. An Opportunity ID is ready.
'
'Recommended: Use programing descriptive not objects repository
'Author: Ana Karina Orduña
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest "CH"
DataTable.Import "..\..\data\TD_NGQ_CPQ_EncoreRetirement_US9400_ItemActions_08.xlsx"

'Hard-coded data.
Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "a")

' Variable Decalration
Dim strQuoteNumberID : strQuoteNumberID = DataTable.Value("QuoteNumberID","Global")
Dim strQuoteVersion : strQuoteVersion = DataTable.Value("QuoteVersion","Global")
Dim strQuoteStatus : strQuoteStatus = DataTable.Value("QuoteStatus","Global")
Dim strQuoteEndDate : strQuoteEndDate = DataTable.Value("QuoteEndDate","Global")
Dim strOpportunityId : strOpportunityId = DataTable.Value("OpportunityID","Global")
Dim strProductNumber : strProductNumber = DataTable.Value("ProductNumber","Global")
Dim strQuoteName : strQuoteName = DataTable.Value("QuoteName","Global") 
Dim strQuotaSelection_Selector : strQuotaSelection_Selector = DataTable.Value("QuotaSelection_Selector","Global") 
Dim strDeliverySpeed : strDeliverySpeed = DataTable.Value("DeliverySpeed","Global") 
Dim strDeliveryTerms : strDeliveryTerms = DataTable.Value("DeliveryTerms","Global") 
Dim strLineItemSelector: strLineItemSelector = DataTable.Value("LineItemSelector","Global")

'START: Core
OpenNgq objUser
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId
Quote_EditQuoteName strQuoteName
click_save_button()

' Build ocs configuration
build_ocs_bom @@ hightlight id_;_Browser("Home").Page("Home 2").Link("NI00155377")_;_script infofile_;_ZIP::ssf2.xml_;_

'CustomerData_ShipTo

'CustomerDataShipTo_SelectSameAsSoldToAddress

' Click shipping data tab
'Quote_ShippingDataTab

' Set speed
'ShippingData_SetDeliverySpeed strDeliverySpeed

' Set Delivery terms
'ShippingData_SetTermsOfDelivery strDeliveryTerms

' Set receipt date
Quote_AdditionalDataTab

'AdditionalData_SetReceiptDateNow

'Refresh Pricing
click_refresh_pricing()


rightClickAddPageBreak()

rightClickAddComment()

selectMultipleLines

removeItemInSubTotal

deleteSubTotalLine

removeItem

click_save_button()

click_refresh_pricing()

editProductNum


Sub rightClickAddPageBreak()
wait(3)
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'752426-B21')])[1]").RightClick
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_7').setAttribute('class', 'submenu');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('item_29').getElementsByTagName('a')[0].click()"
	Browser("name:=Home.*").Page("title:=Home.*").WebList("xpath:=//div[@role='row']//select").SelectByText "Page Break"
End Sub

Sub rightClickAddComment()
wait(3)
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'H1K92A3')])[1]").RightClick
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_7').setAttribute('class', 'submenu');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('item_29').getElementsByTagName('a')[0].click()"
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//span[@class='wrap ng-scope']/input").Click
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//span[@class='wrap ng-scope']/input").Set "A Comment"
End Sub

Sub selectMultipleLines()
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'752426-B21')])[1]").Click
	Dim obj
	Set obj = CreateObject("Mercury.DeviceReplay")
	obj.KeyDown 42
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'HA114A1')])[1]").Click
	obj.KeyUp 42
	UFT.ReplayType = 1
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//a[contains(text(), 'Add Subtotal')])[1]").Click
End Sub

Sub removeItemInSubTotal()
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'HA114A1')])[2]").RightClick
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_7').setAttribute('class', 'submenu');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('item_29').getElementsByTagName('a')[0].click()"
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=//div[@id='grid_msgs']//a").Click
End Sub

Sub removeItem()
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'HA114A1')])[1]").RightClick
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_7').setAttribute('class', 'submenu');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('item_30').getElementsByTagName('a')[0].click()"
End Sub

Sub deleteSubTotalLine()
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[@class='icon-trash ng-scope'])[last()]").Click
End Sub

Sub editProductNum()
	UFT.ReplayType = 2
	Browser("name:=Home.*").Page("title:=Home.*").WebElement("xpath:=(//span[contains(text(),'H1K92A3')])[1]").RightClick
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('menu_7').setAttribute('class', 'submenu');"
	Browser("NGQ").Page("Upload Config").RunScript "document.getElementById('item_31').getElementsByTagName('a')[0].click()"
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//span[@title='Product No.']//input").click
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//span[@title='Product No.']//input").Set "ABC"
	Browser("name:=Home.*").Page("title:=Home.*").WebEdit("xpath:=//span[@title='Product No.']//input").SendKeys("~")
End Sub
