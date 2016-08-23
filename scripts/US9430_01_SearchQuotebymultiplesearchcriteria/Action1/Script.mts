'================================================
'Summary:
'
'Description:
'
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime

InitializeTest

Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>")
Dim strOpportunityId : strOpportunityId = "OPE-0005373487"
Dim strQuoterNumber : strQuoterNumber = "NQ00002291"
Dim strMCDPId : strMCDPId = "145128717"
Dim strAccountName : strAccountName = "XX_WW_TH_TEST_ACCOUNT"

'open brower quick search opportunity ID
OpenNgq(objUser)
QuickSearch_OpportunityId(strOpportunityId)
QuickSearch_Search()

'valdiat OpportuityID in advanced search
Validate_OpportunityId_AdvancedSearch(strOpportunityId) @@ hightlight id_;_Browser("Home 2").Page("Home").WebElement("MDCP Org ID")_;_script infofile_;_ZIP::ssf16.xml_;_

FinalizeTeste





'======================
'function for this scripts
'=======================

'Search Opportunity Id in QUICK SEARCH (HOME)
Function QuickSearch_OpportunityId(strOpportunityId)
	Browser("NGQ").Page("Home").WebEdit("quick opportunity id").Set strOpportunityId
End Function

Function Validate_OpportunityId_AdvancedSearch(opportunityId)

Dim strTempOpportunityId : strTempOpportunityId = Browser("name:=Home.*").Page("title:=Home.*").WebEdit("//span[@class='suffix-colon ng-binding' and text()=""Opportunity ID""]/following-sibling::input")
'//span[@class='suffix-colon ng-binding' and text()="Opportunity ID"]/following-sibling::input
	If  strTempOpportunityId = opportunityId Then
		'do nothing
	Else
		Browser("name:=Home.*").Page("title:=Home.*").WebEdit("//span[@class='suffix-colon ng-binding' and text()=""Opportunity ID""]/following-sibling::input").Set opportunityId
	End If
	
End Function
