﻿'================================================
'Project Number:205713
'User Story: SearchQuotebymultiplesearchcriteria
'Description: Use severals criterias for search/validate in AdvancedSearch 
'Tags:
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime
InitializeTest "IE"

'Change the sync time for this script from 1 min to 3
UFT.BrowserNavigationTimeout = 200000

Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>","<Encrypted DigitalBadge>")
DataTable.Import "..\..\data\US9430_01.xlsx"

Dim strOpportunityId : strOpportunityId = DataTable("strOpportunityId",1)
Dim strQuoterNumber : strQuoterNumber = DataTable("strQuoterNumber",1)
Dim strMCDPId : strMCDPId = DataTable("strMCDPId",1)
Dim strAccountName : strAccountName = DataTable("strAccountName",1)
Dim strEmail : strEmail = DataTable("strEmail",1)
Dim strStartDate : strStartDate = DataTable("strStartDate",1)
Dim strEndDate : strEndDate = DataTable("strEndDate",1)

'open brower  
OpenNgq(objUser)

'search opportunity ID, this will trigger the advanced searchd /valdiat OpportuityID in advanced search
SetOpportunityId(strOpportunityId)
QuickSearch_Search()
Validate_OpportunityId_AdvancedSearch(strOpportunityId)

'reset the search
ClickResetButton_advancedSearch()
ClickNavbarAdvancedSearch()

'search for MDCP ID in advanced Search
MDCPIdAdvancedSearch(strMCDPId)
ClickSearchButton_advancedSearch()
ValidateMDCPIAdvancedSearch(strMCDPId)

''reset the search
ClickResetButton_advancedSearch()

'search for quoteNumber in advanced Search
QuoteNumber_AdvancedSearch(strQuoterNumber)
ClickSearchButton_advancedSearch()
Validate_QuoteNumer_AdvancedSearch(strQuoterNumber)

'reset the search
ClickResetButton_advancedSearch()

'Search and validate email
LastModifedEmail(strEmail)
ClickSearchButton_advancedSearch()
ValidateLasModifedEmail(strEmail)

'reset the search
ClickResetButton_advancedSearch()

'search and validate Accoun Name/Company Name
CompanyNameAccountName(strAccountName)
ClickSearchButton_advancedSearch()
ValidateCompanyName(strAccountName)

'reset the search
ClickResetButton_advancedSearch()

'Search by date
SetStartDate(strStartDate)
SetEndDate(strEndDate)
ClickSearchButton_advancedSearch()
ClickQuoteNumberResult(2)
ValidateDateRange strStartDate,strEndDate

'logout and close the browser
Navbar_Logout()
Browser("NGQ").Close()

FinalizeTest()

