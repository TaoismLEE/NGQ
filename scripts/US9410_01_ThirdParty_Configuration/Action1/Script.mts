﻿'Project Number: 205713
'User Story: US9410_01_ThirdParty_Configuration
'Description: Validates that the third party product can be included in a configuration
'Tags: Configuration, third party, product

Option Explicit
Dim al : Set al = NewActionLifetime
SystemUtil.CloseProcessByName "IEXPLORE.EXE"

InitializeTest "Action1"

DataTable.Import "..\..\data\data_file.xlsx"

' Set opportunity id and 3rd party product number
'Dim objUser : Set objUser = NewRealUser("<username>", "<encrypted password>", "a")
Dim objUser : Set objUser = NewRealUser(DataTable.Value("user", "Global"), DataTable.Value("pass", "Global"), "<Encrypted DigitalBadge>")
Dim strOpportunityId : strOpportunityId = DataTable.Value("oppID", "Global")
Dim strProductNumber : strProductNumber = DataTable.Value("otherProductNumber", "Global")

'NOTE: automation API calls only here. No raw UFT calls!
'For Jenkins Reporting
dumpJenkinsOutput Environment.Value("TestName"), "74229", "CPQEncoreRetirement_US9410_01_IncludethirdpartypartsintheconfigurationfromOCS"

' Open the NGQ website
OpenNgq objUser

'Navigate to "New quote tab" and click "New Quote" and validate it is an empty quote
Navbar_CreateNewQuote
NewQuote_ValidateEmptyQuote "New Quote", "1", "Quote/Configuration Created", "Need Pricing Call"

'Enter an Opportunity ID in the "Import Opportunity ID/Request ID" section. Click import
OpportunityAndQuoteInfo_ImportOpportunityId strOpportunityId

'Add config from OCS
build_ocs_bom

'add 3rd party
add_product_option strProductNumber

' Refresh Pricing
click_refresh_pricing()

validate_product_number_line_item strProductNumber

' Logout and close browser
Navbar_Logout()

browser("NGQ").close()

FinalizeTest



