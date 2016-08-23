'================================================
'Summary: A demo for NGQ
'
'Description:
'Demo is as demo does.
'
'Preconditions:
'Recommended: Use programing descriptive not objects repository
'Author: Ana Karina Orduna
'
'Notes:
'Syncing is a real problem when the app is not responding quickly.
'Spinners/loading dialogs don't appear immediately on section transitions.
'================================================
Option Explicit
Dim al : Set al = NewActionLifetime


'InitializeTest "CH"
InitializeTest ""
'Hard-coded data.
Dim objUser : Set objUser = NewRealUser(Parameter("username"), Parameter("password"))
Dim emptyQuoteNumber : emptyQuoteNumber = "New Quote"
Dim emptyQuoteVersion : emptyQuoteVersion = "01"
Dim emptyQuoteStatus : emptyQuoteStatus = "Quote/Configuration Created"
Dim emptyQuoteEndDate : emptyQuoteEndDate = "Need Pricing Call"
Dim emptyQuoteSelectedTab : emptyQuoteSelectedTab = "Opportunity and Quote Info"
Dim opportunityID : opportunityID = "OPE-0002907630"
Dim quoteName : quoteName = "Test Quote"

'Open browser.
OpenNgq objUser

Navbar_CreateNewQuote

NewQuote_ValideEmptyQuote emptyQuoteNumber,emptyQuoteVersion,emptyQuoteStatus,emptyQuoteEndDate

Quote_currentlySelectedTab emptyQuoteSelectedTab

OpportunityAndQuoteInfo_SetOpportunityId opportunityID

OpportunityAndQuoteInfo_Import

Quote_EditQuoteName quoteName

Quote_save

Dim quoteID : quoteID = Quote_get_quoteNumber

LineItemDetails_AddProductByNumber "U7BM9E", "1"
LineItemDetails_AddProductByNumber "AF556A", "1"
LineItemDetails_AddProductByNumber "BW946A", "1"
LineItemDetails_AddProductByNumber "G1S72A", "1"

Quote_PricingTermsTab
Quote_refreshPricing
verify_price_quality_band
Quote_save
quoteID = Quote_get_quoteNumber

applyEmpowerment "Preferred"
applyEmpowerment "Personal"
applyEmpowerment "Manager"

FinalizeTest

Sub applyEmpowerment(eType)
	Dim row, actual, rowcount, Iterator, xpath, index
	Select Case eType
		Case "Preferred"
			reset_DataTable
			Browser("NGQ").Page("Pricing and Terms").WebElement("ApplyPreferredEmpower").Click
			If Browser("NGQ").Page("Pricing and Terms").WebElement("ExistingDealRemoveAlert").Exist Then
				Browser("NGQ").Page("Pricing and Terms").WebElement("ExistingDealRemoveAlert").Click
			End If
			rowcount = DataTable.GetRowCount
			For Iterator = 1 To rowcount Step 1
				row = DataTable.GetCurrentRow
				'actual = Browser("NGQ").Page("Pricing and Terms").WebElement("TotalRequestedDiscount","index:="&Iterator).GetROProperty("innertext")
				index = Iterator + 1
				xpath = "xpath:=(//div[@id='thresholdLocation']//div[contains(@class,'col13')])[" & index & "]"
				actual = Browser("NGQ").Page("Pricing and Terms").WebElement(xpath).GetROProperty("innertext")
				If actual = DataTable("Preferred") Then
					Reporter.ReportEvent micPass, "Verify Updated Empowerment Pricing", "Verified Correct Empowerment Discount"
				else
					Reporter.ReportEvent micFail, "Verify Updated Empowerment Pricing", "Incorrect Empowerment Discount " & actual & " returned"
				End If
				row = DataTable.SetNextRow
			Next
		Case "Personal"
			reset_DataTable
			Browser("NGQ").Page("Pricing and Terms").WebElement("ApplyPersonalEmpowerment").Click
			If Browser("NGQ").Page("Pricing and Terms").WebElement("ExistingDealRemoveAlert").Exist Then
				Browser("NGQ").Page("Pricing and Terms").WebElement("ExistingDealRemoveAlert").Click
			End If
			rowcount = DataTable.GetRowCount
			For Iterator = 1 To rowcount Step 1
				row = DataTable.GetCurrentRow
				'actual = Browser("NGQ").Page("Pricing and Terms").WebElement("TotalRequestedDiscount","index:="&Iterator).GetROProperty("innertext")
				index = Iterator + 1
				xpath = "xpath:=(//div[@id='thresholdLocation']//div[contains(@class,'col13')])[" & index & "]"
				actual = Browser("NGQ").Page("Pricing and Terms").WebElement(xpath).GetROProperty("innertext")
				If actual = DataTable("Personal") Then
					Reporter.ReportEvent micPass, "Verify Updated Empowerment Pricing", "Verified Correct Empowerment Discount"
				else
					Reporter.ReportEvent micFail, "Verify Updated Empowerment Pricing", "Incorrect Empowerment Discount " & actual & " returned"
				End If
				row = DataTable.SetNextRow
			Next
		Case "Manager"
			reset_DataTable
			Browser("NGQ").Page("Pricing and Terms").WebElement("ApplyManagerEmpowerment").Click
			If Browser("NGQ").Page("Pricing and Terms").WebElement("ExistingDealRemoveAlert").Exist Then
				Browser("NGQ").Page("Pricing and Terms").WebElement("ExistingDealRemoveAlert").Click
			End If
			rowcount = DataTable.GetRowCount
			For Iterator = 1 To rowcount Step 1
				row = DataTable.GetCurrentRow
				'actual = Browser("NGQ").Page("Pricing and Terms").WebElement("TotalRequestedDiscount","index:="&Iterator).GetROProperty("innertext")
				index = Iterator + 1
				xpath = "xpath:=(//div[@id='thresholdLocation']//div[contains(@class,'col13')])[" & index & "]"
				actual = Browser("NGQ").Page("Pricing and Terms").WebElement(xpath).GetROProperty("innertext")
				If actual = DataTable("Manager") Then
					Reporter.ReportEvent micPass, "Verify Updated Empowerment Pricing", "Verified Correct Empowerment Discount"
				else
					Reporter.ReportEvent micFail, "Verify Updated Empowerment Pricing", "Incorrect Empowerment Discount " & actual & " returned"
				End If
				row = DataTable.SetNextRow
			Next
	End Select
End Sub

Sub reset_DataTable()
	Dim it, curRow, numRow
	numRow = DataTable.GetRowCount
	For it = curRow To 1 Step -1
		DataTable.SetPrevRow
	Next
End Sub

