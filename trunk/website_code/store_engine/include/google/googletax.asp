<%
'******************************************************************************
' Copyright (C) 2006 Google Inc.
'  
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'  
'      http://www.apache.org/licenses/LICENSE-2.0
'  
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'******************************************************************************

'******************************************************************************
' Class DefaultTaxRule is the object representation of a <default-tax-rule>.
'******************************************************************************
Class DefaultTaxRule
	Public Rate
	Public ShippingTaxed
	Public CountryArea
	Public States
	Public ZipPatterns

    Private Sub Class_Initialize()
		Rate = 0
		ShippingTaxed = false
		CountryArea = ""
		States = Array()
		ZipPatterns = Array()
    End Sub
End Class

'******************************************************************************
' Class AlternateTaxRule is the object representation of a 
'     <alternate-tax-rule>.
'******************************************************************************
Class AlternateTaxRule
	Public Rate
	Public CountryArea
	Public States
	Public ZipPatterns

    Private Sub Class_Initialize()
		Rate = 0
		CountryArea = ""
		States = Array()
		ZipPatterns = Array()
    End Sub
End Class


'******************************************************************************
' Class AlternateTaxTable is the object representation of a 
'     <alternate-tax-table>.
'******************************************************************************
Class AlternateTaxTable
	Public Name
	Public Standalone
	Public AlternateTaxRuleIndex
	Public AlternateTaxRuleArr()

    Private Sub Class_Initialize()
		Standalone = false
		AlternateTaxRuleIndex = -1
    End Sub

    '**************************************************************************
	' Adds an alternate tax rule to this alternate tax table.
	'
	' Input:    TaxRule    Object: Class AlternateTaxRule
    '**************************************************************************
	Public Sub AddTaxRule(TaxRule)
        AlternateTaxRuleIndex = AlternateTaxRuleIndex + 1
		ReDim Preserve AlternateTaxRuleArr(AlternateTaxRuleIndex)
		Set AlternateTaxRuleArr(AlternateTaxRuleIndex) = TaxRule
	End Sub

	Private Sub Class_Terminate()
		Erase AlternateTaxRuleArr
	End Sub
End Class

%>