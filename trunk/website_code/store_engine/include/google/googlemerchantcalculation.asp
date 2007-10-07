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
' Class MerchantCalculationCallback is the object representation of a 
'     <merchant-calculation-callback>.
'******************************************************************************
Class MerchantCalculationCallback
	Public Xmlns
	Public SerialNumber
	Public ItemIndex
	Public ItemArr()
	Public MerchantPrivateData
	Public CartExpirationDate
	Public BuyerId
	Public BuyerLanguage
	Public Tax
	Public AnonymousAddressIndex
	Public AnonymousAddressArr()
	Public ShippingMethodIndex
	Public ShippingMethodArr()
	Public MerchantCodeIndex
	Public MerchantCodeArr()

    '**************************************************************************
	' Parses through a <merchant-calculation-callback> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		CartExpirationDate = GetElementText(ResponseNode, "good-until-date")
		Tax = GetElementText(ResponseNode, "tax")
		BuyerId = GetElementText(ResponseNode, "buyer-id")
		BuyerLanguage = GetElementText(ResponseNode, "buyer-language")

		ItemIndex = -1
		Dim ItemNodeList
		Set ItemNodeList = ResponseNode.getElementsByTagname("item")
		If ItemNodeList.Length > 0 Then 
			Dim ItemN
			For Each ItemN In ItemNodeList
				Dim MyItem
				Set MyItem = New Item
				MyItem.ParseItem ItemN
				ItemIndex = ItemIndex + 1
				ReDim Preserve ItemArr(ItemIndex)
				Set ItemArr(ItemIndex) = MyItem
				Set MyItem = Nothing
			Next
		End If
		Set ItemNodeList = Nothing

		Dim MerchantPrivateDataNL
		Set MerchantPrivateDataNL = ResponseNode.getElementsByTagname("merchant-private-data")
		If MerchantPrivateDataNL.Length > 0 Then Set MerchantPrivateData = MerchantPrivateDataNL(0)
		Set MerchantPrivateDataNL = Nothing

		AnonymousAddressIndex = -1
		Dim AnonymousAddressNL
		Set AnonymousAddressNL = ResponseNode.getElementsByTagname("anonymous-address")
		If AnonymousAddressNL.Length > 0 Then 
			Dim AnonymousAddressN
			For Each AnonymousAddressN In AnonymousAddressNL
				Dim MyAnonymousAddress
				Set MyAnonymousAddress = New Address
				MyAnonymousAddress.ParseAddress AnonymousAddressN
				AnonymousAddressIndex = AnonymousAddressIndex + 1
				ReDim Preserve AnonymousAddressArr(AnonymousAddressIndex)
				Set AnonymousAddressArr(AnonymousAddressIndex) = MyAnonymousAddress
				Set MyAnonymousAddress = Nothing
			Next
		End If
		Set AnonymousAddressNL = Nothing

		MerchantCodeIndex = -1
		Dim MerchantCodeNL
		Set MerchantCodeNL = ResponseNode.getElementsByTagname("merchant-code-string")
		If MerchantCodeNL.Length > 0 Then 
			Dim MerchantCodeNode
			For Each MerchantCodeNode In MerchantCodeNL
				Dim MyMerchantCode
				MyMerchantCode = MerchantCodeNode.getAttribute("code")
				MerchantCodeIndex = MerchantCodeIndex + 1
				ReDim Preserve MerchantCodeArr(MerchantCodeIndex)
				MerchantCodeArr(MerchantCodeIndex) = MyMerchantCode
			Next
		End If
		Set MerchantCodeNL = Nothing

		ShippingMethodIndex = -1
		Dim ShippingMethodNL
		Set ShippingMethodNL = ResponseNode.getElementsByTagname("method")
		If ShippingMethodNL.Length > 0 Then 
			Dim ShippingMethodNode
			For Each ShippingMethodNode In ShippingMethodNL
				Dim MyShippingMethod
				MyShippingMethod = ShippingMethodNode.getAttribute("name")
				ShippingMethodIndex = ShippingMethodIndex + 1
				ReDim Preserve ShippingMethodArr(ShippingMethodIndex)
				ShippingMethodArr(ShippingMethodIndex) = MyShippingMethod
			Next
		End If
		Set ShippingMethodNL = Nothing

		Set ResponseNode = Nothing
	End Sub

	Private Sub Class_Termnitae()
		Erase ItemArr
		Erase AnonymousAddressArr
		Erase MerchantCodeArr
		Erase ShippingMethodArr
	End Sub
End Class


'******************************************************************************
' Class MerchantCalculationResults is the object representation of a 
'     <merchant-calculation-results>.
'******************************************************************************
Class MerchantCalculationResults
	Public ResultIndex
	Public ResultArr()

	Private Sub Class_Initialize()
		ResultIndex = -1
	End Sub

    '**************************************************************************
	' Adds a a Result to the merchant-calculation-results
	'
	' Input:    Result    Object: Class Result
    '**************************************************************************
	Public Sub AddResult(Result)
        ResultIndex = ResultIndex + 1
		ReDim Preserve ResultArr(ResultIndex)
		Set ResultArr(ResultIndex) = Result
	End Sub

    '**************************************************************************
	' Returns the <merchant-calculation-results> XML in plain text format.
    '**************************************************************************
	Public Property Get GetXml()
		Dim Xml, Attr
		Dim Result, CouponResult, GiftCertResult

		Set Xml = new XmlBuilder

		Attr = Xml.Attribute("xmlns", "http://checkout.google.com/schema/2")
		Xml.Push "merchant-calculation-results", Attr
		Xml.Push "results", ""
		If ResultIndex >= 0 Then
			For Each Result In ResultArr 
				If Result.ShippingName <> "" Then
					Attr = Xml.Attribute("shipping-name", Result.ShippingName)
					Attr = Attr & " " & Xml.Attribute("address-id", Result.AddressId)
					Xml.Push "result", Attr
					Attr = Xml.Attribute("currency", MerchantCurrency)
					Xml.AddElement "shipping-rate", Result.ShippingRate, Attr
					If Result.Shippable Then
						Xml.AddElement "shippable", "true", ""
					Else
						Xml.AddElement "shippable", "false", ""
					End If
				Else 
					Attr = Xml.Attribute("address-id", Result.AddressId)
					Xml.Push "result", Attr
					Attr = Xml.Attribute("currency", MerchantCurrency)
				End If

				If Result.TotalTax <> "" Then
					Xml.AddElement "total-tax", Result.TotalTax, Attr
				End If

				If Result.CouponResultIndex >= 0 Or Result.GiftCertResultIndex >= 0 Then
					Xml.Push "merchant-code-results", ""
					If Result.CouponResultIndex >= 0 Then
						For Each CouponResult In Result.CouponResultArr
							Xml.Push "coupon-result", ""
							Xml.AddElement "code", CouponResult.Code, ""
							If CouponResult.Valid Then
								Xml.AddElement "valid", "true", ""
							Else
								Xml.AddElement "valid", "false", ""
							End If
							Attr = Xml.Attribute("currency", MerchantCurrency)
							Xml.AddElement "calculated-amount", CouponResult.Amount, Attr
							Xml.AddElement "message", CouponResult.Message, ""
							Xml.Pop "coupon-result"
						Next
					End If
					If Result.GiftCertResultIndex >= 0 Then
						For Each GiftCertResult In Result.GiftCertResultArr
							Xml.Push "gift-certificate-result", ""
							Xml.AddElement "code", GiftCertResult.Code, ""
							If GiftCertResult.Valid Then
								Xml.AddElement "valid", "true", ""
							Else
								Xml.AddElement "valid", "false", ""
							End If
							Attr = Xml.Attribute("currency", MerchantCurrency)
							Xml.AddElement "calculated-amount", GiftCertResult.Amount, Attr
							Xml.AddElement "message", GiftCertResult.Message, ""
							Xml.Pop "gift-certificate-result"
						Next
					End If
					Xml.Pop "merchant-code-results"	
				End If					
				Xml.Pop "result"
			Next
		End If
		Xml.Pop "results"
		Xml.Pop "merchant-calculation-results"

		GetXml = Xml.GetXml
		Set Xml = Nothing
	End Property
	
	Private Sub Class_Termnitae()
		Erase ResultArr
	End Sub
End Class 


'******************************************************************************
' Class Result is the object representation of a <result> for 
'     <merchant-calculation-results>.
'******************************************************************************
Class Result
	Public ShippingName
	Public AddressId
	Public ShippingRate
	Public Shippable
	Public TotalTax
	Public CouponResultIndex
	Public CouponResultArr()
	Public GiftCertResultIndex
	Public GiftCertResultArr()

	Private Sub Class_Initialize()
		CouponResultIndex = -1
		GiftCertResultIndex = -1
	End Sub

    '**************************************************************************
	' Adds a CouponResult to the merchant-calculation-results
	'
	' Input:    CouponResult    Object: Class CouponResult
    '**************************************************************************
	Public Sub AddCouponResult(CouponResult)
        CouponResultIndex = CouponResultIndex + 1
		ReDim Preserve CouponResultArr(CouponResultIndex)
		Set CouponResultArr(CouponResultIndex) = CouponResult
	End Sub

    '**************************************************************************
	' Adds a GiftCertResult to the merchant-calculation-results
	'
	' Input:    GiftCertResult    Object: Class GiftCertResult
    '**************************************************************************
	Public Sub AddGiftCertResult(GiftCertResult)
        GiftCertResultIndex = GiftCertResultIndex + 1
		ReDim Preserve GiftCertResultArr(GiftCertResultIndex)
		Set GiftCertResultArr(GiftCertResultIndex) = GiftCertResult
	End Sub

	Private Sub Class_Termnitae()
		Erase CouponResultArr
		Erase GiftCertResultArr
	End Sub
End Class


'******************************************************************************
' Class CouponResult is the object representation of a 
'     <coupon-result> used in the <merchant-calculation-results>.
'******************************************************************************
Class CouponResult
    Public Code
	Public Valid
    Public Amount
    Public Message
End Class


'******************************************************************************
' Class GiftCertResult is the object representation of a 
'     <gift-certificate-result> used in the <merchant-calculation-results>.
'******************************************************************************
Class GiftCertResult
    Public Code
	Public Valid
    Public Amount
    Public Message
End Class

%>
