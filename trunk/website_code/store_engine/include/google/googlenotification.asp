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
' Class NewOrderNotification is the object representation of a 
'     <new-order-notification>.
'******************************************************************************
Class NewOrderNotification
	Public Xmlns
	Public SerialNumber
	Public GoogleOrderNumber
	Public TimeStamp
	Public ItemIndex
	Public ItemArr()
	Public MerchantPrivateData
	Public TaxTotal
	Public ShippingName
	Public ShippingCost
	Public AdjustmentTotal
	Public MarketingEmailAllowed
	Public OrderTotal
	Public FulfillmentOrderState
	Public FinancialOrderState
	Public BuyerId
	Public BuyerBillingAddress
	Public BuyerShippingAddress
	Public MerCalcSuccessful
	Public CouponIndex
	Public CouponArr()
	Public GiftCertIndex
	Public GiftCertArr()

    '**************************************************************************
	' Parses through a <new-order-notification> node and makes the object 
	'     ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		MarketingEmailAllowed = GetElementText(ResponseNode, "email-allowed")
		AdjustmentTotal = GetElementText(ResponseNode, "adjustment-total")
		TaxTotal = GetElementText(ResponseNode, "total-tax")
		ShippingName = GetElementText(ResponseNode, "shipping-name")
		ShippingCost = GetElementText(ResponseNode, "shipping-cost")
		OrderTotal = GetElementText(ResponseNode, "order-total")
		FulfillmentOrderState = GetElementText(ResponseNode, "fulfillment-order-state")
		FinancialOrderState = GetElementText(ResponseNode, "financial-order-state")
		BuyerId = GetElementText(ResponseNode, "buyer-id")
		MerCalcSuccessful = GetElementText(ResponseNode, "merchant-calculation-successful")

		Dim MerchantPrivateDataNL
		Set MerchantPrivateDataNL = ResponseNode.getElementsByTagname("merchant-private-data")
		If MerchantPrivateDataNL.Length > 0 Then Set MerchantPrivateData = MerchantPrivateDataNL(0)
		Set MerchantPrivateDataNL = Nothing

		Dim AddressNL, MyAddress

		Set AddressNL = ResponseNode.getElementsByTagname("buyer-billing-address")
		If AddressNL.Length > 0 Then 
			Set MyAddress = New Address
			MyAddress.ParseAddress AddressNL(0)
			Set BuyerBillingAddress = MyAddress
			Set MyAddress = Nothing
		End If
		Set AddressNL = Nothing
			
		Set AddressNL = ResponseNode.getElementsByTagname("buyer-shipping-address")
		If AddressNL.Length > 0 Then 
			Set MyAddress = New Address
			MyAddress.ParseAddress AddressNL(0)
			Set BuyerShippingAddress = MyAddress
			Set MyAddress = Nothing
		End If
		Set AddressNL = Nothing
		
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

		CouponIndex = -1
		Dim CouponNL
		Set CouponNL = ResponseNode.getElementsByTagname("coupon-adjustment")
		If CouponNL.Length > 0 Then 
			Dim CouponNode
			For Each CouponNode In CouponNL
				Dim MyCoupon
				Set MyCoupon = New CouponAdjustment
				MyCoupon.ParseCoupon CouponNode
				CouponIndex = CouponIndex + 1
				ReDim Preserve CouponArr(CouponIndex)
				Set CouponArr(CouponIndex) = MyCoupon
				Set MyCoupon = Nothing
			Next
		End If
		Set CouponNL = Nothing

		GiftCertIndex = -1
		Dim GiftCertNL
		Set GiftCertNL = ResponseNode.getElementsByTagname("gift-certificate-adjustment")
		If GiftCertNL.Length > 0 Then 
			Dim GiftCertNode
			For Each GiftCertNode In GiftCertNL
				Dim MyGiftCert
				Set MyGiftCert = New GiftCertAdjustment
				MyGiftCert.ParseGiftCert GiftCertNode
				GiftCertIndex = GiftCertIndex + 1
				ReDim Preserve GiftCertArr(GiftCertIndex)
				Set GiftCertArr(GiftCertIndex) = MyGiftCert
				Set MyGiftCert = Nothing
			Next
		End If
		Set GiftCertNL = Nothing

		Set ResponseNode = Nothing
	End Sub

	Private Sub Class_Termnitae()
		Erase ItemArr
		Erase CouponArr
		Erase GiftCertArr
		Set BuyerBillingAddress = Nothing
		Set BuyerShippingAddress = Nothing
	End Sub
End Class


'******************************************************************************
' Class OrderStateChangeNotification is the object representation of a 
'     <order-state-change-notification>.
'******************************************************************************
Class OrderStateChangeNotification
	Public Xmlns
	Public SerialNumber
	Public TimeStamp
	Public GoogleOrderNumber
	Public NewFulfillmentOrderState
	Public NewFinancialOrderState
	Public PreviousFulfillmentOrderState
	Public PreviousFinancialOrderState

    '**************************************************************************
	' Parses through a <order-state-change-notification> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		NewFulfillmentOrderState = GetElementText(ResponseNode, "new-fulfillment-order-state")
		NewFinancialOrderState = GetElementText(ResponseNode, "new-financial-order-state")
		PreviousFulfillmentOrderState = GetElementText(ResponseNode, "previous-fulfillment-order-state")
		PreviousFinancialOrderState = GetElementText(ResponseNode, "previous-financial-order-state")
		Set ResponseNode = Nothing
	End Sub
End Class


'******************************************************************************
' Class RiskInformationNotification is the object representation of a 
'     <risk-information-notification>.
'******************************************************************************
Class RiskInformationNotification
	Public Xmlns
	Public SerialNumber
	Public TimeStamp
	Public GoogleOrderNumber
	Public BillingAddress
	Public IpAddress
	Public EligibleForProtection
	Public AvsResponse
	Public CvnResponse
	Public PartialCC
	Public BuyerAccountAge

    '**************************************************************************
	' Parses through a <risk-information-notification> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		AvsResponse = GetElementText(ResponseNode, "avs-response")
		CvnResponse = GetElementText(ResponseNode, "cvn-response")
		IpAddress = GetElementText(ResponseNode, "ip-address")
		EligibleForProtection = GetElementText(ResponseNode, "eligible-for-protection")
		PartialCC = GetElementText(ResponseNode, "partial-cc-number")
		BuyerAccountAge = GetElementText(ResponseNode, "buyer-account-age")

		Dim AddressNL
		Set AddressNL = ResponseNode.getElementsByTagname("billing-address")
		If AddressNL.Length > 0 Then 
			Dim MyAddress
			Set MyAddress = New Address
			MyAddress.ParseAddress AddressNL(0)
			Set BillingAddress = MyAddress
			Set MyAddress = Nothing
		End If
		Set AddressNL = Nothing

		Set ResponseNode = Nothing
	End Sub

	Private Sub Class_Termnitae()
		Set BillingAddress = Nothing
	End Sub
End Class


'******************************************************************************
' Class ChargeAmountNotification is the object representation of a 
'     <charge-amount-notification>.
'******************************************************************************
Class ChargeAmountNotification
	Public Xmlns
	Public SerialNumber
	Public TimeStamp
	Public GoogleOrderNumber
	Public LatestChargeAmount
	Public TotalChargeAmount

    '**************************************************************************
	' Parses through a <charge-amount-notification> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		LatestChargeAmount = GetElementText(ResponseNode, "latest-charge-amount")
		TotalChargeAmount = GetElementText(ResponseNode, "total-charge-amount")
		Set ResponseNode = Nothing
	End Sub
End Class


'******************************************************************************
' Class ChargebackAmountNotification is the object representation of a 
'     <chargeback-amount-notification>.
'******************************************************************************
Class ChargebackAmountNotification
	
	Public Xmlns
	Public SerialNumber
	Public TimeStamp
	Public GoogleOrderNumber
	Public LatestChargebackAmount
	Public TotalChargebackAmount

    '**************************************************************************
	' Parses through a <chargeback-amount-notification> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		LatestChargebackAmount = GetElementText(ResponseNode, "latest-chargeback-amount")
		TotalChargebackAmount = GetElementText(ResponseNode, "total-chargeback-amount")
		Set ResponseNode = Nothing
	End Sub
End Class


'******************************************************************************
' Class AuthorizationAmountNotification is the object representation of a 
'     <authorization-amount-notification>.
'******************************************************************************
Class AuthorizationAmountNotification
	Public Xmlns
	Public SerialNumber
	Public TimeStamp
	Public GoogleOrderNumber
	Public AvsResponse
	Public CvnResponse
	Public AuthorizationAmount
	Public AuthorizationExpirationDate

    '**************************************************************************
	' Parses through a <authorization-amount-notification> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		AvsResponse = GetElementText(ResponseNode, "avs-response")
		CvnResponse = GetElementText(ResponseNode, "cvn-response")
		AuthorizationAmount = GetElementText(ResponseNode, "authorization-amount")
		AuthorizationExpirationDate = GetElementText(ResponseNode, "authorization-expiration-date")
		Set ResponseNode = Nothing
	End Sub
End Class


'******************************************************************************
' Class RefundAmountNotification is the object representation of a 
'     <refund-amount-notification>.
'******************************************************************************
Class RefundAmountNotification
	Public Xmlns
	Public SerialNumber
	Public TimeStamp
	Public GoogleOrderNumber
	Public LatestRefundAmount
	Public TotalRefundAmount

    '**************************************************************************
	' Parses through a <refund-amount-notification> node and makes the 
	'     object ready to be used.
    '**************************************************************************
	Public Sub ParseNotification(ResponseXml)
		Dim ResponseNode
		Set ResponseNode = GetRootNode(ResponseXml)
		Xmlns = ResponseNode.getAttribute("xmlns")
		SerialNumber = ResponseNode.getAttribute("serial-number")
		TimeStamp = GetElementText(ResponseNode, "timestamp")
		GoogleOrderNumber = GetElementText(ResponseNode, "google-order-number")
		LatestRefundAmount = GetElementText(ResponseNode, "latest-refund-amount")
		TotalRefundAmount = GetElementText(ResponseNode, "total-refund-amount")
		Set ResponseNode = Nothing
	End Sub
End Class


'******************************************************************************
' Class Address is the object representation of <anonymous-address>, 
'     <buyer-billing-address>, <buyer-shipping-address> and <billing-address>
'******************************************************************************
Class Address
	Public ID
	Public ContactName
	Public CompanyName
	Public Address1
	Public Address2
	Public City
	Public Region
	Public PostalCode
	Public CountryCode
	Public Email
	Public Phone
	Public Fax

    '**************************************************************************
	' Parses through an address node and makes the object ready to be used.
    '**************************************************************************
	Public Sub ParseAddress(AddressNode)
		ID = AddressNode.getAttribute("id")
		ContactName = GetElementText(AddressNode, "contact-name")
		CompanyName = GetElementText(AddressNode, "company-name")
		Address1 = GetElementText(AddressNode, "address1")
		Address2 = GetElementText(AddressNode, "address2")
		City = GetElementText(AddressNode, "city")
		Region = GetElementText(AddressNode, "region")
		PostalCode = GetElementText(AddressNode, "postal-code")
		CountryCode = GetElementText(AddressNode, "country-code")
		Email = GetElementText(AddressNode, "email")
		Phone = GetElementText(AddressNode, "phone")
		Fax = GetElementText(AddressNode, "fax")
	End Sub
End Class


'******************************************************************************
' Class Coupon is the object representation of a 
'     <coupon-adjustment> in the <new-order-notification>
'******************************************************************************
Class CouponAdjustment
    Public AppliedAmount
    Public Code
    Public CalculatedAmount
    Public Message

    '**************************************************************************
	' Parses through a <coupon-adjument> node and makes the object 
	'     ready to be used.
    '**************************************************************************
	Public Sub ParseCoupon(CouponNode)
		AppliedAmount = GetElementText(CouponNode, "applied-amount")
		Code = GetElementText(CouponNode, "code")
		CalculatedAmount = GetElementText(CouponNode, "calculated-amount")
		Message = GetElementText(CouponNode, "message")
	End Sub
End Class


'******************************************************************************
' Class GiftCert is the object representation of a 
'     <gift-certificate-adjustment> in the <new-order-notification>
'******************************************************************************
Class GiftCertAdjustment
    Public AppliedAmount
    Public Code
    Public CalculatedAmount
    Public Message

    '**************************************************************************
	' Parses through a <gift-certificate-adjument> node and makes the object 
	'     ready to be used.
    '**************************************************************************
	Public Sub ParseGiftCert(GiftCertNode)
		AppliedAmount = GetElementText(GiftCertNode, "applied-amount")
		Code = GetElementText(GiftCertNode, "code")
		CalculatedAmount = GetElementText(GiftCertNode, "calculated-amount")
		Message = GetElementText(GiftCertNode, "message")
	End Sub
End Class


%>
