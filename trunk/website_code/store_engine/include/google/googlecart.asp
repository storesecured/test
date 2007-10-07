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
' Class Cart is the object representation of a <checkout-shopping-cart>.
' It has methods to create and post a shopping cart to Google Checkout.
'******************************************************************************
Class Cart
	Public MerchantCalculationsUrl
	Public AcceptMerchantCoupons
	Public AcceptGiftCertificates
	Public MerchantCalculatedTax
    Public CartExpiration
    Public MerchantPrivateData
    Public EditCartUrl
    Public ContinueShoppingUrl
    Public RequestBuyerPhoneNumber
	Public RequestInitialAuthDetails
	Public PlatformID

	Public ItemArr()
	Public ShippingArr()
	Public DefaultTaxRuleArr()
	Public AlternateTaxTableArr()

    Private ItemIndex
	Private ShippingIndex
	Private DefaultTaxRuleIndex
	Private AlternateTaxTableIndex

    Private Sub Class_Initialize()
		AcceptMerchantCoupons = false
		AcceptGiftCertificates = false
		MerchantCalculatedTax = false
    	RequestBuyerPhoneNumber = false
		RequestInitialAuthDetails = false

		ItemIndex = -1
		ShippingIndex = -1
		DefaultTaxRuleIndex = -1
		AlternateTaxTableIndex = -1	
	End Sub

    '**************************************************************************
	' Adds an item to the cart.
	'
	' Input:    Item    Object: Class Item
    '**************************************************************************
    Public Sub AddItem(Item)
        ItemIndex = ItemIndex + 1
		ReDim Preserve ItemArr(ItemIndex)
		Set ItemArr(ItemIndex) = Item
    End Sub

    '**************************************************************************
	' Adds a shipping method to the cart.
	'
	' Input:    Shipping    Object: Class FlatRateShipping, or
	'                               Class MerchantCalculatedShipping, or
	'                               Class Pickup
    '**************************************************************************
    Public Sub AddShipping(Shipping)
        ShippingIndex = ShippingIndex + 1
		ReDim Preserve ShippingArr(ShippingIndex)
		Set ShippingArr(ShippingIndex) = Shipping
    End Sub

    '**************************************************************************
	' Adds a defalt tax rule to the default tax table.
	'
	' Input:    TaxRule    Object: Class DefaultTaxRule
    '**************************************************************************
	Public Sub AddDefaultTaxRule(TaxRule)
        DefaultTaxRuleIndex = DefaultTaxRuleIndex + 1
		ReDim Preserve DefaultTaxRuleArr(DefaultTaxRuleIndex)
		Set DefaultTaxRuleArr(DefaultTaxRuleIndex) = TaxRule
	End Sub

    '**************************************************************************
	' Adds an alternate tax table to the alternate tax tables.
	'
	' Input:    TaxTable    Object: Class AlternateTaxTable
    '**************************************************************************
	Public Sub AddAlternateTaxTable(TaxTable)
		AlternateTaxTableIndex = AlternateTaxTableIndex + 1
		ReDim Preserve AlternateTaxTableArr(AlternateTaxTableIndex)
		Set AlternateTaxTableArr(AlternateTaxTableIndex) = TaxTable
	End Sub

    '**************************************************************************
	' Returns the <checkout-shopping-cart> XML in plain text format.
    '**************************************************************************
    Public Property Get GetXml()
		Dim Xml, Attr
		Dim i, Item, TaxRule, TaxTable, Shipping, State, ZipPattern

		Set Xml = new XmlBuilder

		Attr = Xml.Attribute("xmlns", SchemaUri)
		Xml.Push "checkout-shopping-cart", Attr
		Xml.Push "shopping-cart", ""

		If CartExpiration <> "" Then
			Xml.Push "cart-expiration", ""
			Xml.AddElement "good-until-date", CartExpiration, ""
			Xml.Pop "cart-expiration"
		End If
Xml.Push "items", ""
		For Each Item In ItemArr
			Xml.Push "item", ""
			Xml.AddElement "item-name", Item.Name, ""
			Xml.AddElement "item-description", Item.Description, ""
			Attr = Xml.Attribute("currency", MerchantCurrency)
			Xml.AddElement "unit-price", Item.UnitPrice, Attr
			Xml.AddElement "quantity", Item.Quantity, ""
			If Item.MerchantPrivateItemData <> "" Then
				Xml.AddXmlElement "merchant-private-item-data", Item.MerchantPrivateItemData, ""
			End If
			If Item.MerchantItemId <> "" Then
				Xml.AddElement "merchant-item-id", Item.MerchantItemId, ""
			End If
			If Item.TaxTableSelector <> "" Then
				Xml.AddElement "tax-table-selector", Item.TaxTableSelector, ""
			End If
			If Item.ItemWeight <> "" Then
				Attr = Xml.Attribute("unit", Item.ItemWeightUnit)
				Attr = Attr & " " & Xml.Attribute("value", Item.ItemWeight)
				Xml.EmptyElement "item-weight", Attr
			End If
			If Item.DigitalContent Then
			  Xml.Push "digital-content", ""
			  If Item.EmailDelivery Then
				Xml.AddElement "email-delivery", "true", ""
			  End If
			  If Item.DigitalDescription <> "" Then
				Xml.AddElement "description", Item.DigitalDescription, ""
			  End If
			  If Item.Key <> "" Then
				Xml.AddElement "key", Item.Key, ""
			  End If
			  If Item.Url <> "" Then
				Xml.AddElement "url", Item.Url, ""
			  End If
			  Xml.Pop "digital-content"
			End If
			Xml.Pop "item"
		Next
		Xml.Pop "items"


		If MerchantPrivateData <> "" Then
			Xml.AddXmlElement "merchant-private-data", MerchantPrivateData, ""
		End If

		Xml.Pop "shopping-cart"
		Xml.Push "checkout-flow-support", ""
		Xml.Push "merchant-checkout-flow-support", ""

		If EditCartUrl <> "" Then
			Xml.AddElement "edit-cart-url", EditCartUrl, ""
		End If

		If ContinueShoppingUrl <> "" Then
			Xml.AddElement "continue-shopping-url", ContinueShoppingUrl, ""
		End If

		If RequestBuyerPhoneNumber Then
			Xml.AddElement "request-buyer-phone-number", "true", ""
		End If

		If MerchantCalculationsUrl <> "" Then
			Xml.Push "merchant-calculations", ""
			Xml.AddElement "merchant-calculations-url", MerchantCalculationsUrl, ""
			If AcceptMerchantCoupons Then
				Xml.AddElement "accept-merchant-coupons", "true", ""
			End If
			If AcceptGiftCertificates Then
				Xml.AddElement "accept-gift-certificates", "true", ""
			End If
			Xml.Pop "merchant-calculations"
		End If

                If PlatformID <> "" Then
                         Xml.AddElement "platform-id", PlatformID, ""
               	End If
		If ShippingIndex >= 0 Then 
			Xml.Push "shipping-methods", ""
			
			For Each Shipping In ShippingArr
				If Shipping.ShippingType <> "carrier-calculated-shipping" Then
					Attr = Xml.Attribute("name", Shipping.Name)
					Xml.Push Shipping.ShippingType, Attr
					Attr = Xml.Attribute("currency", MerchantCurrency)
					Xml.AddElement "price", Shipping.Price, Attr
				
					If Shipping.ShippingType = "flat-rate-shipping" Or _
						Shipping.ShippingType = "merchant-calculated-shipping" Then

						If Shipping.HasShippingRestrictions Then
							Xml.Push "shipping-restrictions", ""

							Dim Restrictions
							Set Restrictions = Shipping.ShippingRestrictions

							If Restrictions.AllowUsPoBox Then
								Xml.AddElement "allow-us-po-box", "true", ""
							Else
								Xml.AddElement "allow-us-po-box", "false", ""
							End If

							Dim AllowedRestrictions, ExcludedRestrictions

							AllowedRestrictions = Restrictions.AllowedWorldArea Or _
													Restrictions.AllowedPostalAreaIndex >= 0 Or _
													Restrictions.AllowedCountryArea <> "" Or  _
													UBound(Restrictions.AllowedStates) >= 0 Or  _
													UBound(Restrictions.AllowedZipPatterns) >= 0

							ExcludedRestrictions = Restrictions.ExcludedPostalAreaIndex >= 0 Or _
													Restrictions.ExcludedCountryArea <> "" Or  _
													UBound(Restrictions.ExcludedStates) >= 0 Or  _
													UBound(Restrictions.ExcludedZipPatterns) >= 0

							If AllowedRestrictions Then
								Xml.Push "allowed-areas", ""
								If Restrictions.AllowedCountryArea <> "" Then
									Attr = Xml.Attribute("country-area", Restrictions.AllowedCountryArea)
									Xml.EmptyElement "us-country-area", Attr
								End If

								For Each State In Restrictions.AllowedStates
									Xml.Push "us-state-area", ""
									Xml.AddElement "state", State, ""
									Xml.Pop "us-state-area"
								Next

								For Each ZipPattern In Restrictions.AllowedZipPatterns
									Xml.Push "us-zip-area", ""
									Xml.AddElement "zip-pattern", ZipPattern, ""
									Xml.Pop "us-zip-area"
								Next

								If Restrictions.AllowedWorldArea Then
									Xml.EmptyElement "world-area", ""
								End If

								For i = 0 To Restrictions.AllowedPostalAreaIndex
									Xml.Push "postal-area", ""
									Xml.AddElement "country-code", Restrictions.AllowedCountryCodes(i), ""
									If Restrictions.AllowedPostalPatterns(i) <> "" Then
										Xml.AddElement "postal-code-pattern", Restrictions.AllowedPostalPatterns(i), ""
									End If
									Xml.Pop "postal-area"
								Next
		
								Xml.Pop "allowed-areas"
							End If

							If ExcludedRestrictions Then
								If Not AllowedRestrictions Then
									Xml.EmptyElement "allowed-areas", ""
								End If

								Xml.Push "excluded-areas", ""
								If Restrictions.ExcludedCountryArea <> "" Then
									Attr = Xml.Attribute("country-area", Restrictions.ExcludedCountryArea)
									Xml.EmptyElement "us-country-area", Attr
								End If

								For Each State In Restrictions.ExcludedStates
									Xml.Push "us-state-area", ""
									Xml.AddElement "state", State, ""
									Xml.Pop "us-state-area"
								Next

								For Each ZipPattern In Restrictions.ExcludedZipPatterns
									Xml.Push "us-zip-area", ""
									Xml.AddElement "zip-pattern", ZipPattern, ""
									Xml.Pop "us-zip-area"
								Next

								For i = 0 To Restrictions.ExcludedPostalAreaIndex
									Xml.Push "postal-area", ""
									Xml.AddElement "country-code", Restrictions.ExcludedCountryCodes(i), ""
									If Restrictions.ExcludedPostalPatterns(i) <> "" Then
										Xml.AddElement "postal-code-pattern", Restrictions.ExcludedPostalPatterns(i), ""
									End If
									Xml.Pop "postal-area"
								Next
								Xml.Pop "excluded-areas"
							End If
							Xml.Pop "shipping-restrictions"
						End If
						Set Restrictions = Nothing
					End If

					If Shipping.ShippingType = "merchant-calculated-shipping" Then
						If Shipping.HasAddressFilters Then
							Dim Filters
							Set Filters = Shipping.AddressFilters

							Xml.Push "address-filters", ""

							If Filters.AllowUsPoBox Then
								Xml.AddElement "allow-us-po-box", "true", ""
							Else
								Xml.AddElement "allow-us-po-box", "false", ""
							End If

							Dim AllowedFilters, ExcludedFilters
		
							AllowedFilters = Filters.AllowedWorldArea Or _
												Filters.AllowedPostalAreaIndex >= 0 Or _
												Filters.AllowedCountryArea <> "" Or  _
												UBound(Filters.AllowedStates) >= 0 Or  _
												UBound(Filters.AllowedZipPatterns) >= 0

							ExcludedFilters = Filters.ExcludedPostalAreaIndex >= 0 Or _
												Filters.ExcludedCountryArea <> "" Or  _
												UBound(Filters.ExcludedStates) >= 0 Or  _
												UBound(Filters.ExcludedZipPatterns) >= 0

							If AllowedFilters Then
								Xml.Push "allowed-areas", ""
								If Filters.AllowedCountryArea <> "" Then
									Attr = Xml.Attribute("country-area", Filters.AllowedCountryArea)
									Xml.EmptyElement "us-country-area", Attr
								End If

								For Each State In Filters.AllowedStates
									Xml.Push "us-state-area", ""
									Xml.AddElement "state", State, ""
									Xml.Pop "us-state-area"
								Next

								For Each ZipPattern In Filters.AllowedZipPatterns
									Xml.Push "us-zip-area", ""
									Xml.AddElement "zip-pattern", ZipPattern, ""
									Xml.Pop "us-zip-area"
								Next

								If Filters.AllowedWorldArea Then
									Xml.EmptyElement "world-area", ""
								End If

								For i = 0 To Filters.AllowedPostalAreaIndex
									Xml.Push "postal-area", ""
									Xml.AddElement "country-code", Filters.AllowedCountryCodes(i), ""
									If Filters.AllowedPostalPatterns(i) <> "" Then
										Xml.AddElement "postal-code-pattern", Filters.AllowedPostalPatterns(i), ""
									End If
									Xml.Pop "postal-area"
								Next
		
								Xml.Pop "allowed-areas"
							End If

							If ExcludedFilters Then
								If Not AllowedFilters Then
									Xml.EmptyElement "allowed-areas", ""
								End If

								Xml.Push "excluded-areas", ""
								If Filters.ExcludedCountryArea <> "" Then
									Attr = Xml.Attribute("country-area", Filters.ExcludedCountryArea)
									Xml.EmptyElement "us-country-area", Attr
								End If

								For Each State In Filters.ExcludedStates
									Xml.Push "us-state-area", ""
									Xml.AddElement "state", State, ""
									Xml.Pop "us-state-area"
								Next

								For Each ZipPattern In Filters.ExcludedZipPatterns
									Xml.Push "us-zip-area", ""
									Xml.AddElement "zip-pattern", ZipPattern, ""
									Xml.Pop "us-zip-area"
								Next

								For i = 0 To Filters.ExcludedPostalAreaIndex
									Xml.Push "postal-area", ""
									Xml.AddElement "country-code", Filters.ExcludedCountryCodes(i), ""
									If Filters.ExcludedPostalPatterns(i) <> "" Then
										Xml.AddElement "postal-code-pattern", Filters.ExcludedPostalPatterns(i), ""
									End If
									Xml.Pop "postal-area"
								Next

								Xml.Pop "excluded-areas"
							End If
							Xml.Pop "address-filters"
							Set Filters = Nothing
						End If
					End If
					Xml.Pop Shipping.ShippingType
				Else
					Xml.Push "carrier-calculated-shipping", ""
					If Shipping.ShippingOptionIndex >= 0 Then
						Xml.Push "carrier-calculated-shipping-options", ""
						Dim ShippingOption
						For Each ShippingOption In Shipping.ShippingOptions
							Xml.Push "carrier-calculated-shipping-option", ""
							Attr = Xml.Attribute("currency", MerchantCurrency)
							Xml.AddElement "price", ShippingOption.Price, Attr
							Xml.AddElement "shipping-company", ShippingOption.ShippingCompany, ""
							Xml.AddElement "shipping-type", ShippingOption.ShippingType, ""
							If ShippingOption.CarrierPickup <> "" Then
								Xml.AddElement "carrier-pickup", ShippingOption.CarrierPickup, ""
							End If
							If ShippingOption.AdditionalFixedCharge <> "" Then
								Attr = Xml.Attribute("currency", MerchantCurrency)
								Xml.AddElement "additional-fixed-charge", ShippingOption.AdditionalFixedCharge, Attr
							End If
							If ShippingOption.AdditionalVariableChargePercent <> "" Then
								Xml.AddElement "additional-variable-charge-percent", ShippingOption.AdditionalVariableChargePercent, ""
							End If
							Xml.Pop "carrier-calculated-shipping-option"
						Next
						Set ShippingOption = Nothing
						Xml.Pop "carrier-calculated-shipping-options"
					End If
					If Shipping.ShippingPackageIndex >= 0 Then
						Xml.Push "shipping-packages", ""
						Dim ShippingPackage
						For Each ShippingPackage In Shipping.ShippingPackages
							Xml.Push "shipping-package", ""
							Attr = Xml.Attribute("unit", ShippingPackage.HeightUnit)
							Attr = Attr & " " & Xml.Attribute("value", ShippingPackage.Height)
							Xml.EmptyElement "height", Attr
							Attr = Xml.Attribute("unit", ShippingPackage.LengthUnit)
							Attr = Attr & " " & Xml.Attribute("value", ShippingPackage.Length)
							Xml.EmptyElement "length", Attr
							Attr = Xml.Attribute("unit", ShippingPackage.WidthUnit)
							Attr = Attr & " " & Xml.Attribute("value", ShippingPackage.Width)
							Xml.EmptyElement "width", Attr
							If ShippingPackage.Packaging <> "" Then
								Xml.AddElement "packaging", ShippingPackage.Packaging, ""
							End If
							If ShippingPackage.DeliveryAddressCategory <> "" Then
								Xml.AddElement "delivery-address-category", ShippingPackage.DeliveryAddressCategory, ""
							End If
							Dim ShipFromAddress
							Set ShipFromAddress = ShippingPackage.ShipFromAddress
							Attr = Xml.Attribute("id", ShipFromAddress.ID)
							Xml.Push "ship-from", Attr
							Xml.AddElement "city", ShipFromAddress.City, ""
							Xml.AddElement "region", ShipFromAddress.Region, ""
							Xml.AddElement "country-code", ShipFromAddress.CountryCode, ""
							Xml.AddElement "postal-code", ShipFromAddress.PostalCode, ""
							Xml.Pop "ship-from"
							Set ShipFromAddress = Nothing
							Xml.Pop "shipping-package"
						Next
						Set ShippingPackage = Nothing
						Xml.Pop "shipping-packages"
					End If
					Xml.Pop "carrier-calculated-shipping"
				End If
			Next
		Xml.Pop "shipping-methods"
		End If



		If DefaultTaxRuleIndex >= 0 Or AlternateTaxTableIndex >= 0 Then
			If MerchantCalculatedTax Then
				Attr = Xml.Attribute("merchant-calculated", "true")
				Xml.Push "tax-tables", Attr
			Else
				Xml.Push "tax-tables", ""
			End If

			If DefaultTaxRuleIndex >= 0 Then
				Xml.Push "default-tax-table", ""
				Xml.Push "tax-rules", ""
				For Each TaxRule In DefaultTaxRuleArr
					If TaxRule.CountryArea <> "" Then
						Xml.Push "default-tax-rule", ""
						If TaxRule.ShippingTaxed Then
							Xml.AddElement "shipping-taxed", "true", ""
						End If
						Xml.AddElement "rate", TaxRule.Rate, ""
						Xml.Push "tax-area", ""
						Attr = Xml.Attribute("country-area", TaxRule.CountryArea)
						Xml.EmptyElement "us-country-area", Attr
						Xml.Pop "tax-area"
						Xml.Pop "default-tax-rule"
					End If

					For Each State In TaxRule.States
						Xml.Push "default-tax-rule", ""
						If TaxRule.ShippingTaxed Then
							Xml.AddElement "shipping-taxed", "true", ""
						End If
						Xml.AddElement "rate", TaxRule.Rate, ""
						Xml.Push "tax-area", ""
						Xml.Push "us-state-area", ""
						Xml.AddElement "state", State, ""
						Xml.Pop "us-state-area"
						Xml.Pop "tax-area"
						Xml.Pop "default-tax-rule"
					Next

					For Each ZipPattern In TaxRule.ZipPatterns
						Xml.Push "default-tax-rule", ""
						If TaxRule.ShippingTaxed Then
							Xml.AddElement "shipping-taxed", "true", ""
						End If
						Xml.AddElement "rate", TaxRule.Rate, ""
						Xml.Push "tax-area", ""
						Xml.Push "us-zip-area", ""
						Xml.AddElement "zip-pattern", ZipPattern, ""
						Xml.Pop "us-zip-area"
						Xml.Pop "tax-area"
						Xml.Pop "default-tax-rule"
					Next
				Next
				Xml.Pop "tax-rules"
				Xml.Pop "default-tax-table"
			End If

			If AlternateTaxTableIndex >= 0 Then
				Xml.Push "alternate-tax-tables", ""

				For Each TaxTable In AlternateTaxTableArr
					Attr = Xml.Attribute("name", TaxTable.Name)
					If TaxTable.Standalone Then
						Attr = Attr & " " & Xml.Attribute("standalone", "true")
					End If
					Xml.Push "alternate-tax-table", Attr
					Xml.Push "alternate-tax-rules", ""

					For Each TaxRule In TaxTable.AlternateTaxRuleArr
						If TaxRule.CountryArea <> "" Then
							Xml.Push "alternate-tax-rule", ""
							Xml.AddElement "rate", TaxRule.Rate, ""
							Xml.Push "tax-area", ""
							Attr = Xml.Attribute("country-area", TaxRule.CountryArea)
							Xml.EmptyElement "us-country-area", Attr
							Xml.Pop "tax-area"
							Xml.Pop "alternate-tax-rule"
						End If

						For Each State In TaxRule.States
							Xml.Push "alternate-tax-rule", ""
							Xml.AddElement "rate", TaxRule.Rate, ""
							Xml.Push "tax-area", ""
							Xml.Push "us-state-area", ""
							Xml.AddElement "state", State, ""
							Xml.Pop "us-state-area"
							Xml.Pop "tax-area"
							Xml.Pop "alternate-tax-rule"
						Next

						For Each ZipPattern In TaxRule.ZipPatterns
							Xml.Push "alternate-tax-rule", ""
							Xml.AddElement "rate", TaxRule.Rate, ""
							Xml.Push "tax-area", ""
							Xml.Push "us-zip-area", ""
							Xml.AddElement "zip-pattern", ZipPattern, ""
							Xml.Pop "us-zip-area"
							Xml.Pop "tax-area"
							Xml.Pop "alternate-tax-rule"
						Next
					Next
					Xml.Pop "alternate-tax-rules"
					Xml.Pop "alternate-tax-table"
				Next
				Xml.Pop "alternate-tax-tables"
			End If
			Xml.Pop "tax-tables"
		End If

		Xml.Pop "merchant-checkout-flow-support"
		Xml.Pop "checkout-flow-support"

		If RequestInitialAuthDetails Then
			Xml.Push "order-processing-support", ""
			Xml.AddElement "request-initial-auth-details", "true", ""
			Xml.Pop "order-processing-support"
		End If

		Xml.Pop "checkout-shopping-cart"

		GetXml = Xml.GetXml
		Set Xml = Nothing
    End Property

    '**************************************************************************
	' Posts the cart to Google Checkout and redirects the user to 
	'     the Google Checkout page accordingly.
    '**************************************************************************
	Public Sub PostCartToGoogle()
		Dim ResponseXml
		ResponseXml = SendRequest(GetXml, PostUrl)
		ProcessSynchronousResponse(ResponseXml)
	End Sub

    '**************************************************************************
	' Diagnoses the shopping cart XML by posting to the Diagnose URL.
	' It displays the response in the browser.
    '**************************************************************************
	Public Sub DiagnoseXml()
		Dim ResponseXml
		ResponseXml = SendRequest(GetXml, DiagnoseUrl)
		ProcessSynchronousResponse(ResponseXml)
	End Sub

    '**************************************************************************
	' Processes the synchronous response received from Google Checkout 
	'     when a cart is posted.
    '**************************************************************************
	Public Sub ProcessSynchronousResponse(ResponseXml)
		Dim ResponseNode, RootTag
		Set ResponseNode = GetRootNode(ResponseXml)
		RootTag = ResponseNode.Tagname

		Select Case RootTag
			Case "request-received"
				DisplayResponse ResponseXml
			Case "error"
				DisplayResponse ResponseXml
			Case "diagnosis"
				DisplayResponse ResponseXml
			Case "checkout-redirect"
				ProcessCheckoutRedirect ResponseXml
			Case Else
		End Select 
	End Sub

    '**************************************************************************
	' Displays the response in the browser.
    '**************************************************************************
	Public Sub DisplayResponse(ResponseXml)
		Response.Write "<pre>"
		Response.Write Server.HTMLEncode(ResponseXml)
		Response.Write "</pre>"
	End Sub

    '**************************************************************************
	' Redirects the user to the Google Checkout page when a cart is 
	'     successfully posted.
    '**************************************************************************
	Public Sub ProcessCheckoutRedirect(ResponseXml)
		Dim ResponseNode, RedirectUrl
		Set ResponseNode = GetRootNode(ResponseXml)
		RedirectUrl = GetElementText(ResponseNode, "redirect-url")
		Response.Redirect RedirectUrl
		Set ResponseNode = Nothing
	End Sub

	Private Sub Class_Terminate()
		Erase ItemArr
		Erase ShippingArr
		Erase DefaultTaxRuleArr
		Erase AlternateTaxTableArr
	End Sub
End Class



'******************************************************************************
' Class Item is the object representation of an <item>.
'******************************************************************************
Class Item
    Public Name
    Public Description
    Public Quantity
    Public UnitPrice
    Public MerchantItemId
    Public TaxTableSelector
    Public MerchantPrivateItemData
    Public ItemWeight
	Public ItemWeightUnit
    Public DigitalContent
    Public EmailDelivery
    Public DigitalDescription
    Public Key
    Public Url

    Private Sub Class_Initialize()
		ItemWeightUnit = "LB"
    End Sub

    '**************************************************************************
	' Parses through an <item> node and makes the Item object ready to be used.
    '**************************************************************************
	Public Sub ParseItem(ItemNode)
		Name = GetElementText(ItemNode, "item-name")
		Description = GetElementText(ItemNode, "item-description")
		Quantity = GetElementText(ItemNode, "quantity")
		ItemWeight=GetElementText(ItemNode, "item-weight")
		UnitPrice = GetElementText(ItemNode, "unit-price")
		MerchantItemId = GetElementText(ItemNode, "merchant-item-id")
		TaxTableSelector = GetElementText(ItemNode, "tax-table-selector")

		Dim MerchantPrivateItemDataNL
		Set MerchantPrivateItemDataNL = ItemNode.getElementsByTagname("merchant-private-item-data")
		If MerchantPrivateItemDataNL.Length > 0 Then Set MerchantPrivateItemData = MerchantPrivateItemDataNL(0)
		Set MerchantPrivateItemDataNL = Nothing

		Dim DigitalContentNL
		Set DigitalContentNL = ResponseNode.getElementsByTagname("digital-content")
		If DigitalContentNL.Length > 0 Then 
			Dim DigitalContentNode
			Set DigitalContentNode = DigitalContentNL(0)
			DigitalDescription = GetElementText(DigitalContentNode, "description")
			EmailDelivery = GetElementText(DigitalContentNode, "email-delivery")
			Key = GetElementText(DigitalContentNode, "key")
			Url = GetElementText(DigitalContentNode, "url")
			Set DigitalContentNode = Nothing
		End If
		Set DigitalContentNL = Nothing
	End Sub
End Class


'******************************************************************************
' The DisplayButton function displays a Google Checkout button.
' When a buyer clicks on the button, it will invoke the CartProcessingPage
'
' Input:    ButtonSize    "LARGE" OR "MEDIUM" OR "SMALL"
'           Enabled       true - displays a Google Checkout button
'                         false - displays a disabled Google Checkout button
'******************************************************************************
Sub DisplayButton(ButtonSize, Enabled,button_style) 
	Dim ButtonWidth, ButtonHeight, ButtonStyle, ButtonVariant, ButtonLocale, ButtonSrc, ButtonGif

	ButtonSize = UCase(ButtonSize)
	Select Case ButtonSize
		Case "LARGE"
			ButtonWidth = "180"
			ButtonHeight = "46"
		Case "MEDIUM"
			ButtonWidth = "168"
			ButtonHeight = "44"
		Case "SMALL"
			ButtonWidth = "160"
			ButtonHeight = "43"
		Case Else
	End Select 

	If Enabled = True Then
		ButtonVariant = "text"
	ElseIf Enabled = False Then
		ButtonVariant = "disabled"
	End If

	EnvType = UCase(EnvType)
	If EnvType = "SANDBOX" Then
		ButtonGif = "https://sandbox.google.com/checkout/buttons/checkout.gif"
	ElseIf EnvType = "PRODUCTION" Then
		ButtonGif = "https://checkout.google.com/buttons/checkout.gif"
	End If
        if button_style="white" then
	ButtonStyle = "white"
	else
	ButtonStyle = "trans"
        end if

        ButtonLocale = "en_US"
	ButtonSrc = ButtonGif & _
				"?merchant_id=" & MerchantId & _
				"&w=" & ButtonWidth & _
				"&h=" & ButtonHeight & _
				"&style=" & ButtonStyle & _
				"&variant=" & ButtonVariant & _
				"&loc=" & ButtonLocale

	If Enabled = True Then
		Response.Write VbCrLf & "<form method=""POST"" action=""" & CartProcessingPage & """>" & VbCrLf
		Response.Write "  <input type=""image"" name=""Checkout"" alt=""Checkout"" src=""" & ButtonSrc & """ height=""" & ButtonHeight & """ width=""" & ButtonWidth & """>" & VbCrLf
		Response.Write "</form>" & VbCrLf
	ElseIf Enabled = False Then
		Response.Write VbCrLf & "<img alt=""Checkout"" src=""" & ButtonSrc & """ height=""" & ButtonHeight & """ width=""" & ButtonWidth & """ />" & VbCrLf
	End If
End Sub


%>
