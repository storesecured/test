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

' Class FlatRateShipping is the object representation of a
'     <flat-rate-shipping>.
'******************************************************************************
Class FlatRateShipping
	Public Name
	Public Price
	Public ShippingType
	Public ShippingRestrictions
	Public HasShippingRestrictions

    Private Sub Class_Initialize()
		ShippingType = "flat-rate-shipping"
		Set ShippingRestrictions = Nothing
		HasShippingRestrictions = false
    End Sub

	Public Sub AddShippingRestrictions(Restrictions)
		Set ShippingRestrictions = Restrictions
		HasShippingRestrictions = true
	End Sub

	Private Sub Class_Terminate()
		Set ShippingRestrictions = Nothing
	End Sub
End Class


'******************************************************************************
' Class MerchantCalculatedShipping is the object representation of a
'     <merchant-calculated-shipping>.
'******************************************************************************
Class MerchantCalculatedShipping
	Public Name
	Public Price
	Public ShippingType
	Public ShippingRestrictions
	Public AddressFilters
	Public HasShippingRestrictions
	Public HasAddressFilters

    Private Sub Class_Initialize()
		ShippingType = "merchant-calculated-shipping"

		Set ShippingRestrictions = Nothing
		Set AddressFilters = Nothing

		HasShippingRestrictions = false
		HasAddressFilters = false
    End Sub

	Public Sub AddShippingRestrictions(Restrictions)
		Set ShippingRestrictions = Restrictions
		HasShippingRestrictions = true
	End Sub

	Public Sub AddAddressFilters(Filters)
		Set AddressFilters = Filters
		HasAddressFilters = true
	End Sub

	Private Sub Class_Terminate()
		Set ShippingRestrictions = Nothing
		Set AddressFilters = Nothing
	End Sub
End Class

'******************************************************************************
' Class ShippingFilters is the object representation of a
'     <shipping-restricitons> or <address-filters>.
'******************************************************************************
Class ShippingFilters
	Public AllowUsPoBox

	Public AllowedWorldArea
	Public AllowedCountryCodes()
	Public AllowedPostalPatterns()
	Public AllowedCountryArea
	Public AllowedStates
	Public AllowedZipPatterns

	Public ExcludedCountryCodes()
	Public ExcludedPostalPatterns()
	Public ExcludedCountryArea
	Public ExcludedStates
	Public ExcludedZipPatterns

	Public AllowedPostalAreaIndex
	Public ExcludedPostalAreaIndex

    Private Sub Class_Initialize()
		AllowUsPoBox = true

		AllowedWorldArea = false
		AllowedCountryArea = ""
		AllowedStates = Array()
		AllowedZipPatterns = Array()
		
		ExcludedCountryArea = ""
		ExcludedStates = Array()
		ExcludedZipPatterns = Array()

		AllowedPostalAreaIndex = -1
        ExcludedPostalAreaIndex = -1
    End Sub

    '**************************************************************************
	' Adds an allowed postal-area to the shipping restrictions.
	'
	' Input:    CountryCode      Two letter country code. Ex: "US", "GB", etc.
	'           PostalPattern    Postal code pattern. Ex: "9404*", "SW*"
    '**************************************************************************
	Public Sub AddAllowedPostalArea(CountryCode, PostalPattern)
        AllowedPostalAreaIndex = AllowedPostalAreaIndex + 1
		ReDim Preserve AllowedCountryCodes(AllowedPostalAreaIndex)
		ReDim Preserve AllowedPostalPatterns(AllowedPostalAreaIndex)

        AllowedCountryCodes(AllowedPostalAreaIndex) = CountryCode
        AllowedPostalPatterns(AllowedPostalAreaIndex) = PostalPattern
	End Sub

    '**************************************************************************
	' Adds an excluded postal-area to the shipping restrictions.
	'
	' Input:    CountryCode      Two letter country code. Ex: "US", "GB", etc.
	'           PostalPattern    Postal code pattern. Ex: "9404*", "SW*"
    '**************************************************************************
	Public Sub AddExcludedPostalArea(CountryCode, PostalPattern)
        ExcludedPostalAreaIndex = ExcludedPostalAreaIndex + 1
		ReDim Preserve ExcludedCountryCodes(ExcludedPostalAreaIndex)
		ReDim Preserve ExcludedPostalPatterns(ExcludedPostalAreaIndex)

		ExcludedCountryCodes(ExcludedPostalAreaIndex) = CountryCode
		ExcludedPostalPatterns(ExcludedPostalAreaIndex) = PostalPattern
	End Sub

	Private Sub Class_Terminate()
		Erase AllowedCountryCodes
		Erase AllowedPostalPatterns
		Erase ExcludedCountryCodes
		Erase ExcludedPostalPatterns
	End Sub
End Class
'******************************************************************************
' Class CarrierCalculatedShipping is the object representation of a
'     <carrier-calculated-shipping>.
'******************************************************************************
Class CarrierCalculatedShipping
	Public ShippingType
	Public ShippingOptions()
	Public ShippingPackages()
	Public ShippingOptionIndex
	Public ShippingPackageIndex

    Private Sub Class_Initialize()
		ShippingType = "carrier-calculated-shipping"

		ShippingOptionIndex = -1
		ShippingPackageIndex = -1
    End Sub

	Public Sub AddShippingOption(ShippingOption)
		ShippingOptionIndex = ShippingOptionIndex + 1
		ReDim Preserve ShippingOptions(ShippingOptionIndex)
		Set ShippingOptions(ShippingOptionIndex) = ShippingOption
	End Sub

	Public Sub AddShippingPackage(ShippingPackage)
		ShippingPackageIndex = ShippingPackageIndex + 1
		ReDim Preserve ShippingPackages(ShippingPackageIndex)
		Set ShippingPackages(ShippingPackageIndex) = ShippingPackage
	End Sub

	Private Sub Class_Terminate()
		Erase ShippingOptions
		Erase ShippingPackages
	End Sub
End Class

'******************************************************************************
' Class ShippingOption is the object representation of a 
'     <carrier-calculated-shipping-option>.
'******************************************************************************
Class ShippingOption
	Public Price
	Public ShippingCompany
	Public CarrierPickup
	Public ShippingType
	Public AdditionalFixedCharge
	Public AdditionalVariableChargePercent
End Class

'******************************************************************************
' Class ShippingPackage is the object representation of a 
'     <shipping-package>.
'******************************************************************************
Class ShippingPackage

	Public Height
	Public HeightUnit
	Public Length
	Public LengthUnit
	Public Width
	Public WidthUnit
	Public Packaging
	Public DeliveryAddressCategory
	Public ShipFromAddress

    Private Sub Class_Initialize()
		HeightUnit = "IN"
		LengthUnit = "IN"
		WidthUnit = "IN"
    End Sub

	Public Sub AddShipFrom(Address)
		Set ShipFromAddress = Address
	End Sub

	Private Sub Class_Terminate()
		Set ShipFromAddress = Nothing
	End Sub
End Class

'******************************************************************************
' Class Pickup is the object representation of a <pickup>.
'******************************************************************************
Class PickupShipping
	Public Name
	Public Price
	Public ShippingType

    Private Sub Class_Initialize()
		ShippingType = "pickup"
    End Sub
End Class

%>
