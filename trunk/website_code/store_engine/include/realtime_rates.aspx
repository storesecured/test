<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="dotnetSHIP" %>
<%@ Page Language="VB" Debug="true" %>
<script language="VB" runat="server">



Dim objShip,objShip2 As dotnetSHIP.Ship 

Sub Page_Load(Src As Object, E As EventArgs)

Dim Service,sFinalString as string
Dim sShippingOptions as string
Dim iMult as double
Dim iLeftover, iMaxWeight as double
Dim iConwayWeight as integer, iUseUPS as integer, iUseUSPS as integer
Dim rate,rate2 as rate
Dim sCompanies,sName,sCharge,sNone,sFinal,sUrl,sRate as String
Dim is_Residential as String,sDestSt as string

on error resume next
      
        objShip = New dotnetSHIP.Ship()
            is_Residential = Request.Params("Res")
            iMaxWeight = Request.Params("Max_Weight")
            If Request.Params("sWt") > iMaxWeight Then
                objShip.Weight = iMaxWeight
                iLeftover = Request.Params("sWt") Mod iMaxWeight
                iMult = (Request.Params("sWt") - iLeftover) / iMaxWeight

            Else
                objShip.Weight = Request.Params("sWt")
                iMult = 1
                iLeftover = 0
            End If
            objShip.OrigCountry = Request.Params("OrgCty")
            If Request.Params("OrgCty") = "ZW" Then
                sFinalString = "Rates_" & Request.Params("Res") & "=|&"
                Response.Write(Replace((sFinalString), "&&", "&"))
            End If
            objShip.DestCountry = Request.Params("DestCty")
            If Request.Params("DestCty") = "US" Then
                objShip.DestZipPostal = Left(Request.Params("DestZip"), 5)
            Else
                objShip.DestZipPostal = Request.Params("DestZip")
            End If
            If Request.Params("OrgCty") = "US" Then
                objShip.OrigZipPostal = Left(Request.Params("OrgZip"), 5)
            Else
                objShip.OrigZipPostal = Request.Params("OrgZip")
            End If

            objShip.OrigStateProvince = Request.Params("OrgST")
            sDestSt=Request.Params("DestST")
            iUseUPS=Request.Params("USE_UPS")
            iUseUSPS=Request.Params("USE_USPS")
            
		  if iUseUSPS and (sDestSt="AA" or sDestSt="AE" or sDestSt="AP") then
            	iUseUPS=0
		  else
            	objShip.DestStateProvince = sDestSt
		  end if

            If iUseUPS Then
                objShip.UPSLogin = Request.Params("UPS_Info") ' Please enter your login data here
                'UPSPickup: Pickup options for UPS including: 1) "customer counter" 2)"letter center" 3)"on call air" 4) "one time pickup" 5)"daily pickup" 6)"authorized shipping outlet" 7)"air service center". The default value is set to "customer counter."
                objShip.UPSPickup=Request.Params("UPS_Pickup")
                'UPSPack: Packing options for UPS including: 1) "Shipper Supplied Packaging" 2) "UPS Letter Envelope" or "UPS Letter" 3) "Your Packaging" or "Package" 4) "UPS Tube" 5) "UPS Pak" 6) "UPS Express Box" 7) "International UPS 25KG Box" or "UPS 25KG Box" 8) "International UPS 10KG Box" or "UPS 10KG Box" The default value is set to "Your Packaging."
                objShip.UPSPack=Request.Params("UPS_Pack")

                If sShippingOptions <> "" Then
                    sShippingOptions += ","
                End If
                sShippingOptions += "UPS"
            End If
            If iUseUSPS Then
                objShip.USPSLogin = "876EASYS2859,610EY45TV868" ' Please enter your login data here
                If sShippingOptions <> "" Then
                    sShippingOptions += ","
                End If
                sShippingOptions += "USPS"
            End If
            If Request.Params("USE_DHL") Then
                'objShip.DHLServices = request.params("DHL_Service")
                'if sShippingOptions<>"" then
                '   sShippingOptions += ","
                'end if
                'sShippingOptions += "DHL"
            End If
            If Request.Params("USE_FEDEX") and 1=0 Then
                'FedExPack: Packing Options for FedEx include: 1) "FedEx Express Envelope" or "FedEx Letter" , 2) "FedEx Express Pak" or "FedEx Pak" 3) "FedEx Express Box" or "FedEx Box" , 4) "FedEx Express Tube" or "FedEx Tube" 5) "My Packaging" or "Your Packaging". The default for this property is "My Packaging"
                'objShip.FedExURL="https://gatewaybeta.fedex.com/GatewayDC"
                'objShip.FedExLogin = "510064429,1156376"  ' "Account number, meter number"
                'objShip.FedExLogin = "296454281,1263356"  ' "Account number, meter number"
                'objShip.FedExLogin = "239367186,1828724"  ' "Account number, meter number"

                'objShip.FedExPack = Request.Params("FedEx_Pack")
                'If Request.Params("Fedex_Ground") Then
                '    objShip.FedExGround = True
                'Else
                '    objShip.FedExGround = False
                'End If
                'If sShippingOptions <> "" Then
                '    sShippingOptions += ","
                'End If
                'sShippingOptions += "Fedex"
            End If
            If Request.Params("USE_AIRBORNE") Then
                If sShippingOptions <> "" Then
                    sShippingOptions += ","
                End If
                sShippingOptions += "Airborne"
            End If
            If Request.Params("USE_CANADA") Then
                objShip.CanadaPostLogin = Request.Params("Canada_Info")
                If sShippingOptions <> "" Then
                    sShippingOptions += ","
                End If
                sShippingOptions += "CanadaPost"
                objShip.OrigStateProvince = Request.Params("OrgST")
                objShip.DestStateProvince = sDestSt
            End If
            If Request.Params("USE_CONWAY") Then

                objShip.ConWayLogin = Request.Params("Conway_Info")
                objShip.ConWay.SingleShipment = True
                If sShippingOptions <> "" Then
                    sShippingOptions += ","
                End If
                sShippingOptions += "Conway"
                iConwayWeight = FormatNumber(iMaxWeight, 0)
                objShip.Weight = FormatNumber(Request.Params("sWt"), 0)
            End If

            objShip.Timeout = 40

            objShip.Sort = "price"
            If is_Residential <> "False" Then
                objShip.Residential = True
            Else
                objShip.Residential = False
            End If

            objShip.Rate(sShippingOptions)

            If iLeftover <= 0 Then
                sFinal = ""
                For Each rate In objShip.Rates
                If rate.Charge <> 0 Then
                    sRate = rate.Charge * iMult
                    sFinal += rate.Company + "," + rate.Name + "," + sRate + "|"
                End If
                Next
                
            Else
                objShip2 = New dotnetSHIP.Ship()

                objShip2.Weight = iLeftover
                objShip2.OrigCountry = Request.Params("OrgCty")
                objShip2.DestCountry = Request.Params("DestCty")
                objShip2.DestZipPostal = Request.Params("DestZip")
                objShip2.OrigZipPostal = Request.Params("OrgZip")
                objShip2.OrigStateProvince = Request.Params("OrgST")
                objShip2.DestStateProvince = sDestSt
                If iUseUPS Then
                    objShip2.UPSLogin = Request.Params("UPS_Info") ' Please enter your login data here
                    'UPSPickup: Pickup options for UPS including: 1) "customer counter" 2)"letter center" 3)"on call air" 4) "one time pickup" 5)"daily pickup" 6)"authorized shipping outlet" 7)"air service center". The default value is set to "customer counter."
                    objShip2.UPSPickup=Request.Params("UPS_Pickup")
                    'UPSPack: Packing options for UPS including: 1) "Shipper Supplied Packaging" 2) "UPS Letter Envelope" or "UPS Letter" 3) "Your Packaging" or "Package" 4) "UPS Tube" 5) "UPS Pak" 6) "UPS Express Box" 7) "International UPS 25KG Box" or "UPS 25KG Box" 8) "International UPS 10KG Box" or "UPS 10KG Box" The default value is set to "Your Packaging."
                    objShip2.UPSPack=Request.Params("UPS_Pack")

                    If sShippingOptions <> "" Then
                        sShippingOptions += ","
                    End If
                    sShippingOptions += "UPS"
                End If
                If iUseUSPS Then
                    objShip2.USPSLogin = "876EASYS2859,610EY45TV868" ' Please enter your login data here
                    If sShippingOptions <> "" Then
                        sShippingOptions += ","
                    End If
                    sShippingOptions += "USPS"
                End If
                If Request.Params("USE_DHL") Then
                    'objShip2.DHLServices = request.Params("DHL_Service")
                    If sShippingOptions <> "" Then
                        sShippingOptions += ","
                    End If
                    sShippingOptions += "DHL"
                End If
                If Request.Params("USE_FEDEX") and 1=0 Then
                    'FedExPack: Packing Options for FedEx include: 1) "FedEx Express Envelope" or "FedEx Letter" , 2) "FedEx Express Pak" or "FedEx Pak" 3) "FedEx Express Box" or "FedEx Box" , 4) "FedEx Express Tube" or "FedEx Tube" 5) "My Packaging" or "Your Packaging". The default for this property is "My Packaging"
                    'objShip2.FedExURL="https://gatewaybeta.fedex.com/GatewayDC"
                    'objShip2.FedExLogin = "510064429,1156376"  ' "Account number, meter number"
                    'objShip2.FedExLogin = "296454281,1263356"  ' "Account number, meter number"
                    'objShip2.FedExPack = Request.Params("FedEx_Pack")
                    'If Request.Params("Fedex_Ground") Then
                    '    objShip2.FedExGround = True
                    'Else
                    '    objShip2.FedExGround = False
                    'End If
                    'If sShippingOptions <> "" Then
                    '    sShippingOptions += ","
                    'End If
                    'sShippingOptions += "Fedex"
                End If
                If Request.Params("USE_AIRBORNE") Then
                    If sShippingOptions <> "" Then
                        sShippingOptions += ","
                    End If
                    sShippingOptions += "Airborne"
                End If
                If Request.Params("USE_CANADA") Then
                    objShip2.CanadaPostLogin = Request.Params("Canada_Info")
                    If sShippingOptions <> "" Then
                        sShippingOptions += ","
                    End If
                    sShippingOptions += "CanadaPost"
                End If
                If Request.Params("USE_CONWAY") Then
                    objShip2.ConWayLogin = Request.Params("Conway_Info")
                    objShip2.ConWay.SingleShipment = True
                End If
		
                objShip2.Timeout = 40

                objShip2.Sort = "price"
                If is_Residential <> "False" Then
                    objShip2.Residential = True
                Else
                    objShip2.Residential = False
                End If

                objShip2.Rate(sShippingOptions)

                If (objShip2.Error) <> "" And Request.Params("OrgCty") <> "ZW" Then
                    Response.Write("Rates=&Error=" & Server.UrlEncode(objShip2.Error))
                Else
                    sFinal = ""
                    For Each rate2 In objShip2.Rates
                        For Each rate In objShip.Rates
                            If (rate2.Company = rate.Company) And (rate2.Name = rate.Name) Then
                                sRate = (rate.Charge * iMult) + rate2.Charge
                                sFinal += rate.Company + "," + rate.Name + "," + sRate + "|"
                            End If
                        Next
                    Next
                End If

            End If

        Response.Write("Rates=" + sFinal)
        If objShip.Error <> "" Then
            Response.Write("&Error=" + objShip.Error)
        End If
end sub
</script>


