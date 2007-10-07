<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="dotnetSHIP" %>
<%@ Page Language="VB" Debug="true" %>
<script language="VB" runat="server">



Dim objShip,objShip2 As dotnetSHIP.Ship 

Sub Page_Load(Src As Object, E As EventArgs)

Dim Service,sFinalString as string
Dim sArray(7) as string
Dim sShippingOptions as string
Dim iMult as double
Dim iLeftover, iMaxWeight as double
Dim iConwayWeight,iUPS as integer
Dim rate,rate2 as rate
Dim sCompanies,sName,sCharge,sNone,sFinal,sUrl,sRate as String
Dim is_Residential as String

on error resume next
For Each Service In split(Request.Params("Info"),"|")
      
	objShip = new dotnetSHIP.Ship()

	sArray=split(Service,",")
	
	if Service <> "" then
	  is_Residential=sArray(8)
	  iMaxWeight=request.params("Max_Weight")
		if sArray(6)>iMaxWeight then
			objShip.Weight= iMaxWeight
			iLeftover=sArray(6) mod iMaxWeight
			iMult=(sArray(6)-iLeftover)/iMaxWeight
			
		else
			objShip.Weight= sArray(6)
			iMult=1
			iLeftover=0
		end if
		objShip.OrigCountry=sArray(2)
      if sArray(2)="ZW" then
         sFinalString += "Rates_"&sArray(7)&"=|&"
         response.redirect (replace((Request.Params("Site")+"before_payment.asp?"+sFinalString+request.params("Query")),"&&","&"))
		end if
		objShip.DestCountry= sArray(5)
		if sArray(5) = "US" then
		   objShip.DestZipPostal= left(sArray(3),5)
		else
         objShip.DestZipPostal= sArray(3)
		end if
      if sArray(2) = "US" then
		   objShip.OrigZipPostal= left(sArray(0),5)
		else
         objShip.OrigZipPostal= sArray(0)
		end if

		'objShip.OrigZipPostal= sArray(0)
		'objShip.OrigStateProvince = sArray(1)
		'objShip.DestStateProvince = sArray(4)

                'response.end
                if Request.Params("USE_CONWAY")<>1 and Request.Params("USE_CANADA")<>1 and Request.Params("USE_AIRBORNE")<>1 and Request.Params("USE_DHL")<>1 and Request.Params("USE_UPS")<>1 and Request.Params("USE_UPS")<>1 then
      	           response.end
                   response.redirect (Request.Params("Site")+"before_payment.asp?Error="+server.urlencode("No shipping providers were selected."))

              end if

		if Request.Params("USE_UPS")=1 or iUPS=1 then
			objShip.UPSLogin = Request.Params("UPS_Info") ' Please enter your login data here
			'UPSPickup: Pickup options for UPS including: 1) "customer counter" 2)"letter center" 3)"on call air" 4) "one time pickup" 5)"daily pickup" 6)"authorized shipping outlet" 7)"air service center". The default value is set to "customer counter."
			objShip.UPSPickup=Request.Params("UPS_Pickup")
			'UPSPack: Packing options for UPS including: 1) "Shipper Supplied Packaging" 2) "UPS Letter Envelope" or "UPS Letter" 3) "Your Packaging" or "Package" 4) "UPS Tube" 5) "UPS Pak" 6) "UPS Express Box" 7) "International UPS 25KG Box" or "UPS 25KG Box" 8) "International UPS 10KG Box" or "UPS 10KG Box" The default value is set to "Your Packaging."
			objShip.UPSPack=Request.Params("UPS_Pack")

			if sShippingOptions<>"" then
				sShippingOptions += ","
			end if
			sShippingOptions += "UPS"
		end if
		if Request.Params("USE_USPS")=1 then
			objShip.USPSLogin = "876EASYS2859,610EY45TV868" ' Please enter your login data here
      		if sShippingOptions<>"" then
				sShippingOptions += ","
			end if
      		sShippingOptions += "USPS"
      	end if
      	if Request.Params("USE_DHL")=1 then
      		'objShip.DHLServices = request.params("DHL_Service")
      		'if sShippingOptions<>"" then
				'   sShippingOptions += ","
			   'end if
      	   'sShippingOptions += "DHL"
      	end if
      	if Request.Params("USE_FEDEX") and 1=0 then
			'FedExPack: Packing Options for FedEx include: 1) "FedEx Express Envelope" or "FedEx Letter" , 2) "FedEx Express Pak" or "FedEx Pak" 3) "FedEx Express Box" or "FedEx Box" , 4) "FedEx Express Tube" or "FedEx Tube" 5) "My Packaging" or "Your Packaging". The default for this property is "My Packaging"
			objShip.FedExPack=Request.Params("FedEx_Pack")
			if request.params("Fedex_Ground") then
						objShip.FedExGround = true
					else
						objShip.FedExGround = false
					end if
      		if sShippingOptions<>"" then
				sShippingOptions += ","
            end if
      		sShippingOptions += "Fedex"
      	end if
      	if Request.Params("USE_AIRBORNE")=1 then
			if sShippingOptions<>"" then
				sShippingOptions += ","
            end if
      		sShippingOptions += "Airborne"
      	end if
        if Request.Params("USE_CANADA")=1 then
			objShip.CanadaPostLogin = Request.Params("Canada_Info")
            if sShippingOptions<>"" then
				sShippingOptions += ","
            end if
      		sShippingOptions += "CanadaPost"
      		objShip.OrigStateProvince = sArray(1)
		      objShip.DestStateProvince = sArray(4)
      	end if
      	if Request.Params("USE_CONWAY")=1 then
      		objShip.ConWayLogin = Request.Params("Conway_Info")
      		objShip.ConWay.SingleShipment = true
      		if sShippingOptions<>"" then
				sShippingOptions += ","
                end if
      		sShippingOptions += "Conway"
      		iConwayWeight=formatnumber(iMaxWeight,0)
            objShip.Weight=formatnumber(sArray(6),0)
      	end if

      			
      	objShip.Timeout=40
      
      	objShip.Sort="price"
      	if is_Residential<>"False" then
            objShip.Residential = true
         else
            objShip.Residential = false
         end if


        objShip.Rate(sShippingOptions)

        'if (objShip.Error)<> "" then
			'response.redirect(Request.Params("Site")&"/before_payment.asp?Rates_"&sArray(7)&"="&Server.urlencode(objShip.Error)&"&Error="&Server.urlencode(objShip.Error)&"&Info="&request.params("Info"))
		'else
			if iLeftover <= 0 then
			   sFinal=""
				For each rate in objShip.Rates
					sRate = rate.Charge*iMult
					sFinal += rate.Company+","+rate.Name+","+sRate+"|"
				Next
			else
				objShip2 = new dotnetSHIP.Ship()

				sArray=split(Service,",")
				objShip2.Weight= iLeftover
				objShip2.OrigCountry=sArray(2)
				objShip2.DestCountry= sArray(5)
				objShip2.DestZipPostal= sArray(3)
				objShip2.OrigZipPostal= sArray(0)
				'objShip2.OrigStateProvince = sArray(1)
				'objShip2.DestStateProvince = sArray(4)

				if Request.Params("USE_UPS") then
					objShip2.UPSLogin = Request.Params("UPS_Info") ' Please enter your login data here
					'UPSPickup: Pickup options for UPS including: 1) "customer counter" 2)"letter center" 3)"on call air" 4) "one time pickup" 5)"daily pickup" 6)"authorized shipping outlet" 7)"air service center". The default value is set to "customer counter."
					objShip2.UPSPickup=Request.Params("UPS_Pickup")
					'UPSPack: Packing options for UPS including: 1) "Shipper Supplied Packaging" 2) "UPS Letter Envelope" or "UPS Letter" 3) "Your Packaging" or "Package" 4) "UPS Tube" 5) "UPS Pak" 6) "UPS Express Box" 7) "International UPS 25KG Box" or "UPS 25KG Box" 8) "International UPS 10KG Box" or "UPS 10KG Box" The default value is set to "Your Packaging."
					objShip2.UPSPack=Request.Params("UPS_Pack")

					if sShippingOptions<>"" then
						sShippingOptions += ","
					end if
					sShippingOptions += "UPS"
				end if
				if Request.Params("USE_USPS") then
					objShip2.USPSLogin = "876EASYS2859,610EY45TV868" ' Please enter your login data here
      				if sShippingOptions<>"" then
						sShippingOptions += ","
					end if
      				sShippingOptions += "USPS"
      			end if
      			if Request.Params("USE_DHL") then
      				objShip2.DHLServices = request.Params("DHL_Service")
      				if sShippingOptions<>"" then
						sShippingOptions += ","
					end if
      			sShippingOptions += "DHL"
      			end if
      			if Request.Params("USE_FEDEX") and 1=0 then
					'FedExPack: Packing Options for FedEx include: 1) "FedEx Express Envelope" or "FedEx Letter" , 2) "FedEx Express Pak" or "FedEx Pak" 3) "FedEx Express Box" or "FedEx Box" , 4) "FedEx Express Tube" or "FedEx Tube" 5) "My Packaging" or "Your Packaging". The default for this property is "My Packaging"
					objShip2.FedExPack=Request.Params("FedEx_Pack")
					if request.params("Fedex_Ground") then
						objShip2.FedExGround = true
					else
						objShip2.FedExGround = false
					end if
      				if sShippingOptions<>"" then
						sShippingOptions += ","
					end if
      				sShippingOptions += "Fedex"
      			end if
      			if Request.Params("USE_AIRBORNE") then
					if sShippingOptions<>"" then
						sShippingOptions += ","
					end if
      				sShippingOptions += "Airborne"
      			end if
				if Request.Params("USE_CANADA") then
					objShip2.CanadaPostLogin = Request.Params("Canada_Info")
					if sShippingOptions<>"" then
						sShippingOptions += ","
					end if
      				sShippingOptions += "CanadaPost"
      			end if
      			if Request.Params("USE_CONWAY") then
      				objShip2.ConWayLogin = Request.Params("Conway_Info")
      				objShip2.ConWay.SingleShipment = true
      			end if
      					
      			objShip2.Timeout=40
      
      			objShip2.Sort="price"
				if is_Residential<>"False" then
            objShip2.Residential = true
         else
            objShip2.Residential = false
         end if


				objShip2.Rate(sShippingOptions)

				if (objShip2.Error)<> "" and sArray(2)<>"ZW" then
					response.redirect(Request.Params("Site")&"/before_payment.asp?Rates_"&sArray(7)&"="&request.params("Query")&"&Error="&Server.urlencode(objShip2.Error)&"&Info="&request.params("Info"))
				else
				   sFinal=""	
               For each rate2 in objShip2.Rates
						For each rate in objShip.Rates
							if (rate2.Company = rate.Company) and (rate2.Name = rate.Name) then
								sRate = (rate.Charge*iMult) + rate2.Charge
								sFinal += rate.Company+","+rate.Name+","+sRate+"|"
							end if
						Next
					Next
				end if
            
            end if
		end if
		sFinalString += "Rates_"&sArray(7)&"="+sFinal+"&"
	'end if
	if objShip.Error <> "" and sArray(2)<>"ZW" then
	   sFinalString += "Error="&Server.urlencode(objShip.Error)&"&Info="&request.params("Info")&"&"
	end if
Next

response.redirect (replace((Request.Params("Site")+"before_payment.asp?"+sFinalString+request.params("Query")),"&&","&"))


end sub
</script>


