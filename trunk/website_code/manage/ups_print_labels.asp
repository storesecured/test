
<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

	calendar=1
    Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")

	sTitle = "Print Labels"
	sFormAction = "ups_print_labels.asp"
	thisRedirect = "ups_print_labels.asp"
	sFormName ="Create_Labels"
	sMenu = "orders"
	addPicker=1
	createHead thisRedirect
	if Service_Type < 7  then %>
	<TR bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>

	</td></tr>
   <% createFoot thisRedirect, 0%>
<% else
	
   if Request.Form("RefNo") = "" then		'before selecting service/pack type

	Start_Date = FormatDateTime(now(),2)
	
	EditOrders = request.querystring("Orders")
	if EditOrders = "" then
	   fn_error "You must select at least one order"
	else

		sql_select = "select * from Store_purchases WHERE (OID in("&EditOrders&")) AND store_id = "&Store_Id
		rs_Store.open sql_select,conn_store,1,1
		if rs_Store.eof and rs_Store.bof then
		   response.write "Could not find any records to create labels for."
		else


				FromCompany=Store_Company
				FromName= Store_Name
				FromFirm= Store_Name
				FromAddress1 = Store_Address1
				FromAddress2 = Store_Address2
				
				FromCountry = Store_Country

				FromCity = Store_City
				FromState = Store_state
				FromZip5  = Store_Zip
				FromZip4 = ""

				ToCompany =  rs_Store("ShipCompany")
				ToName= rs_Store("ShipFirstName") & " " & rs_store("ShipLastName")
				ToAddress1	= rs_Store("ShipAddress1")
				 if rs_Store("ShipAddress2") <> "" then
					  ToAddress2 =  rs_Store("ShipAddress2")
				 end if
				ToCity = rs_Store("ShipCity")
				ToState = rs_Store("ShipState")
				ToZip5 = rs_Store("ShipZip")	
				ToZip4=""
				ToCountry = rs_Store("ShipCountry")

				ShipMethod=rs_Store("Shipping_method_name")
				RefNo=EditOrders

				rs_Store.close

					sql_real_time = "select * from Store_Real_Time_Settings where store_id="&store_id
					rs_Store.open sql_real_time, conn_store, 1, 1
						
						UPS_User = rs_store("UPS_User")
						UPS_Password = rs_store("UPS_Password")
						UPS_AccessLicense = rs_store("UPS_AccessLicense")
						UPS_Account_Number = rs_store("UPS_Account_Number")
						UPS_Pack=rs_store("UPS_Pack")
						UPS_Pickup=rs_store("UPS_Pickup")
					rs_Store.Close


				sql_weight="select sum (item_weight*quantity) as total_weight from Store_Transactions where oid="&RefNo&" and store_id= "&store_id
				set rs_weight= conn_store.Execute(sql_weight)
				Item_Weight=rs_weight("total_weight")
				set rs_weight=nothing
				
				Total_Weight=Item_Weight + Handling_Weight
				%>
					
				
						<TR bgcolor='#FFFFFF'>
	<td colspan=4>
		This feature is currently undergoing certification with UPS.  If you do not already have permission from UPS to generate shipping labels with your access license you will be unable to use this functionality until the application is fully certified with UPS.

	</td></tr><TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Service</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<select class="inputvalue" name="serviceType" id="serviceType">
										<option value="">Select Service Type</option>
										<option value="01">UPS Next Day Air</option>
										<option value="02">UPS 2nd Day Air</option>
										<option value="03">UPS Ground</option>
										<option value="07">UPS Worldwide ExpressSM</option>
										<option value="08">UPS Worldwide ExpeditedSM</option>
										<option value="11">UPS Standard</option>
										<option value="12">UPS 3-Day Select</option>
										<option value="13">UPS Next Day Air Saver</option>
										<option value="14">UPS Next Day Air. Early AM</option>
										<option value="54">UPS Worldwide Express PlusSM</option>
										<option value="59">UPS 2nd Day Air A.M.</option>
										<option value="65">UPS Express Saver</option>
									</select>
									<% small_help "Service" %>
								</td>
						</tr>
								 

						<TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Packing</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<select class="inputvalue" name="packType" id="packType"  value="Package">
										<option value="">Select Packing Type</option>
										<option value="00">Your Packaging</option>
										<option value="01">UPS Letter / Express Envelope</option>
										<option value="02">Package</option>
										<option value="03">UPS Tube</option>
										<option value="04">UPS Pak</option>
										<option value="21">UPS Express Box</option>
										<option value="24">UPS 25KG Box®</option>
										<option value="25">UPS 10KG Box®</option>
									</select>
										<% small_help "Packing" %>
								</td>
						</tr>

								
						<TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Package Weight</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<input name="Weight" id="Weight" type="text" size="5" maxlength="5" value="<%=Total_Weight%>" onKeyPress="return goodchars(event,'0123456789.')">
								<input type="hidden" name="Weight_C" value="Op|Integer|||||Weight">
								<% small_help "Weight" %>
								</td>
						</tr>


							<TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Package  Length</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<input name="Pack_Length" id="Pack_Length" type="text" size="3" maxlength="3"  value="<%=Pack_Length%>" onKeyPress="return goodchars(event,'0123456789/')">
								<% small_help "Pack_Length" %>
								</td>
						</tr>

							<TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Package Width</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<input name="Pack_Width" id="Pack_Width" type="text" size="3" maxlength="3"  value="<%=Pack_Width%>" onKeyPress="return goodchars(event,'0123456789/')">
								<% small_help "Pack_Width" %>
								</td>
						</tr>

							<TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Package Height</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<input name="Pack_Height" id="Pack_Height" type="text" size="3" maxlength="3" value="<%=Pack_Height%>" onKeyPress="return goodchars(event,'0123456789/')">
								<% small_help "Pack_Height" %>
								</td>
						</tr>
						
						<TR bgcolor='#FFFFFF'>
								<td width="35%" class="inputname"><B>Declared Amount</B></td>
								<td width="65%" class="inputvalue" colspan="2" >
									<input name="Insured_Amount" id="Insured_Amount" type="text" size="6" maxlength="6"  value="<%=Insured_Amount%>" onKeyPress="return goodchars(event,'0123456789.')">
								<% small_help "Insured_Amount" %>
								</td>
						</tr>


							<TR bgcolor='#FFFFFF'>
									<td width="35%" class="inputname"><B>Saturday Delivery</B></td>
									<td width="65%" class="inputvalue" colspan="2" >
										<input class="image" type="checkbox" <%= Saturday_Delivery %> name="Saturday_Delivery" value="-1" >
										<% small_help "Saturday Delivery" %>
									</td>
							</tr>	


					<% if UPS_Pickup="on call air" then %>

								
								<TR bgcolor='#FFFFFF'>
										<td width="35%" class="inputname"><B>Pickup date</B></td>
										<td width="65%" class="inputvalue" colspan="2" >
									
										<SCRIPT LANGUAGE="JavaScript" ID="jscal1">
									      var now = new Date();
									      var cal1 = new CalendarPopup("testdiv1");
									      cal1.showNavigationDropdowns();
			 						  </SCRIPT>
										<input name="Pickupdate"  size="10" maxlength=10 value="<%= FormatDateTime(Start_date,2) %>" onKeyPress="return goodchars(event,'0123456789/')">
								
									   <A HREF="#" onClick="cal1.select(document.forms[0].Pickupdate,'anchor1','MM/dd/yyyy',(document.forms[0].Pickupdate.value=='')?document.forms[0].Pickupdate.value:null); return false;"  NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>	<input type="hidden"  name="Pickupdate_C" value="Re|date|||||Start Date">
										<% small_help "Pickup date" %>
										</td>
								</tr>
								
										<TR bgcolor='#FFFFFF'>
										<td width="35%" class="inputname"><B>Earliest Time Ready</B></td>
										<td width="65%" class="inputvalue" colspan="2" >
											<input name="Earliest_Time_Ready" id="Earliest_Time_Ready" type="text" size="4" maxlength="4" onKeyPress="return goodchars(event,'0123456789/')">
										<% small_help "Earliest Time Ready" %>
										</td>
								</tr>


								<TR bgcolor='#FFFFFF'>
										<td width="35%" class="inputname"><B>Latest Time Ready</B></td>
										<td width="65%" class="inputvalue" colspan="2" >
											<input name="Latest_Time_Ready" id="Latest_Time_Ready" type="text" size="4" maxlength="4"onKeyPress="return goodchars(event,'0123456789/')">
											<% small_help "Latest Time Ready" %>
										</td>
								</tr>


								<TR bgcolor='#FFFFFF'>
										<td width="35%" class="inputname"><B>Contact Name</B></td>
										<td width="65%" class="inputvalue" colspan="2" >
											<input name="Contact_Name" id="Contact_Name" type="text" size="35" maxlength="35">
											<% small_help "Contact Name" %>
										</td>
								</tr>
								
								<TR bgcolor='#FFFFFF'>
										<td width="35%" class="inputname"><B>Contact Phone Number</B></td>
										<td width="65%" class="inputvalue" colspan="2" >
											<input name="Contact_Phone_Number" id="Contact_Phone_Number" type="text" size="15" maxlength="15" onKeyPress="return goodchars(event,'0123456789/')">
											<% small_help "Contact Phone Number" %>
										</td>
								</tr>
				
					<%end if%>
			
				
						<input name="FromCompany" type="hidden" value="<%=FromCompany%>">	
						<input name="FromName" type="hidden" value="<%=FromName%>">
						<input name="FromFirm" type="hidden" value="<%=FromFirm%>">
						<input name="FromAddress1" type="hidden" value="<%=FromAddress1%>">
						<input name="FromAddress2" type="hidden" value="<%=FromAddress2%>">
						<input name="FromCity" type="hidden" value="<%=FromCity%>">
						<input name="FromState" type="hidden" value="<%=FromState%>">
						<input name="FromCountry" type="hidden" value="<%=FromCountry%>">
						<input name="FromZip5" type="hidden" value="<%=FromZip5%>">
						<input name="FromZip4" type="hidden" value="<%=FromZip4%>">

						
						<input name="ToCompany" type="hidden" value="<%=ToCompany%>">	
						<input name="ToName" type="hidden" value="<%=ToName%>">
						<input name="ToAddress1" type="hidden" value="<%=ToAddress1%>">
						<input name="ToAddress2" type="hidden" value="<%=ToAddress2%>">
						<input name="ToCity" type="hidden" value="<%=ToCity%>">
						<input name="ToState" type="hidden" value="<%=ToState%>">
						<input name="ToZip5" type="hidden" value="<%=ToZip5%>">
						<input name="ToZip4" type="hidden" value="<%=ToZip4%>">
						<input name="ToCountry" type="hidden" value="<%=ToCountry%>">
				 
						<input name="RefNo" type="hidden" value="<%=RefNo%>">


						<script language="JavaScript">
							var obj;
							obj= document.forms[0].serviceType;
							for (k=0;k<obj.options.length;k++)
								if (obj.options[k].innerText == "<%=ShipMethod%>")	
									obj.options[k].selected="selected"
							
							obj= document.forms[0].packType;
							for (k=0;k<obj.options.length;k++)
								if (obj.options[k].innerText == "<%=UPS_Pack%>")	
									obj.options[k].selected="selected"
									
							 
							</script>


						<TR bgcolor='#FFFFFF'>
							<td width="100%"  colspan="4" align="center">
								<input name="Get_Label_Now " type="button" class="Buttons" value="Get Label" onclick="document.forms[0].submit();">
							</td>
						</tr>


					
		<%
		end if
		rs_Store.close
	end if
	createFoot thisRedirect, 2
	%>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

frmvalidator.addValidation("serviceType","req","Please select a service.");
frmvalidator.addValidation("packType","req","Please select a packing type.");
frmvalidator.addValidation("Weight","req","Please enter Weight.");

</script>




<% else				'after sending confirm request
%>

	<script>
		function getLabel()
		{

			if (confirm("This will charge your UPS Account, Are you sure?"))
			{
				document.forms[0].submit(); 
			}
		}

	</script>


<%
		sTitle = "Print Labels"
		sFormAction = "ups_print_labels_file.asp"
	'	thisRedirect = "ups_print_labels.asp"
		sFormName ="Create_Labels_Confirm"
		sMenu = "Advanced"
		addPicker=1
		createHead thisRedirect
		


		FromName= Request.Form("FromName")
		FromFirm=  Request.Form("FromFirm")
		FromCompany=Request.Form("FromCompany")
		FromAddress1= Request.Form("FromAddress1")
		FromAddress2=  Request.Form("FromAddress2")
        FromAddress2=FromAddress1
		FromCity= Request.Form("FromCity")
		FromState=  Request.Form("FromState")
		FromZip5= Request.Form("FromZip5")
		FromZip4=  Request.Form("FromZip4")
		FromCountry=Request.Form("FromCountry")

		ToCompany=Request.Form("ToCompany")
		ToName= Request.Form("ToName")
		ToAddress1=  Request.Form("ToAddress1")
		ToAddress2= Request.Form("ToAddress2")
		ToCity=  Request.Form("ToCity")
		ToState= Request.Form("ToState")
		ToZip5=  Request.Form("ToZip5")
		ToZip4= Request.Form("ToZip4")
		ToCountry=Request.Form("ToCountry")



		'Pickupdate=Request.Form("Pickupdate")
		  if Request.Form("Pickupdate")<>"" then
			if isdate(Request.Form("Pickupdate")) then
				Start_Date = Request.Form("Pickupdate")
				arr_Start_date=split(Start_Date, "/")
				Pickup_Date= arr_Start_date(2)+arr_Start_date(0)+arr_Start_date(1)
			end if
		  end if
			
		


		
		Earliest_Time_Ready=Request.Form("Earliest_Time_Ready")
		Latest_Time_Ready=Request.Form("Latest_Time_Ready")
		Contact_Name=Request.Form("Contact_Name")
		Contact_Phone_Number=Request.Form("Contact_Phone_Number")


		'mapping country codes	
		select case LCase(Request.Form("ToCountry"))
			case "united states":
				ToCountry="US"
		end select
		

		select case Lcase(Request.Form("FromCountry"))
			case "united states":
				FromCountry="US"
		end select

		RefNo= Request.Form("RefNo")
		
		serviceCode=Request.Form("serviceType")
		packCode=Request.Form("packType")
		weight=Request.Form("weight")


		Pack_Length=Request.Form("Pack_Length")
		Pack_Width=Request.Form("Pack_Width")
		Pack_Height=Request.Form("Pack_Height")
		Insured_Amount=Request.Form("Insured_Amount")

		Saturday_Delivery=Request.Form("Saturday_Delivery")

		sql_real_time = "select * from Store_Real_Time_Settings where store_id="&store_id
		rs_Store.open sql_real_time, conn_store, 1, 1
			UPS_User = rs_store("UPS_User")
			UPS_Password = rs_store("UPS_Password")
			UPS_AccessLicense = rs_store("UPS_AccessLicense")
			UPS_Account_Number = rs_store("UPS_Account_Number")
			UPS_Pack=rs_store("UPS_Pack")
			UPS_Pickup=rs_store("UPS_Pickup")
		rs_Store.Close

		'getting the access info
		Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.load(Server.MapPath( "xml/access.xml")) 
		
		set nodeUPS_User = xmlDoc.documentElement.selectsinglenode("UserId")
		nodeUPS_User.text=UPS_User
		'nodeUPS_User.text="melanieeasystore"				'need to remove after testing

		set nodeUPS_Password = xmlDoc.documentElement.selectsinglenode("Password")
		nodeUPS_Password.text=UPS_Password
		'nodeUPS_Password.text="ankle237"						'need to remove after testing

		set nodeUPS_AccessLicense = xmlDoc.documentElement.selectsinglenode("AccessLicenseNumber")
		nodeUPS_AccessLicense.text=UPS_AccessLicense
		'nodeUPS_AccessLicense.text="FBBA98A10FDA0F48"		'need to remove after testing

		strAccess=xmlDoc.xml
		

		'sending the confirm request			
		Set xmlDocReq= Server.CreateObject("MSXML2.DOMDocument")
		xmlDocReq.load(Server.MapPath( "xml/ups_confirm.xml")) 
		
		set nodeService = xmlDocReq.documentElement.selectsinglenode("Shipment/Service/Code")
		nodeService.text=serviceCode
		
		set nodePack = xmlDocReq.documentElement.selectsinglenode("Shipment/Package/PackagingType/Code")
		nodePack.text=PackCode
		
		set nodeWeight = xmlDocReq.documentElement.selectsinglenode("Shipment/Package/PackageWeight/Weight")
		nodeWeight.text=weight

		set nodeReferenceNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/Package/ReferenceNumber/Value")
		nodeReferenceNumber.text=RefNo

		set nodePackage = xmlDocReq.documentElement.selectsinglenode("Shipment/Package")
		
		if Pack_Length <> ""  or Pack_Width <> "" or Pack_Height <> "" then
			set nodePackageDimensions=nodePackage.appendchild(xmlDocReq.createElement("Dimensions"))
		end if

		if Pack_Length <> "" then
			set nodePackageLength=nodePackageDimensions.appendchild(xmlDocReq.createElement("Length"))
			nodePackageLength.text=Pack_Length
		end if

		if Pack_Width <> "" then
			set nodePackageWidth=nodePackageDimensions.appendchild(xmlDocReq.createElement("Width"))
			nodePackageWidth.text=Pack_Width
		end if

		if Pack_Height <> "" then 
			set nodePackageHeight=nodePackageDimensions.appendchild(xmlDocReq.createElement("Height"))
			nodePackageHeight.text=Pack_Height
		end if		

		if Insured_Amount <> "" then
			set nodePackageServiceOptions=nodePackage.appendchild(xmlDocReq.createElement("PackageServiceOptions"))
			set nodeInsuredValue=nodePackageServiceOptions.appendchild(xmlDocReq.createElement("InsuredValue"))
			'set nodeInsuredCurrencyCode=nodeInsuredValue.appendchild(xmlDocReq.createElement("CurrencyCode"))
			set nodeInsuredMonetaryValue=nodeInsuredValue.appendchild(xmlDocReq.createElement("MonetaryValue"))
			nodeInsuredMonetaryValue.text = Insured_Amount
		end if

		'-----------------------  computing Oversize ----------------------------------------------------------------------------------------------------------------
		'OSP=0			
		'LengthGirth = Pack_Length + 2  *  (Pack_Width + Pack_Height)

		'if LengthGirth > 84 and LengthGirth <= 108 and weight < 30 then
		'		OSP=1
		'elseif LengthGirth > 108 and weight < 70 then
		'		OSP=2
		'elseif LengthGirth > 108 and weight >= 70 then
		'		OSP=3
		'end if

		'if OSP <> 0 and serviceCode = "03" then
		'	set nodeOversizePackage=nodePackage.appendchild(xmlDocReq.createElement("OversizePackage"))
		'	nodeOversizePackage.text =OSP
		'end if

		'-----------------------  computing Oversize ends ----------------------------------------------------------------------------------------------------------------



'		set nodeAccountNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/PaymentInformation/Prepaid/BillShipper/AccountNumber")
'		nodeAccountNumber.text=UPS_Account_Number
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------


		'----------------------------------------------------------------Shipper/Ship From Info----------------------------------------------------------	
	
		set nodeShipperName = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Name")
		nodeShipperName.text=FromCompany

		set nodeCompanyAttn = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/AttentionName")
		nodeCompanyAttn.text=FromName

		set nodeShipperNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/ShipperNumber")
		nodeShipperNumber.text=UPS_Account_Number

		set nodeShipperAddressLine1 = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Address/AddressLine1")
		nodeShipperAddressLine1.text=FromAddress1
	
		set nodeShipperAddressLine2 = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Address/AddressLine2")		
		if ToAddress2 <> "" then
			nodeShipperAddressLine2.text=FromAddress2
		else
			nodeShipperAddressLine2.text=""
		end if

		set nodeShipperCity = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Address/City")
		nodeShipperCity.text=FromCity

		set nodeShipperStateProvinceCode = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Address/StateProvinceCode")
		nodeShipperStateProvinceCode.text=FromState

		set nodeShipperPostalCode = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Address/PostalCode")
		nodeShipperPostalCode.text=FromZip5

		set nodeShipperCountryCode = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/Address/CountryCode")
		nodeShipperCountryCode.text=FromCountry

		'set nodePhoneCountryCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/PhoneNumber/StructuredPhoneNumber/PhoneCountryCode")
		'nodePhoneCountryCode.text="34"

		'set nodeShipperPhoneDialPlanNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/PhoneNumber/StructuredPhoneNumber/PhoneDialPlanNumber")
		'nodePhoneDialPlanNumber.text="333656"

		'set nodeShipperPhoneLineNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/Shipper/PhoneNumber/StructuredPhoneNumber/PhoneLineNumber")
		'nodePhoneLineNumber.text="7337"


		'------------------------------------------------------------------------------Ship To Info------------------------------------------------
		set nodeToAttn = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/AttentionName")
		nodeToAttn.text=ToCompany

		set nodeToCompanyName = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/CompanyName")
		nodeToCompanyName.text=ToName

		set nodeToAddressLine1 = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/Address/AddressLine1")
		nodeToAddressLine1.text=ToAddress1

		set nodeToAddressLine2 = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/Address/AddressLine2")
		if ToAddress2 <> "" then
			nodeToAddressLine2.text=ToAddress2
		else
			nodeToAddressLine2.text=""
		end if
		
		set nodeToCity = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/Address/City")
		nodeToCity.text=ToCity

		set nodeToStateProvinceCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/Address/StateProvinceCode")
		nodeToStateProvinceCode.text=ToState

		set nodeToPostalCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/Address/PostalCode")
		nodeToPostalCode.text=ToZip5

		set nodeToCountryCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/Address/CountryCode")
		nodeToCountryCode.text=ToCountry


		'set nodeToPhoneCountryCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/PhoneNumber/StructuredPhoneNumber/PhoneCountryCode")
		'nodeFromPhoneCountryCode.text="111616"

		'set nodeToPhoneDialPlanNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/PhoneNumber/StructuredPhoneNumber/PhoneDialPlanNumber")
		'nodeFromPhoneDialPlanNumber.text="7117"

		
		'set nodeToPhoneLineNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipTo/PhoneNumber/StructuredPhoneNumber/PhoneLineNumber")
		'nodeFromPhoneLineNumber.text="9119"

		'------------------------------------------------------------------------------Ship From Info------------------------------------------------
		set nodeFromAttn = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/AttentionName")
		nodeFromAttn.text=FromCompany

		set nodeFromCompanyName = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/CompanyName")
		nodeFromCompanyName.text=FromName

		set nodeFromAddressLine1 = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/Address/AddressLine1")
		nodeFromAddressLine1.text=ToAddress1

		set nodeFromAddressLine2 = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/Address/AddressLine2")
		if FromAddress2 <> "" then
			nodeFromAddressLine2.text=FromAddress2
		else
			nodeFromAddressLine2.text=""
		end if
		
		set nodeFromCity = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/Address/City")
		nodeFromCity.text=FromCity

		set nodeFromStateProvinceCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/Address/StateProvinceCode")
		nodeFromStateProvinceCode.text=FromState

		set nodeFromPostalCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/Address/PostalCode")
		nodeFromPostalCode.text=FromZip5

		set nodeFromCountryCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/Address/CountryCode")
		nodeFromCountryCode.text=FromCountry


		'set nodeFromPhoneCountryCode = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/PhoneNumber/StructuredPhoneNumber/PhoneCountryCode")
		'nodeFromPhoneCountryCode.text="111616"

		'set nodeFromPhoneDialPlanNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/PhoneNumber/StructuredPhoneNumber/PhoneDialPlanNumber")
		'nodeFromPhoneDialPlanNumber.text="7117"

		
		'set nodeFromPhoneLineNumber = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipFrom/PhoneNumber/StructuredPhoneNumber/PhoneLineNumber")
		'nodeFromPhoneLineNumber.text="9119"

			'set nodeShipment = xmlDocReq.documentElement.selectsinglenode("Shipment")

			'set nodeShipmentServiceOptions=nodeShipment.appendchild(xmlDocReq.createElement("ShipmentServiceOptions"))

		'------------------------------------------------------- On Call Air---------------------------------------------------------------------------------

		'set nodePickupDate = xmlDocReq.documentElement.selectsinglenode("Shipment/ShipmentServiceOptions/OnCallAir/PickupDetails/PickupDate")

		if UPS_Pickup = "On Call Air" then
		
			set nodeShipment = xmlDocReq.documentElement.selectsinglenode("Shipment")

			set nodeShipmentServiceOptions=nodeShipment.appendchild(xmlDocReq.createElement("ShipmentServiceOptions"))

			set nodeOnCallAir=nodeShipmentServiceOptions.appendchild(xmlDocReq.createElement("OnCallAir"))

			set nodePickupDetails=nodeOnCallAir.appendchild(xmlDocReq.createElement("PickupDetails"))
				
			set nodePickupDate=nodePickupDetails.appendchild(xmlDocReq.createElement("PickupDate"))
			nodePickupDate.text=Pickup_Date

			set nodeEarliestTimeReady=nodePickupDetails.appendchild(xmlDocReq.createElement("EarliestTimeReady"))
			nodeEarliestTimeReady.text=Earliest_Time_Ready

			set nodeLatestTimeReady=nodePickupDetails.appendchild(xmlDocReq.createElement("LatestTimeReady"))
			nodeLatestTimeReady.text=Latest_Time_Ready

			set nodeContactInfo=nodeOnCallAir.appendchild(xmlDocReq.createElement("ContactInfo"))
				
			set nodeContact_Info_Name=nodeContactInfo.appendchild(xmlDocReq.createElement("Name"))
			nodeContact_Info_Name.text=Contact_Name

			set nodeContact_Info_PhoneNumber=nodeContactInfo.appendchild(xmlDocReq.createElement("PhoneNumber"))
			nodeContact_Info_PhoneNumber.text=Contact_Phone_Number
		
			if Saturday_Delivery <> "" then
				set nodeSaturdayDelivery=nodeShipmentServiceOptions.appendchild(xmlDocReq.createElement("SaturdayDelivery"))
			end if
	

			'-----------------------------------------------------------------------------------------------------------------------------------------------------------
		end if
		'---------------------------------------------------------------------- Billing Info ------------------------------------------------------------------------------------
		set nodePaymentInformation=xmlDocReq.documentElement.selectsinglenode("Shipment/PaymentInformation")

		set nodePrepaid=nodePaymentInformation.appendchild(xmlDocReq.createElement("Prepaid"))
		set nodeBillShipper=nodePrepaid.appendchild(xmlDocReq.createElement("BillShipper"))
		set nodeBillShipper_AccountNumber=nodeBillShipper.appendchild(xmlDocReq.createElement("AccountNumber"))

		nodeBillShipper_AccountNumber.text=UPS_Account_Number

		'------------------------------------------------- Bill Third Party ------------------------------------------------------------------------

	'	set nodePaymentInformation=xmlDocReq.documentElement.selectsinglenode("Shipment/PaymentInformation")
	'	set nodeBillThirdParty=nodePaymentInformation.appendchild(xmlDocReq.createElement("BillThirdParty"))
	'	set nodeBillThirdPartyShipper=nodeBillThirdParty.appendchild(xmlDocReq.createElement("BillThirdPartyShipper"))


	'	set nodeThirdParty_AccountNumber=nodeBillThirdPartyShipper.appendchild(xmlDocReq.createElement("AccountNumber"))
	'	nodeThirdParty_AccountNumber.text=UPS_Account_Number


	'	set nodeThirdParty=nodeBillThirdPartyShipper.appendchild(xmlDocReq.createElement("ThirdParty"))
	'	set nodeThirdParty_Address=nodeThirdParty.appendchild(xmlDocReq.createElement("Address"))
		
	'	set nodeThirdParty_PostalCode=nodeThirdParty_Address.appendchild(xmlDocReq.createElement("PostalCode"))
	'	nodeThirdParty_PostalCode.text=FromZip5	

	'	set nodeThirdParty_CountryCode=nodeThirdParty_Address.appendchild(xmlDocReq.createElement("CountryCode"))
	'	nodeThirdParty_CountryCode.text=FromCountry


		'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------	
		strReq=xmlDocReq.xml
	
		' sending the confirm request
		Set xmlHTTP = Server.CreateObject("Microsoft.XMLHTTP")
		xmlHTTP.open "POST","https://www.ups.com/ups.app/xml/ShipConfirm", False
		xmlHTTP.send(strAccess+strReq)

        Upload_Folder = fn_get_sites_folder(Store_Id,"Upload")
		'saving the confirm response to an xml file	
		Set oConfirm= Server.CreateObject("MSXML2.DOMDocument")
		oConfirm.loadXML(xmlHTTP.ResponseText)
		oConfirm.save(Upload_Folder&"ups_confirm_op_"&RefNo&".xml")
	
		set nodeResponseCode= oConfirm.documentElement.selectsinglenode("Response/ResponseStatusCode")

		if nodeResponseCode.text <> 0 then	 ' if confirm request is successful (Code=0)
			
			
			Set oConfirm= Server.CreateObject("MSXML2.DOMDocument")
			oConfirm.load(Upload_Folder&"ups_confirm_op_"&RefNo&".xml")

			
			set nodeConfResponse=oConfirm.documentElement.selectsinglenode("Response")
			if nodeConfResponse.selectsinglenode("Error") <> null then
					set nodeConfError=oConfirm.documentElement.selectsinglenode("Response/Error/ErrorDescription")
					set nodeConfErrorCode=oConfirm.documentElement.selectsinglenode("Response/Error/ErrorCode")
					set nodeConfErrorSeverity=oConfirm.documentElement.selectsinglenode("Response/Error/ErrorSeverity")
			end if

			set nodeTransportationChargesCurrency = oConfirm.documentElement.selectsinglenode("ShipmentCharges/TransportationCharges/CurrencyCode")
			
			set nodeTransportationChargesMonetaryValue = oConfirm.documentElement.selectsinglenode("ShipmentCharges/TransportationCharges/MonetaryValue")

			set nodeServiceOptionsChargesCurrency = oConfirm.documentElement.selectsinglenode("ShipmentCharges/ServiceOptionsCharges/CurrencyCode")
			
			set nodeServiceOptionsChargesMonetaryValue = oConfirm.documentElement.selectsinglenode("ShipmentCharges/ServiceOptionsCharges/MonetaryValue")

			set nodeTotalChargesCurrency = oConfirm.documentElement.selectsinglenode("ShipmentCharges/TotalCharges/CurrencyCode")
	
			set nodeTotalChargesMonetaryValue = oConfirm.documentElement.selectsinglenode("ShipmentCharges/TotalCharges/MonetaryValue")


%>

			<TR bgcolor='#FFFFFF'>
				<td>
					<div class="bar" style="padding-left: 5px;padding-top: 10px;">
					<font size="2" face="tahoma" color="black"><b>Confirm Shipment</b></font>
					</div>
				</td>
			</tr>
		
			<TR bgcolor='#FFFFFF'>
					<td width="25%" class="inputname"><B>Transportation Charges</B></td>
					<td width="75%" class="inputvalue" colspan="2" >
					<% =nodeTransportationChargesMonetaryValue.text & " " & nodeTransportationChargesCurrency.text %>
					</td>
			</tr>


			<TR bgcolor='#FFFFFF'>
					<td width="25%" class="inputname"><B>Service Options Charges</B></td>
					<td width="75%" class="inputvalue" colspan="2" >
					<% =nodeServiceOptionsChargesMonetaryValue.text & " " & nodeServiceOptionsChargesCurrency.text %>
					</td>
			</tr>


			

			<TR bgcolor='#FFFFFF'>
					<td width="25%" class="inputname"><B>Total Charges</B></td>
					<td width="75%" class="inputvalue" colspan="2" >
					<% =nodeTotalChargesMonetaryValue.text & " " & nodeTotalChargesCurrency.text %>
					</td>
			</tr>

			<TR bgcolor='#FFFFFF'>
				<td width="50%" align="center">
					<input name="Back_Label " type="button" class="Buttons" OnClick=JavaScript:self.location="ups_print_labels.asp?Orders=<%=RefNo%>" value="Back">
				</td>
				<td width="50%"  align="center" colspan=4>
				<!--	<input name="Get_Label_Now " type="submit" class="Buttons" value="Get Label Now">-->
						<input name="Get_Label_Now " type="button" class="Buttons" value="Get Label Now" onclick="javascript:getLabel()">
				</td>
			</tr>
			<input name="RefNo" type="hidden" value="<%=RefNo%>">


<%
		else			' confirm request failed (Code = 1)
			set nodeError=oConfirm.documentElement.selectsinglenode("Response/Error/ErrorDescription")
			 fn_error nodeError.text
		end if

	set oConfirm=nothing
end if		'main Request.Form if condition

end if
%>
