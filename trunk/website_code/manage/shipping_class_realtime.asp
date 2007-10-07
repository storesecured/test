<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="help/shipping_class_realtime.asp"-->

<% 

'LOAD CURRENT VALUES FROM THE DATABASE

on error resume next
sql_real_time = "select Store_Real_Time_Settings.*,Shipping_Classes,use_ups,use_usps,use_fedex,use_dhl,use_canada,use_airborne,ship_multi,show_shipping,handling_fee,handling_weight from Store_Real_Time_Settings inner join store_settings on store_real_time_settings.store_id=store_settings.store_id where store_settings.store_id="&store_id
rs_Store.open sql_real_time, conn_store, 1, 1
	Ship_Multi = rs_store("Ship_Multi")
        Show_Shipping = rs_store("Show_Shipping")
        sHandling_Fee = rs_store("Handling_Fee")
        sHandling_Weight = rs_store("Handling_Weight")
        UPS_User = rs_store("UPS_User")
	UPS_Password = rs_store("UPS_Password")
	USPS_User = rs_store("USPS_User")
	USPS_Password = rs_store("USPS_Password")
	UPS_AccessLicense = rs_store("UPS_AccessLicense")
	UPS_Account_Number=rs_store("UPS_Account_Number")
	UPS_Shipper_Name=rs_store("UPS_Shipper_Name")
	Canada_Login = rs_store("Canada_Login")
	UPS_Pickup = rs_store("UPS_Pickup")
	UPS_Pack = rs_store("UPS_Pack")
	Fedex_Pack = rs_store("Fedex_Pack")
	Fedex_Ground = rs_store("Fedex_Ground")
	if Fedex_Ground then
	   Fedex_Ground="checked"
	end if
	DHL_Service = rs_store("DHL_Service")
	Countries=rs_store("Countries")
	Restrict_Options=rs_store("Restrict_Options")
	Max_Weight=rs_Store("Max_Weight")
	Shipping_Class=rs_store("Shipping_Classes")
        Use_UPS=rs_Store("Use_UPS")
        Use_USPS=rs_Store("Use_USPS")
        Use_Fedex=rs_Store("Use_Fedex")
        Use_Canada=rs_Store("Use_Canada")
        Use_DHL=rs_Store("Use_DHL")
        Use_Airborne=rs_Store("Use_Airborne")
        
        Fedex_Account_Number=rs_Store("Fedex_Account_Number")
        Fedex_Meter_Number=rs_Store("Fedex_Meter_Number")
        Fedex_Transaction_Identifier=rs_Store("Fedex_Transaction_Identifier")
        Fedex_Carrier_Code=rs_Store("Fedex_Carrier_Code")
        FedEx_DropTypeValue = rs_store("Fedex_DropoffType")
        if FedEx_DropTypeValue = "REGULARPICKUP" then
           FedEx_DropTypeValue_Display="Regular Pickup"
        elseif FedEx_DropTypeValue="REQUESTCOURIER" then
           FedEx_DropTypeValue_Display="Request Courier"
        elseif FedEx_DropTypeValue="STATION" then
           FedEx_DropTypeValue_Display="Station Dropoff"
        end if


rs_Store.Close

if Show_Shipping <> 0	then
	 Show_Shipping_Checked = "Checked"
else
	 Show_Shipping_Checked = ""
end if
if Ship_Multi <> 0 then
	 Ship_Multi = "checked"
else
	 Ship_Multi = ""
end if  

if Use_UPS=-1 then
	USE_UPS_checked = "checked"
end if
if Use_USPS=-1 then
	USE_USPS_checked = "checked"
end if
if Use_DHL=-1 then
	USE_DHL_checked = "checked"
end if
if Use_Fedex=-1 then
	USE_FEDEX_checked = "checked"
end if
if Use_Canada=-1 then
	USE_CANADA_checked = "checked"
end if
if Use_Airborne=-1 then
	USE_AIRBORNE_checked = "checked"
end if
if instr(Shipping_Class,"1")>0 then
	checked_Shipping_Class_1 = "checked"
end if
if instr(Shipping_Class,"2")>0 then
	checked_Shipping_Class_2 = "checked"
end if
if instr(Shipping_Class,"3")>0 then
	checked_Shipping_Class_3 = "checked"
end if
if instr(Shipping_Class,"4")>0 then
	checked_Shipping_Class_4 = "checked"
end if
if instr(Shipping_Class,"5")>0 then
	checked_Shipping_Class_5 = "checked"
end if
if instr(Shipping_Class,"6")>0 then
	checked_Shipping_Class_6 = "checked"
end if
if instr(Shipping_Class,"7")>0 then
	checked_Shipping_Class_7 = "checked"
end if

sNeedTabs=1
sFormAction = "Store_Settings.asp"
sName = "Store_Shipping_class"
sFormName = "Shipping_Class"
sCommonName="Shipping Settings"
sTitle = "Shipping Settings"
sFullTitle = "Shipping > Settings"
sSubmitName = "Update_Store_Shipping_class"
thisRedirect = "shipping_class_realtime.asp"
sTopic="Shipping"
sMenu = "shipping"
createHead thisRedirect

%>



<tr bgcolor='#FFFFFF'><td width="724" align=center valign=top height=35>
<table border=0 cellspacing=0 cellpadding=0 width=724>
			 
	<!-- TAB MENU STARTS HERE -->

		<tr>
		<td align="center" valign=top height=45 width='100%'><br>
		<script type="text/javascript" language="JavaScript1.2" src="include/tabs-xp.js"></script>
		<script language="javascript1.2">
		var bselectedItem   = 0;
		var bmenuItems =
		[
		["Advanced Shipping Options", "content3",,,"Advanced Shipping Options","0"],
		["Realtime Shipping", "content1",,,"Realtime Shipping","0"],
		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
		
		
<!-- CONTENT 1 STARTS HERE -->
			<div id="content1" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>


				<TR bgcolor='#FFFFFF'><td class="inputname" width="40%"><b>Enable Realtime Rates</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class_Real" <%= checked_Shipping_Class_7 %> type="checkbox" value="7">
				<% small_help "Real Shipping Rates" %></td></tr>
				


						<TR bgcolor='#FFFFFF'>
							<td width="20%" class=inputname><B>Use UPS</b></font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="Use_UPS" value="-1" <%= USE_UPS_checked %>>
                        <a href=ups.asp<%= sAddString %> class=link>Instructions for UPS setup</a>
                        <% small_help "Use UPS" %></td>
						</tr>
						
						<TR bgcolor='#FFFFFF'>
							<td width="20%" class=inputname>Username</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="20" name="UPS_User" value="<%= UPS_User %>">
                        <% small_help "UPS Username" %></td>
						</tr>
					
						<TR bgcolor='#FFFFFF'>
							<td width="20%" class=inputname>Password</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="20" name="UPS_Password" value="<%= UPS_Password %>">
                        <% small_help "UPS Password" %></td>
						</tr>
						
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Access License</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="20" name="UPS_AccessLicense" value="<%= UPS_AccessLicense %>">
                        <% small_help "UPS License" %></td>
						</tr>

						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Account Number</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="6" name="UPS_Account_Number" maxlength="8" value="<%= UPS_Account_Number %>">
                        <% small_help "UPS Account Number" %></td>
						</tr>
						

						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Shipper Name</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="35" name="UPS_Shipper_Name" maxlength="35" value="<%= UPS_Shipper_Name %>">
                        <% small_help "Shipper Name" %></td>
						</tr>
					
			<!--   other shipper info relevant fields

						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Address Line 1</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Address1" maxlength="35" value="<%= UPS_Shipper_Address1 %>">
                        <% small_help "UPS Shipper Address Line 1" %></td>
						</tr>

						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Address Line 2</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Address2" maxlength="35" value="<%= UPS_Shipper_Address2 %>">
                        <% small_help "UPS Shipper Address Line 2" %></td>
						</tr>

						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>City</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_City" maxlength="35" value="<%= UPS_Shipper_City %>">
                        <% small_help "UPS Shipper City" %></td>
						</tr>


						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>State</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_State" maxlength="35" value="<%= UPS_Shipper_State %>">
                        <% small_help "UPS Shipper State" %></td>
						</tr>

						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Country</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Country" maxlength="35" value="<%= UPS_Shipper_Country %>">
                        <% small_help "UPS Shipper Country" %></td>
						</tr>
						
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Zip Code</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="10" name="UPS_Shipper_ZipCode" maxlength="10" value="<%= UPS_Shipper_ZipCode %>">
                        <% small_help "UPS Zip Code" %></td>
						</tr>
			
				-->


						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Pickup</font></td>
							<td width="80%" class=inputvalue >
								<select name="UPS_Pickup">
								<option value='<%=UPS_Pickup %>'><%= UPS_Pickup %></option>
                <option value='customer counter'>Customer Counter</option>
                <option value='letter center'>Letter Center</option>
                <option value='on call air'>On Call Air</option>
                <option value='one time pickup'>One Time Pickup</option>
                <option value='daily pickup'>Daily Pickup</option>
                <option value='authorized shipping outlet'>Authorized Shipping Outlet</option>
                <option value='air service center'>Air Service Center</option>
                </select>
								<% small_help "UPS Pickup" %>
							</td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Pack</font></td>
							<td width="80%" class=inputvalue >
								<select name="UPS_Pack">
								<option value='<%=UPS_Pack %>'><%= UPS_Pack %></option>
                <option value='Shipper Supplied Packaging'>Shipper Supplied Packaging</option>
                <option value='UPS Letter Envelope'>UPS Letter Envelope</option>
                <option value='Your Packaging'>Your Packaging</option>
                <option value='UPS Tube'>UPS Tube</option>
                <option value='UPS Pak'>UPS Pak</option>
                <option value='UPS Express Box'>UPS Express Box</option>
                <option value='UPS 25KG Box'>UPS 25KG Box</option>
                <option value='UPS 10KG Box'>UPS 10KG Box</option>
                </select>
                        <% small_help "UPS Pack" %></td>
						</tr>
					
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname><b>Use USPS</b></font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="Use_USPS" value="-1" <%= USE_USPS_checked %>>
                        <% small_help "Use USPS" %></td>
						</tr>

						


						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname><b>Use Fedex</b></font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="Use_FedEx" value="-1" <%= USE_FEDEX_checked %>> <a href=fedex_instructions.asp<%= sAddString %> class=link>Instructions for FedEx setup</a>
                        <% small_help "Use Fedex" %></td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname> FedEx Account Number</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="20" name="Fedex_Account_Number" maxlength="18" value="<%= Fedex_Account_Number %>">
                        <% small_help "Fedex Account Number" %></td>
						</tr>

							<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname> Transaction Identifier</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="20" name="Fedex_Transaction_Identifier" maxlength="18" value="<%= Fedex_Transaction_Identifier %>">
                        <% small_help "Transaction Identifier" %></td>
						</tr>
							<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>FedEx Meter Number</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="20" name="Fedex_Meter_Number" maxlength="18" value="<%= Fedex_Meter_Number %>">
                        <% small_help "Fedex Meter Number" %></td>
						</tr>

							<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>FedEx carrier code</font></td>
							<td width="80%" class=inputvalue >
								<input type="text" size="10" name="Fedex_Carrier_Code" maxlength="18" value="<%= Fedex_Carrier_Code %>">
                        <% small_help "Fedex carrier Code" %></td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Pack</font></td>
							<td width="80%" class=inputvalue >
								<select name="Fedex_Pack">
								<option value='<%=Fedex_Pack %>'><%= Fedex_Pack %></option>
                <option value='FedEx Express Envelope'>FedEx Express Envelope</option>
                <option value='FedEx Express Pak'>FedEx Express Pak</option>
                <option value='FedEx Express Box'>FedEx Express Box</option>
                <option value='FedEx Express Tube'>FedEx Express Tube</option>
                <option value='Your Packaging'>Your Packaging</option>
                </select>
                        <% small_help "Fedex Pack" %></td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							<td width="20%" class=inputname>Dropoff Type</font></td>
							<td width="80%" class=inputvalue >
								<select name="FedEx_DropType">
								<option value='<%=FedEx_DropTypeValue %>'><%=FedEx_DropTypeValue_Display%></option>
                <option value='REGULARPICKUP'>Regular Pickup</option>
                <option value='REQUESTCOURIER'>Request Courier</option>
                <option value='STATION'>Station</option>
                </select>
                            <% small_help "Fedex Pack" %></td>
						</tr>

						<TR bgcolor='#FFFFFF'>

							
							<td width="20%" class=inputname>Ground</font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="Fedex_Ground" value="-1" <%= Fedex_Ground %>>
                        <% small_help "Fedex Ground" %></td>
						</tr>
      <!--   other shipper info relevant fields
      <TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname><b>Use Airborne</b></font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="Use_AIRBORNE" value="-1" <%= USE_AIRBORNE_checked %>>
                        <% small_help "Use Airborne" %></td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname><b>Use DHL</b></font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="Use_DHL" value="-1" <%= Use_DHL_checked %>>
                        <% small_help "Use DHL" %></td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Service</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<select name="DHL_Service">
								<option value='<%=DHL_Service %>'><%= DHL_Service %></option>
                <option value='Second Day'>Second Day</option>
                <option value='Ground'>Ground</option>
                <option value='Express'>Express</option>
                <option value='Next Afternoon'>Next Afternoon</option>
                <option value='Early Delivery'>Early Delivery</option>
                </select>
                        <% small_help "DHL Service" %></td>
						</tr>
						-->
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname><b>Use Canada Post</b></font></td>
							<td class=inputvalue>
								<input class=image type="Checkbox" name="USE_CANADA" value="-1" <%= USE_CANADA_checked %>>
                        <a href=canada_post_instructions.asp<%= sAddString %> class=link>Instructions for Canada Post setup</a><% small_help "Use Canada" %></td>
						</tr>
						<TR bgcolor='#FFFFFF'>
							
							
							<td width="20%" class=inputname>Login</font></td>
							<td width="80%" class=inputvalue >&nbsp;
								<input type="text" size="20" name="Canada_Login" value="<%= Canada_Login %>">
                        <% small_help "Canada Login" %></td>
						</tr>

						






          <TR bgcolor='#FFFFFF'>
					
					<td width="20%"class="inputname"><B>Max Weight Per Box</b></td>
					<td width="80%" class="inputvalue">
							<input id=maxweight type=text name="Max_Weight" onKeyPress="return goodchars(event,'0123456789.')" value=<%= Max_Weight %>>
						<input type="hidden" name="Max_Weight_C" value="Re|Integer|0|150|||Max Weight">
						<% small_help "Max_Weight" %></td>
					</tr>
          <TR bgcolor='#FFFFFF'>
					
					<td width="20%"class="inputname"><B>Restricted Options</b></td>
					<td width="80%" class="inputvalue">
							<textarea id= ro name="Restrict_Options" rows=5 cols=40><%= Restrict_Options %></textarea>
						<input type="hidden" name="Restrict_Options_C" value="Op|String|0|2000|||Restricted Options">
						<% small_help "Name" %></td>
					</tr>

					<TR bgcolor='#FFFFFF'>

					<td align="left" width="188" height="20" class="inputname"><B>Use Realtime Shipping for</b></td>
					<td align="left" width="186" height="20" class="inputvalue">
						<select multiple name="Countries" size="6">
							<Option value="">Select countries where realtime shipping applies
							<%
							sql_region = "SELECT Country,Country_Id FROM Sys_Countries ORDER BY Country;"
							set myfields1=server.createobject("scripting.dictionary")
							Call DataGetrows(conn_store,sql_region,mydata1,myfields1,noRecords1)
							if noRecords1 = 0 then
							FOR rowcounter1= 0 TO myfields1("rowcount")
								' set the selected flag

									CountriesArray=split(Countries,",")
									Found=false
									For i=0 to ubound(CountriesArray)
										CountryElem=replace(CountriesArray(i)," ","")
										Country=replace(mydata1(myfields1("country"),rowcounter1)," ","")
										If CountryElem = Country  Then
											Found=true
											exit for
										End If
									Next
									if Found then
										selected = "selected"
									else
										selected=""
									end if

								response.write "<Option value='"&mydata1(myfields1("country"),rowcounter1)&"' "&selected&" >"&mydata1(myfields1("country"),rowcounter1)&"</option>"
							Next
							End If
							%>
						</select>
						<input type="hidden" name="Countries_C" value="Re|String|0|4000|||Countries">
						<% small_help "Countries" %></td>
					</tr>
</table>
			</div>		
			<!-- CONTENT 1 ENDS HERE -->
			

<!-- CONTENT 3 STARTS HERE -->
			<div id="content3" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Shipping</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_Shipping" value="-1" <%= Show_Shipping_Checked %>>&nbsp;
				<% small_help "Show Shipping" %></td>
				</tr>
                                <TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Allow Separate Shipments</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Ship_Multi" value="-1" <%= Ship_Multi %>>&nbsp;
				<% small_help "Ship Multi" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Handling Fee</td>
				<td width="70%" class="inputvalue"><%= Store_Currency %><input type="text" name="Handling_Fee" value="<%= sHandling_Fee %>" size=5 onKeyPress="return goodchars(event,'-0123456789.')">&nbsp;
				<input type="hidden" name="Handling_Fee_C" value="Re|Integer|||||Handling Fee"><% small_help "Handling_Fee" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Handling Weight</td>
				<td width="70%" class="inputvalue"><input type="text" name="Handling_Weight" value="<%= sHandling_Weight %>" size=5 onKeyPress="return goodchars(event,'-0123456789.')">&nbsp;
				<input type="hidden" name="Handling_Weight_C" value="Re|Integer|||||Handling Weight"><% small_help "Handling_Weight" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td class="inputname" width="40%">Enable Flat fee</td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_1 %> type="checkbox" value="1">
				<a href="shipping_all_class.asp?type=1" class=link>Add Flat Fee Shipping Method</a>
				<% small_help "Flat fee" %></td></tr>

				<TR bgcolor='#FFFFFF'><td class="inputname" width="40%">Enable Flat Fee + Weight</td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_2 %> type="checkbox" value="2">
				<a href="shipping_all_class.asp?type=2" class=link>Add Flat Fee + weight Shipping Method</a>
				<% small_help "Flat Fee Plus Weight" %></td></tr>

				<TR bgcolor='#FFFFFF'><td class="inputname" width="40%">Enable Per Item</td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_3 %> type="checkbox" value="3">
				<a href="shipping_all_class.asp?type=3" class=link>Add Per Item Shipping Method</a>
				<% small_help "No Shipping (Per Item Only)" %></td></tr>

				<TR bgcolor='#FFFFFF'><td class="inputname" width="40%">Enable Total Order Matrix</td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_4 %> type="checkbox" value="4">
				<a href="shipping_all_class.asp?type=4" class=link>Add Total Order Matrix Shipping Method</a>
				<% small_help "Total Order Matrix" %></td></tr>

				<TR bgcolor='#FFFFFF'><td class="inputname" width="40%">Enable % of Total Order</td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_5 %> type="checkbox" value="5">
				<a href="shipping_all_class.asp?type=5" class=link>Add % of Total Order Shipping Method</a>
				<% small_help "Percentage of Total Order" %></td></tr>

				<TR bgcolor='#FFFFFF'><td class="inputname" width="40%">Enable Total Weight Matrix</td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_6 %> type="checkbox" value="6">
				<a href="shipping_all_class.asp?type=6" class=link>Add Total Weight Matrix Shipping Method</a>
				<% small_help "Total Weight Matrix" %></td></tr>

			</table>
			</div>		
			<!-- CONTENT 3 ENDS HERE -->
			</td>
		</tr>	
			 
</table></td></tr>
<tr bgcolor=#FFFFFF>
<td colspan='4' class='tpage'>		
<% createFoot thisRedirect,1 %>
</td>
</tr>

<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Countries","req","Please select the countries that realtime shipping will apply to.");
 frmvalidator.addValidation("Max_Weight","req","Please enter a maximum weight.");
 frmvalidator.addValidation("Handling_Fee","req","Please enter a handling fee");
 frmvalidator.addValidation("Handling_Weight","req","Please enter a handling weight");

</script>
