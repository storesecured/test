<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

'LOAD CURRENT VALUES FROM THE DATABASE

if Use_UPS=-1 then
	USE_UPS = "checked"
end if
if Use_USPS=-1 then
	USE_USPS = "checked"
end if
if Use_DHL=-1 then
	USE_DHL = "checked"
end if
if Use_Fedex=-1 then
	USE_FEDEX = "checked"
end if
if use_canada=-1 then
	USE_CANADA = "checked"
end if
if Use_Conway=-1 then
	USE_CONWAY = "checked"
end if
if Use_Aireborne=-1 then
	USE_AIRBORNE = "checked"
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

if instr(Shipping_Class,"7")>0 then
	 checkedenable = "checked"
	 sDiv="block"
else 
	 checkedenable = ""
	 sDiv="none"
end if

on error resume next
sql_real_time = "select * from Store_Real_Time_Settings where store_id="&store_id
rs_Store.open sql_real_time, conn_store, 1, 1
	UPS_User = rs_store("UPS_User")
	UPS_Password = rs_store("UPS_Password")
	USPS_User = rs_store("USPS_User")
	USPS_Password = rs_store("USPS_Password")
	UPS_AccessLicense = rs_store("UPS_AccessLicense")
	UPS_Account_Number=rs_store("UPS_Account_Number")
	UPS_Shipper_Name=rs_store("UPS_Shipper_Name")
	Canada_Login = rs_store("Canada_Login")
	Conway_Login = rs_store("Conway_Login")
	Conway_Password = rs_store("Conway_Password")
	UPS_Pickup = rs_store("UPS_Pickup")
	UPS_Pack = rs_store("UPS_Pack")
	Fedex_Pack = rs_store("Fedex_Pack")
	Fedex_Ground = rs_store("Fedex_Ground")
	if Fedex_Ground then
	   Fedex_Ground="checked"
	end if
	DHL_Service = rs_store("DHL_Service")
rs_Store.Close

sFormAction = "Store_Settings.asp"
sName = "Store_Shipping_class"
sFormName = "Shipping_Class"
sTitle = "Shipping"
sSubmitName = "Update_Store_Shipping_class"
thisRedirect = "shipping_class.asp"
sTopic="Shipping"
sMenu = "general"
sQuestion_Path = "general/shipping.htm"
createHead thisRedirect

%>
		<tr><td colspan=4 align=right class=HelpAvailable>Help Available
		<a class="link" href=http://videos.storesecured.com/shipping/shipping.htm target=_blank><img src=images/flash.gif border=0></a>
		<a class="link" href=http://videos.storesecured.com/shipping/shipping.wmv target=_blank><img src=images/wm.gif border=0></a>
		<a class="link" href=http://videos.storesecured.com/shipping/shipping.zip target=_blank><img src=images/zip.gif border=0></a>
		</td></tr>
		<tr><td colspan=4 class=instructions>Please select the check box for each of the ways that you would like the store
				to calculate shipping.  You may select more then 1 method of calculating the shipping cost if 
				appropriate for your store.  Your customers will be shown all applicable prices for their 
				cart at checkout and given the choice of which method they would like to use.  Once you have 
				selected the type(s) of shipping to accept make sure you use the 
				details link to setup your pricing structure.</td></tr>

		<tr>
				<td class="inputname" width="40%"><b>Flat fee</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_1 %> type="checkbox" value="1">
				</td><td class=inputvalue><a href=shipping_class1_list.asp<%= sAddString %> class=link>Details</a><% small_help "Flat fee" %></td></tr>

				<tr><td class="inputname" width="40%"><b>Flat fee + weight</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_2 %> type="checkbox" value="2">
				</td><td class=inputvalue><a href=shipping_class2_list.asp<%= sAddString %> class=link>Details</a><% small_help "Flat Fee Plus Weight" %></td></tr>

				<tr><td class="inputname" width="40%"><b>Per Item Shipping</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_3 %> type="checkbox" value="3">
				</td><td class=inputvalue><a href=shipping_class3_list.asp<%= sAddString %> class=link>Details</a><% small_help "No Shipping (Per Item Only)" %></td></tr>

				<tr><td class="inputname" width="40%"><b>Total Order Matrix</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_4 %> type="checkbox" value="4">
				</td><td class=inputvalue><a href=shipping_class4_list.asp<%= sAddString %> class=link>Details</a><% small_help "Total Order Matrix" %></td></tr>

				<tr><td class="inputname" width="40%"><b>% of total order</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_5 %> type="checkbox" value="5">
				</td><td class=inputvalue><a href=shipping_class5_list.asp<%= sAddString %> class=link>Details</a><% small_help "Percentage of Total Order" %></td></tr>

				<tr><td class="inputname" width="40%"><b>Total Weight Matrix</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class" <%= checked_Shipping_Class_6 %> type="checkbox" value="6">
				</td><td class=inputvalue><a href=shipping_class6_list.asp<%= sAddString %> class=link>Details</a><% small_help "Total Weight Matrix" %></td></tr>

			 <% if Service_Type > 5 then %>
				<tr><td class="inputname" width="40%"><b>Real Shipping Rates</b></td>
				<td class="inputvalue" width="60%"><input class="image" name="Shipping_Class_Real" OnClick="showHideForm('real','Shipping_Class_Real');" <%= checked_Shipping_Class_7 %> type="checkbox" value="7">
				</td><td class=inputvalue><a href=realtime_options.asp<%= sAddString %> class=link>Details</a><% small_help "Real Shipping Rates" %></td></tr>

<tr>
				<td width="90%" colspan="4">
				<DIV NAME="real" ID="real" style="display: <%= sDiv %>;">
					<table border="0" width="100%" cellspacing="0" cellpadding="0">
						
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><B>Use UPS</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Use_UPS" value="-1" <%= USE_UPS %>>
                        <a href=ups.asp<%= sAddString %> class=link>Instructions for UPS setup</a>
                        <% small_help "Use UPS" %></td>
						</tr>
						
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Username</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="20" name="UPS_User" value="<%= UPS_User %>">
                        <% small_help "UPS Username" %></td>
						</tr>
					
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Password</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="20" name="UPS_Password" value="<%= UPS_Password %>">
                        <% small_help "UPS Password" %></td>
						</tr>
						
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Access License</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="20" name="UPS_AccessLicense" value="<%= UPS_AccessLicense %>">
                        <% small_help "UPS License" %></td>
						</tr>

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Account Number</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="6" name="UPS_Account_Number" maxlength="8" value="<%= UPS_Account_Number %>">
                        <% small_help "UPS Account Number" %></td>
						</tr>
						

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Shipper Name</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Name" maxlength="35" value="<%= UPS_Shipper_Name %>">
                        <% small_help "Shipper Name" %></td>
						</tr>
					
			<!--   other shipper info relevant fields

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Address Line 1</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Address1" maxlength="35" value="<%= UPS_Shipper_Address1 %>">
                        <% small_help "UPS Shipper Address Line 1" %></td>
						</tr>

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Address Line 2</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Address2" maxlength="35" value="<%= UPS_Shipper_Address2 %>">
                        <% small_help "UPS Shipper Address Line 2" %></td>
						</tr>

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">City</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_City" maxlength="35" value="<%= UPS_Shipper_City %>">
                        <% small_help "UPS Shipper City" %></td>
						</tr>


						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">State</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_State" maxlength="35" value="<%= UPS_Shipper_State %>">
                        <% small_help "UPS Shipper State" %></td>
						</tr>

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Country</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="35" name="UPS_Shipper_Country" maxlength="35" value="<%= UPS_Shipper_Country %>">
                        <% small_help "UPS Shipper Country" %></td>
						</tr>
						
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Zip Code</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="10" name="UPS_Shipper_ZipCode" maxlength="10" value="<%= UPS_Shipper_ZipCode %>">
                        <% small_help "UPS Zip Code" %></td>
						</tr>
			
				-->


						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Pickup</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
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
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Pack</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
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
					
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><b>Use USPS</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Use_USPS" value="-1" <%= USE_USPS %>>
                        <% small_help "Use USPS" %></td>
						</tr>

						


						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><b>Use Fedex</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Use_FedEx" value="-1" <%= USE_FEDEX %>>
                        <% small_help "Use Fedex" %></td>
						</tr>
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Pack</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
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
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Ground</font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Fedex_Ground" value="-1" <%= Fedex_Ground %>>
                        <% small_help "Fedex Ground" %></td>
						</tr>
      <!--   other shipper info relevant fields
      <tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><b>Use Airborne</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Use_AIRBORNE" value="-1" <%= USE_AIRBORNE %>>
                        <% small_help "Use Airborne" %></td>
						</tr>
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><b>Use DHL</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Use_DHL" value="-1" <%= Use_DHL %>>
                        <% small_help "Use DHL" %></td>
						</tr>
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Service</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
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
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><b>Use Canada Post</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="USE_CANADA" value="-1" <%= USE_CANADA %>>
                        Email eparcel@canadapost.ca to request a merchant ID for the XML system.<% small_help "Use Canada" %></td>
						</tr>
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Login</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="20" name="Canada_Login" value="<%= Canada_Login %>">
                        <% small_help "Canada Login" %></td>
						</tr>

						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2"><b>Use Conway</b></font></td>
							<td colspan="3" class=inputvalue>
								<input class=image type="Checkbox" name="Use_CONWAY" value="-1" <%= USE_CONWAY %>>
                        <% small_help "Use Conway" %></td>
						</tr>
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Login</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="20" name="Conway_Login" value="<%= Conway_Login %>">
                        <% small_help "Conway Password" %></td>
						</tr>
						<tr>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap><font face="Arial" size="2">&nbsp;&nbsp;</font></td>
							<td width="1%" nowrap class=inputname><font face="Arial" size="2">Password</font></td>
							<td width="97%" class=inputvalue colspan="3" >&nbsp;
								<input type="text" size="20" name="Conway_Password" value="<%= Conway_Password %>">
                        <% small_help "Conway Password" %></td>
						</tr>



					</table>
					</div>
					<% end if %>
				</td>
			</tr>




<% createFoot thisRedirect,1 %>
