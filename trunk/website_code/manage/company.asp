<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/country_list.asp"-->
<!--#include file="include/List_of_countries.asp"-->
<!--#include file="help/company.asp"-->
<%

if Newsletter_receive <> 0 then
	newsletter_checked = "checked"
else
	newsletter_checked = ""
end if

if show_tax_exempt <> 0 then
	exempt_checked = "checked"
else
	exempt_checked = ""
end if

if show_residential <> 0 then
	residential_checked = "checked"
else
	residential_checked = ""
end if

if show_special_offers <> 0 then
	special_checked = "checked"
else
	special_checked = ""
end if

if Store_public <>	0	then
	 Store_public_Checked = "Checked"
else
	 Store_public_Checked = ""
end if

if Gift_Service <>	0	then
	 Gift_Service_Checked = "Checked"
	 sDivGift="block"
else
	 Gift_Service_Checked = ""
	 sDivGift="none"
end if

if Gift_Message <>	0	then
	 Gift_Message_Checked = "Checked"
else
	 Gift_Message_Checked = ""
end if

if ExpressCheckout <> 0 then
	 express_checked = "checked"
else
	 express_checked = ""
end if

if Enable_IP_Tracking <> 0 then
	 ip_checked = "checked"
else
	 ip_checked = ""
end if

if No_Login <> 0 then
	 No_Login = "checked"
else
	 No_Login = ""
end if

if AllowCookies <> 0 then
	 cookies_checked = "checked"
else
	 cookies_checked = ""
end if

if When_Adding <> 0	then
	 When_Adding_Checked = "Checked"
else
	 When_Adding_Checked = ""
end if

if Save_Cart <>	0	then
	 Save_Cart_Checked = "Checked"
else
	 Save_Cart_Checked = ""
end if

if hide_retail_price<>0 then
	Show_Hide_Retail_Price_Checked = "Checked"	
else
	Show_Hide_Retail_Price_Checked=""
end if

if Cart_Thumbnails<>0 then
	Cart_Thumbnails_Checked = "Checked"
else
	Cart_Thumbnails_Checked=""
end if

sFormAction = "store_settings.asp"
sName = "Store_Company_info"
sFormName = "company"
sQuestion_Path = "general/store_info.htm"

sNeedTabs=1
sTitle = "Settings"
sCommonName = "Settings"
sFullTitle = "General > Settings"
sSubmitName = "Update_Store_Company"
thisRedirect = "company.asp"
sMenu = "general"
sTopic = "Store_Info"
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
		["Contact Info", "content1",,,"Contact Info","Contact Info"],
		["Login Options", "content2",,,"Login Options","Login Options"],
		["Cart Options ",  "content7",,,"Cart Options","Cart Options"],
		];

		apy_tabsInit();
		</script>
		</td>
		</tr>
		
		<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='25'>
		
		
		<!-- CONTENT 1 MAIN -->
			<div id="content1" style="visibility: hidden;" class=tpage>
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				
				
			
				<tr bgcolor=#FFFFFF>
					<td width="24%" height="23" class="inputname">Store Name</td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_name" value="<%= Store_name %>" size="60" maxlength=200>*
							<input type="hidden" name="Store_name_C" value="Re|String|0|200|||Store Name">
							<% small_help "Store Name" %></td>
					</tr>
					

			<tr bgcolor=#FFFFFF>
					<td width="24%" height="23" class="inputname">Company Name</td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_company" value="<%= Store_company %>" size="60" maxlength=250>
					<input type="hidden" name="Store_company_C" value="Op|String|0|250|||Company Name">
					<% small_help "Company Name" %></td>
					</tr>
		
					<tr bgcolor=#FFFFFF>
					<td width="24%" height="23" class="inputname">Address Line 1</td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_address1" value="<%= Store_address1 %>" size="60" maxlength=250>
					<input type="hidden" name="Store_address1_C" value="Op|String|0|250|||Address Line 1">
					<% small_help "Address Line 1" %></td>
					</tr>

					<tr bgcolor=#FFFFFF>
					<td width="24%" height="23" class="inputname">Address Line 2</td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_address2" value="<%= Store_address2 %>" size="60" maxlength=250>
					<input type="hidden" name="Store_address2_C" value="Op|String|0|250|||Address Line 2">
					<% small_help "Address Line 2" %></td>
					</tr>

					<tr bgcolor=#FFFFFF>
					<td width="24%" height="23" class="inputname">City</td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_city" value="<%= Store_city %>" size="60" maxlength=100>
					<input type="hidden" name="Store_city_C" value="Op|String|0|100|||City">
					<% small_help "City" %></td>
					</tr>
					<tr bgcolor=#FFFFFF>
						<td width="24%" height="23" class="inputname">Country</td>
					<td width="76%" height="23" class="inputvalue">
							 
							 <select id='Store_Country' name='Store_Country' onChange="validateState()">
										<option selected="selected" value='<%= Store_Country %>'><%= Store_Country %></option>
										<option><%=create_country_list ("Store_country",Store_country,1,"") %></option>												
							 </select>			
							
							<% small_help "Country" %></td>
					</tr>
                <tr bgcolor="#ffffff">
                    <td class="inputname" height="23" width="24%">
                        State</td>
                    <td class="inputvalue" height="23" width="76%">
			
			
								<%
								   
									set rsState = server.CreateObject("Adodb.Recordset")
									
								   	sql_STMT_USA = "SELECT State, State_Long FROM Sys_States WHERE Country = 'United States' order by State_Long asc"
									sql_STMT_Canada = "SELECT State, State_Long FROM Sys_States WHERE Country = 'Canada' order by State_Long asc"
								

								   if Store_Country = "United States" then
								    Div1="block"									
								    Div2="none"
								    Div3="none"	
									

								rsState.open sql_STMT_USA,conn_store,1,1									
									if not(rsState.eof or rsState.bof) then
										do while not rsState.eof
											if rsState.fields("State") = Store_State then
												USA_State_Value = rsState.fields("State")
												USA_State_Text = rsState.fields("State_Long")
												exit do												
											end if											
										rsState.MoveNext
										loop										
									end if
									rsState.close
																		
									Canada_State = ""
									Other_State=""
									
								   elseif Store_Country = "Canada" then
								    Div1="none"
								    Div2="block"
								    Div3="none"
									USA_State=""
									
									
								rsState.open sql_STMT_Canada,conn_store,1,1									
									if not(rsState.eof or rsState.bof) then
										do while not rsState.eof
											if rsState.fields("State") = Store_State then
												Canada_State_Value = rsState.fields("State")
												Canada_State_Text = rsState.fields("State_Long")
												exit do												
											end if											
										rsState.MoveNext
										loop										
									end if
									rsState.close									
									Other_State=""
								   
								   else
								    Div1="none"
								    Div2="none"
								    Div3="block"	
									USA_State=""
									Canada_State = ""
									Other_State = Store_State								   
								   end if
								   %>
								   
								   
								   <div id="UnitedStatesDiv" style="display: <%= Div1 %>">								   
								   	<select id='State_UnitedStates' name='State_UnitedStates' onChange="storeStateData()">
								   		<option selected="selected" value='<%=USA_State_Value%>'><%=USA_State_Text%></option>
										<%
										rsState.open sql_STMT_USA,conn_store,1,1
											if not (rsState.eof or rsState.bof) then
												do while not rsState.eof
										%>
												<option value="<%=rsState.fields("State")%>"><%=rsState.fields("State_Long")%></option>
										<%
												rsState.MoveNext
												loop
											end if
										rsState.close		
										
										%>		
									</select>*								
							</div>
						
									<div id="CanadaDiv" style="display: <%= Div2 %>">
									<select id='State_Canada' name='State_Canada' onChange="storeStateData()">
										<option selected="selected" value='<%=Canada_State_Value%>'><%=Canada_State_Text%></option>
										<%
										rsState.open sql_STMT_Canada,conn_store,1,1
											if not (rsState.eof or rsState.bof) then
												do while not rsState.eof
										%>
												<option value="<%=rsState.fields("State")%>"><%=rsState.fields("State_Long")%></option>
										<%
												rsState.MoveNext
												loop
											end if
										rsState.close		
										'set rsState = nothing
										%>																
								  </select>*
								  </div>

												
									<div id="optDiv" style="display: <%= Div3 %>">
										<input type="text" id='txtStore_State' name='txtStore_State' value="<%= Other_State %>" size="40" maxlength=100 onBlur="storeStateData()">
							        </div>
		
							
							<input type="hidden" id="Store_State" name="Store_State" value="<%=Store_State%>">
							<input type="hidden" name="Store_State_C" value="Op|String|0|100|||State">
					
					
					<% small_help "State" %>
                    </td>
                </tr>
		  

		
					<tr bgcolor=#FFFFFF>
				<td width="24%" height="23" class="inputname">Zip Code</td>
				<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_zip" value="<%= Store_zip %>" size="60" maxlength=10>
							<INPUT type="hidden"  name=Store_zip_C value="Op|String|0|14|||Zip Code">
							<% small_help "Zip Code" %></td>
					</tr>
		
					<tr bgcolor=#FFFFFF>
				<td width="24%" height="23" class="inputname">Phone</td>
					<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_Phone" value="<%= Store_Phone %>" size="60" maxlength="20">
					<INPUT type="hidden"  name=Store_Phone_C value="Op|String|0|20|||Phone">
					<% small_help "Phone" %></td>
						</tr>
		
					<tr bgcolor=#FFFFFF>
				<td width="24%" height="23" class="inputname">Fax</td>
				<td width="76%" height="23" class="inputvalue">
							<input type="text" name="Store_Fax" value="<%= Store_Fax %>" size="60" maxlength="20">
					<INPUT type="hidden"  name=Store_Fax_C value="Op|String|0|20|||Fax">
					<% small_help "Fax" %></td>
					</tr>

					<tr bgcolor=#FFFFFF>
				<td width="24%" height="17" class="inputname">Email</td>
				<td width="76%" height="17" class="inputvalue">
							<input type="text" name="Store_email" value="<%= Store_email %>" size="60" maxlength=50>*
					<INPUT type="hidden"  name=Store_email_C value="Re|Email|0|50|@,.||Email">
					<% small_help "Email" %></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Receive Newsletter</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Newsletter_Receive" value="-1" <%= newsletter_checked %>>&nbsp;
					 <% small_help "Receive newsletter" %></td>
				</tr>
			</table>
			</div>
		<!-- CONTENT 2 MAIN -->
		
		
		
		
		
		<!-- CONTENT 2 STARTS HERE -->
			<div id="content2" style="visibility: hidden;" class="tabPage">
			
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Public Store</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Store_Public" value="-1" <%= Store_Public_checked %>>&nbsp;
					 <% small_help "Public Store" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Skip Login Page</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="No_Login" value="-1" <%= No_Login %>>&nbsp;
					<% small_help "No Login" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Allow Cookies</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="AllowCookies" value="-1" <%= cookies_checked %>>&nbsp;
					 <% small_help "Allow Cookies" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Express Checkout</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="ExpressCheckout" value="-1" <%= express_checked %>>&nbsp;
					<% small_help "Members Only" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Enable IP Tracking</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Enable_IP_Tracking" value="-1" <%= ip_checked %>>&nbsp;
					<% small_help "IP Tracking" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Residential</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_Residential" value="-1" <%= residential_checked %>>&nbsp;
					<% small_help "Show_Residential" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Special Offers</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_Special_Offers" value="-1" <%= special_checked %>>&nbsp;
					<% small_help "Show_Special_Offers" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Tax Exempt</td>
				<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_Tax_Exempt" value="-1" <%= exempt_checked %>>&nbsp;
					<% small_help "Show_Tax_Exempt" %></td>
				</tr>
			</table>
			</div>		
			<!-- CONTENT 2 ENDS HERE -->
			
			
			
			<!-- CONTENT 7 STARTS HERE -->
			<div id="content7" style="visibility: hidden;" class="tabPage">
			<table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>

				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show/Hide Cart When Adding</td>
				<td width="70%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="When_Adding" value="-1" <%= When_Adding_Checked %>>&nbsp;
				<% small_help "When Adding" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Save / Retrieve Cart</td>
				<td width="70%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="Save_Cart" value="-1" <%= Save_Cart_Checked %>>&nbsp;
				<% small_help "Show_SaveCart" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Hide Retail Price in Cart</td>
				<td width="70%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="Hide_Retail_Price" value="-1" <%= Show_Hide_Retail_Price_Checked %>>&nbsp;
				<% small_help "Hide_Retail" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Show Thumbnails in Cart</td>
				<td width="70%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="Cart_Thumbnails" value="-1" <%= Cart_Thumbnails_Checked %>>&nbsp;
				<% small_help "Cart_Thumbnails" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Offer Gift Wrapping</td>
				<td width="35%" class="inputvalue"><input class="image" type="checkbox" name="Gift_Service" value="-1" <%= Gift_Service_Checked %>></td>
				<td width="35%" class="inputvalue">Additional <%= Store_Currency %>
				<input type="text" size=5 name="Gift_Wrapping_Surcharge" value="<%= Gift_Wrapping_Surcharge %>" onKeyPress="return goodchars(event,'0123456789.')">
				<input type="hidden" name="Gift_Wrapping_Surcharge_C" value="Op|Integer|||||Gift Wrapping Surcharge">
				<% small_help "Gift Wrapping" %></td>
				</tr>
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Offer Gift Message</td>
				<td width="70%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="Gift_Message" value="-1" <%= gift_message_checked %>>&nbsp;
				<% small_help "Gift Message" %></td>
				</tr>
				<tr bgcolor=#FFFFFF>
				<td width="30%" class="inputname">Currency Symbol</td>
				<td width="70%" class="inputvalue" colspan=2><input type="text" name="Store_Currency" value="<%= Store_Currency %>" maxlength=5 size=3>*
				<input type="hidden" name="Store_Currency_C" value="Re|String|0|5|||Currency">
					 <% small_help "Currency" %></td>
				</tr>
            <TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Continue Shopping URL</td>
				<td width="70%" class="inputvalue" colspan=2><input type="text" name="Continue_Shopping" value="<%= Continue_Shopping %>" size=60>&nbsp;
				<input type="hidden" name="Continue_Shopping_C" value="Op|String|0|100|||Continue Shopping"><% small_help "Continue_Shopping" %></td>
				</tr>
            <TR bgcolor='#FFFFFF'>
								<td width="30%" class="inputname">Minimum Order 
                                Amount</td>
				<td width="70%" class="inputvalue" colspan=2> <%= Store_Currency %>
				<input type="text" size=5 name="Minimum_Amount" value="<%= Minimum_Amount %>" onKeyPress="return goodchars(event,'0123456789.')">
								<% small_help "Minimum_Amount" %></td>
				</tr>
			</table>
			</div>
			<!-- CONTENT 7 ENDS HERE -->




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
 frmvalidator.addValidation("Store_email","email","Please enter a valid email.");
 frmvalidator.addValidation("Store_email","req","Please enter a valid email.");
 frmvalidator.addValidation("Store_Currency","req","Please enter a currency."); 
 
 function validateState()
 {
 	if(window.document.getElementById("Store_Country").value=="United States")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="block"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="none" 	       
 	        window.document.getElementById("State_UnitedStates").options[1].selected=true
 	        window.document.getElementById("Store_State").value = window.document.getElementById("State_UnitedStates").value
 	    }
 	else if(window.document.getElementById("Store_Country").value=="Canada")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="block"
 	        window.document.getElementById("optDiv").style.display="none" 
 	        window.document.getElementById("State_Canada").options[1].selected=true
 	        window.document.getElementById("Store_State").value = window.document.getElementById("State_Canada").value	    	    
 	    }
 	else
 	   {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="block"
 	        window.document.getElementById("txtStore_State").value=""
			window.document.getElementById("Store_State").value = ""	    
 	    }
    
 }	    
 
 
 function storeStateData()
    {
    if(window.document.getElementById("Store_Country").value=="United States")
 	    {
            window.document.getElementById("Store_State").value = window.document.getElementById("State_UnitedStates").value
        }   
    else if(window.document.getElementById("Store_Country").value=="Canada")
 	    {
            window.document.getElementById("Store_State").value = window.document.getElementById("State_Canada").value	    
        }
    else
        {    
            window.document.getElementById("Store_State").value = window.document.getElementById("txtStore_State").value	    
        }
    }    	    
 	

</script>

