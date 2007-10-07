<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

'SELECT ITEM TO QUICK EDIT
Calendar=1
sFormAction = "customer_quick_edit_action.asp"
sTitle = "Quick Edit Customers"
sSubmitName = "Quick_Edit_Update"
thisRedirect = "quick_edit_customers.asp"
sFormName = "Quick_Edit_Update"
addPicker =1
sMenu = "inventory"
createHead thisRedirect
if Service_Type < 7 then %>
	<tr>
	<td colspan=2>
		This feature is not available at your current level of service.
	</td></tr>

<% createFoot thisRedirect, 0%>
<% else

	select_str="SELECT * FROM store_customers WHERE Store_id="&Store_id&" and record_type= 0 AND  ccid in ("&request.form("DELETE_IDS")&")" 
	set customerFields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,select_str,customerData,customerFields,noRecords)

	
%>
<input type="hidden" value="<%=request.form("DELETE_IDS")%>" name= "DELETE_IDS">
<%if noRecords= 0 then%>


	<%
	for customerRowCounter = 0 to customerFields("rowcount")
		
		Disp_customer_ID = customerData(customerFields("ccid"),customerRowcounter) 
		customer_ID = customerData(customerFields("cid"),customerRowcounter) 

		spam = customerData(customerFields("spam"),customerRowcounter) 
		Protected_page_access = 	customerData(customerFields("protected_page_access"),customerRowcounter)  
		Tax_Exempt =customerData(customerFields("tax_exempt"),customerRowcounter)    


		%>

					<tr>
						<td colspan="3"><b><br>Customer #<%= disp_customer_ID %><br></b></td>
					</tr>
					<tr>
						<td width="20%" class="inputname">User Name</td>
						<td width="80%" class="inputvalue">
									<input name="user_id_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("user_id"),customerRowcounter) %>">
									<input type="hidden" name="user_id_<%= disp_customer_ID %>_C" value="Re|String|0|200|||user_id">
									<% small_help "UserID" %></td>

					</tr>


					<tr>
						<td width="20%" class="inputname">Password</td>
						<td width="80%" class="inputvalue">
									<input name="password_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("password"),customerRowcounter) %>">
									<input type="hidden" name="password_<%= disp_customer_ID %>_C" value="Re|String|0|200|||Password">
									<% small_help "Password" %></td>
					</tr>

					<tr>
						<td width="20%" class="inputname">Budget</td>
						<td width="80%" class="inputvalue">
									<input name="budget_orig_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("budget_orig"),customerRowcounter) %>">
									<input type="hidden" name="budget_orig_<%= disp_customer_ID %>_C" value="Re|String|0|200|||Budget">
									<% small_help "Budget" %></td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Budget Left</td>
						<td width="80%" class="inputvalue">
									<input name="budget_left_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("budget_left"),customerRowcounter) %>">
									<input type="hidden" name="budget_left_<%= disp_customer_ID %>_C" value="Re|String|0|200|||budget_left">
									<% small_help "Budget" %></td>
					</tr>

						
					<tr>
						<td class="inputname">Promo Emails </td>
						<td class="inputvalue">
						<%if customerData(customerFields("spam"),customerRowcounter) then %>
								<input class="image" type="checkbox"  checked  name="Spam_<%= disp_customer_ID %>" >
						<%else%>
								<input class="image" type="checkbox"  name="Spam_<%= disp_customer_ID %>" >
						<%end if%>
						
						
						<% small_help "Promo Emails" %></td>
					</tr>
						
					
						
					<tr>
						<td class="inputname">Tax Exempt </td>
						<td class="inputvalue">
						<% 
						if customerData(customerFields("tax_exempt"),customerRowcounter)=true  then   
						%>
										<input class="image" type="checkbox" value="a" checked name="Tax_Exempt_<%= disp_customer_ID %>">
						<%
						else
						%>
										<input class="image" type="checkbox" name="Tax_Exempt_<%= disp_customer_ID %>">
						<%
						end if
						%>
						
						<% small_help "Tax Exempt" %></td>
					</tr>


					<tr>
						<td class="inputname">Protected Access</td>
						<td class="inputvalue">
						<% 
						if customerData(customerFields("protected_page_access"),customerRowcounter) = true   then
						%> 
								<input class="image" type="checkbox" value="a"  checked name="protected_page_access_<%= disp_customer_ID %>"  >
						<%
						else
						%>
								<input class="image" type="checkbox"  name="protected_page_access_<%= disp_customer_ID %>"  >
						<%
						end if
						%>
						
						<% small_help "Protected Access" %></td>
					</tr>

					<tr>
						<td colspan = "3" >&nbsp;</td>
					</tr>

					<tr>
						<td colspan = "3" ><b>Billing Information</b></td>
					</tr>

					<tr>
						<td colspan = "3" >&nbsp;</td>
					</tr>

					<tr>
						<td width="20%" class="inputname">First Name</td>
						<td width="80%" class="inputvalue">
									<input name="first_name_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("first_name"),customerRowcounter) %>">
									<input type="hidden" name="first_name_<%= disp_customer_ID %>_C" value="Re|String|0|200|||first_name">
									<% small_help "FirstName" %></td>
					</tr>

					<tr>
						<td width="20%" class="inputname">Last Name</td>
						<td width="80%" class="inputvalue">
									<input name="last_name_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("last_name"),customerRowcounter) %>">
									<input type="hidden" name="last_name_<%= disp_customer_ID %>_C" value="Re|String|0|200|||last_name">
									<% small_help "LastName" %></td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Address Line 1</td>
						<td width="80%" class="inputvalue">
									<input name="address1_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("address1"),customerRowcounter) %>">
									<input type="hidden" name="Address1_<%= disp_customer_ID %>_C" value="Re|String|0|200|||address1">
									<% small_help "Address" %></td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Address Line 2</td>
						<td width="80%" class="inputvalue">
									<input name="address2_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("address2"),customerRowcounter) %>">
									<input type="hidden" name="address2_<%= disp_customer_ID %>_C" value="Op|String|0|200|||address2">
									<% small_help "Address" %></td>
					</tr>





					<tr>
						<td width="20%" class="inputname">City</td>
						<td width="80%" class="inputvalue">
									<input name="city_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("city"),customerRowcounter) %>">
									<input type="hidden" name="city_<%= disp_customer_ID %>_C" value="Re|String|0|200|||city">
									<% small_help "City" %></td>
					</tr>

					<tr>
						<td width="20%" class="inputname">State</td>
						<td width="80%" class="inputvalue">
									<input name="state_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("state"),customerRowcounter) %>">
									<input type="hidden" name="sate_<%= disp_customer_ID %>_C" value="Re|String|0|200|||state">
									<% small_help "State" %></td>
					</tr>

					<tr>
					<td width="20%" class="inputname">Country</td>
						<td width="80%" class="inputvalue">
							<select name="country_<%= disp_customer_ID %>" size="1">
							<% 
							'getting the country name of the customer
							 Country = customerData(customerFields("country"),customerRowcounter)
							 'getting the list of countries
							sql_region = "SELECT Country FROM Sys_Countries ORDER BY Country;"
							set myfields=server.createobject("scripting.dictionary")
							Call DataGetrows(conn_store,sql_region,mydata,myfields,noRecords)
								if noRecords = 0 then
								FOR rowcounter= 0 TO myfields("rowcount")
									' set the selected flag
									if Country=mydata(myfields("country"),rowcounter) then
										selected = "selected"
									ELSE
										selected = ""
									end if
									response.write "<Option "&selected&" value='"&mydata(myfields("country"),rowcounter)&"'>"&mydata(myfields("country"),rowcounter)&"</option>"
								Next
								End If %>
							</select>
							</td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Zipcode</td>
						<td width="80%" class="inputvalue">
									<input name="zip_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("zip"),customerRowcounter) %>">
									<input type="hidden" name="zip_<%= disp_customer_ID %>_C" value="Re|String|0|200|||Zip">
									<% small_help "Zip" %></td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Phone</td>
						<td width="80%" class="inputvalue">
									<input name="phone_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("phone"),customerRowcounter) %>">
									<input type="hidden" name="phone_<%= disp_customer_ID %>_C" value="Re|String|0|200|||phone">
									<% small_help "Phone" %></td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Email</td>
						<td width="80%" class="inputvalue">
									<input name="email_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("email"),customerRowcounter) %>">
									<input type="hidden" name="email_<%= disp_customer_ID %>_C" value="Re|String|0|200|||email">
									<% small_help "Email" %></td>
					</tr>


					<tr>
						<td width="20%" class="inputname">Fax</td>
						<td width="80%" class="inputvalue">
									<input name="fax_<%= disp_customer_ID %>" size="30" value="<%= customerData(customerFields("fax"),customerRowcounter) %>">
									<input type="hidden" name="fax_<%= disp_customer_ID %>_C" value="op|String|0|200|||fax">
									<% small_help "Fax " %></td>
					</tr>

<tr>
	<td colspan = "3">
		<hr color=#000000 size=1>
	</td>
</tr>

	<%
	next
	%>
<%		
end if
%>

<% createFoot thisRedirect, 1%>
<% end if %>
