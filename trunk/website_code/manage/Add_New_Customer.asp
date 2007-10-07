<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<script language="javascript" type="text/javascript">
 function validateState()
 {
 	if(window.document.getElementById("Country").value=="United States")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="block"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="none" 	       
 	        //window.document.getElementById("State_UnitedStates").options[1].selected=true
 	        window.document.getElementById("State").value = window.document.getElementById("State_UnitedStates").value
 	    }
 	else if(window.document.getElementById("Country").value=="Canada")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="block"
 	        window.document.getElementById("optDiv").style.display="none" 
 	        //window.document.getElementById("State_Canada").options[1].selected=true
 	        window.document.getElementById("State").value = window.document.getElementById("State_Canada").value	    	    
 	    }
 	else
 	   {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="block"
 	        window.document.getElementById("State").value=""	    
 	    }
 	    
 }	
 
 
  function storeStateData()
    {
    if(window.document.getElementById("Country").value=="United States")
 	    {
            window.document.getElementById("State").value = window.document.getElementById("State_UnitedStates").value
        }   
    else if(window.document.getElementById("Country").value=="Canada")
 	    {
            window.document.getElementById("State").value = window.document.getElementById("State_Canada").value	    
        }
    else
        {    
            window.document.getElementById("State").value = window.document.getElementById("txtStore_State").value	    
        }
    }      

</script>


<%
sFormAction = "Update_records_action.asp"
sTitle = "Add Customer"
sFullTitle = "<a href=my_customer_base.asp class=white>Customers</a> > Add"
sCommonName="Customer"
sCancel="my_customer_base.asp"
sSubmitName = "Register_User"
thisRedirect = "modify_customer.asp"
sMenu = "customers"
sQuestion_Path = "reports/customers.htm"
createHead thisRedirect
%>

<input type="hidden"  name="Add_New_Customer" value="Add">
<input type="Hidden" name="Record_type" value="0">


    

				

				
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Username</td>
						<td class="inputvalue">
							<INPUT	name="User_Id" size="40" >
							<INPUT type="hidden"  name=User_Id_C value="Re|String||||Username">
                                   <% small_help "Username" %></td>
					</TR>
				
					<tr bgcolor='#FFFFFF'> 
						<td class="inputname">Password</td>
						<td class="inputvalue">
							<INPUT name="Password" type=password size="40" >
							<INPUT name=Password_C type=hidden value="Re|String||||Password" size="40" >
                                   <% small_help "Password" %></td>
					</TR>
				
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Password Confirm</td>
						<td class="inputvalue">
							<INPUT name="Password_Confirm" type=password size="40" >
							<INPUT name=Password_Confirm_C type=hidden value="Re|String||||Password Confirm" size="40" >
                                   <% small_help "Password Confirm" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Budget</td>
						<td class="inputvalue">
							<INPUT name="Budget_Orig" type=text size="40" >
							<INPUT name=Budget_Orig_C type=hidden value="Re|Integer||||Budget" size="40" >
                                   <% small_help "Budget" %></td>
					</TR>
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Rewards</td>
						<td class="inputvalue">
							<INPUT name="Reward_Total" type=text size="40" >
							<INPUT name=Reward_Total_C type=hidden value="Re|Integer||||Rewards" size="40" >
                                   <% small_help "Reward_Total" %></td>
					</TR>
					<tr bgcolor='#FFFFFF'>
			<td class="inputname">Promo Emails </td>
			<td class="inputvalue">
						<input class="image" type="checkbox" name="Spam" value="-1" >
			<% small_help "Promo Emails" %></td>
		</tr>
			
				
			
				<tr bgcolor='#FFFFFF'>
			<td class="inputname">Tax Exempt </td>
			<td class="inputvalue">
						<input class="image" type="checkbox" name="Tax_Exempt" value="-1" >
			<% small_help "Tax Exempt" %></td>
		</tr>
		<% if session("service_type") >=5 then %>
			<tr bgcolor='#FFFFFF'>
			<td class="inputname">Protected Access</td>
			<td class="inputvalue">
						<input class="image" type="checkbox" <%= checked_protected %> name="protected_access" value="-1" >
			<% small_help "Protected Access" %></td>
		</tr>
		<% end if %>
	
					<tr bgcolor='#FFFFFF'>
						<TD colspan="3"><hr></TD>
					</TR>

					<tr bgcolor='#FFFFFF'>
						<TD colspan="3"><b>Billing Info</b></TD>
					</TR> 

					<tr bgcolor='#FFFFFF'>
						<td class="inputname">First Name</td>
						<td class="inputvalue">
							<INPUT	name=First_name size="40" >
							<INPUT type="hidden"  name=First_name_C value="Re|String||||First Name">
                                   <% small_help "First Name" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Last Name</td>
						<td class="inputvalue">
							<INPUT	name=last_name size="40" >
							<INPUT type="hidden"  name=Last_name_C value="Re|String||||Last Name">
                                   <% small_help "Last Name" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Company</td>
						<td class="inputvalue">
							<INPUT	name=company size="40" >
                                   <% small_help "Company" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Address Line 1</td>
						<td class="inputvalue">
							<INPUT	name=address1 size="40" >
                                   <% small_help "Address Line 1" %></td>
					</TR>
				
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Address Line 2</td>
						<td class="inputvalue">
							<INPUT	name=address2 size="40" >
                                   <% small_help "Address Line 2" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">City</td>
						<td class="inputvalue">
							<INPUT	name=city size="40" >
                                   <% small_help "City" %></td>
					</TR>

					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Country</td>
						<td class="inputvalue">
							<select size="1" name="Country" id="Country" onChange="validateState()">
							<option selected="selected" value="United States">United States</option>							
								<% sql_region = "SELECT Country FROM Sys_Countries ORDER BY Country;"
								   set myfields=server.createobject("scripting.dictionary")
									Call DataGetrows(conn_store,sql_region,mydata,myfields,noRecords)
								   if noRecords = 0 then
									FOR rowcounter= 0 TO myfields("rowcount")
										response.write "<Option value='"&mydata(myfields("country"),rowcounter)&"'>"&mydata(myfields("country"),rowcounter)&"</option>"
									Next
									End if
							response.write "</select>"
                     small_help "Country" %></td>
					</TR>

	            <tr bgcolor='#FFFFFF'>
						<td class="inputname">State</td>
						<td class="inputvalue">
							
								   	<%
									set rsState = server.CreateObject("Adodb.Recordset")
									sql_STMT_USA = "SELECT State, State_Long FROM Sys_States WHERE Country = 'United States' order by State_Long asc"
									sql_STMT_Canada = "SELECT State, State_Long FROM Sys_States WHERE Country = 'Canada' order by State_Long asc"
									%>
									
							<div id="UnitedStatesDiv" style="display:block">								   
								   	<select id='State_UnitedStates' name='State_UnitedStates' onChange="storeStateData()">								   		
										<option selected="selected"  value='AL'>Alabama</option>
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
										</select>								
							</div>
						
									<div id="CanadaDiv" style="display:none">
									<select id='State_Canada' name='State_Canada' onChange="storeStateData()">										
										<option selected="selected" value='AB'>Alberta</option>
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
										%>																
								  </select>
								  </div>

												
									<div id="optDiv" style="display:none">
									<input type="text" id='txtStore_State' name='txtStore_State' value="" size="40" maxlength=100 onBlur="storeStateData()">
							        </div>
							
							
							<input type="hidden" name="State" id="State" value="AL">
							<INPUT type="hidden" name=State_C value="Re|String||||State">
                     		<% small_help "State" %>
						</td>
					</tr>

					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Zipcode</td>
						<td class="inputvalue">
							<INPUT	name=zip size="10">
                                   <% small_help "Zipcode" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Phone</td>
						<td class="inputvalue">
							<INPUT	name=phone size="40" >
                                   <% small_help "Phone" %></td>
					</TR>

					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Email</td>
						<td class="inputvalue">
							<INPUT	name=Email size="40" >
                                   <% small_help "Email" %></td>
					</TR>
	
					<tr bgcolor='#FFFFFF'>
						<td class="inputname">Fax</td>
						<td class="inputvalue">
							<INPUT	name=Fax size="40"
                                   <% small_help "Fax" %></td>
					</TR>
	



<% createFoot thisRedirect, 1%>
