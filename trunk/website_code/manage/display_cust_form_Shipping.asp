<!--#include file="pagedesign.asp"-->
<!--#include file="include/country_list.asp"-->
<!--#include file="include/List_of_countries.asp"-->

<script language="javascript" type="text/javascript" src="js/display_cust_form_Shipping.js"></script>

<%
' RETRIEVE VALUES FROM THE DATABASE, USING
' RECORD_TYPE PARAM: 
'	0=BILLING ADDRESS
'	1=FIRST SHIPPING ADDRESS
'	2=SECOND SHIPPING ADDRESS, ETC
Record_type = 1
sql_select_cust =  "SELECT Store_Customers.User_id, Store_Customers.Password, Store_Customers.Budget_left,Store_Customers.Last_name, Store_Customers.First_name, Store_Customers.Company, Store_Customers.Address1, Store_Customers.Address2, Store_Customers.City, Store_Customers.State, Store_Customers.Zip, Store_Customers.Country, Store_Customers.EMail, Store_Customers.Phone,Store_Customers.FAX,Store_Customers.Is_Residential FROM Store_Customers WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type =	"&Record_type&" and store_id="&Store_Id
rs_Store.open sql_select_cust,conn_store,1,1
if not rs_store.eof then
	rs_Store.MoveFirst 
	User_id= Rs_store("User_id")
	Password= Rs_store("Password")
	Company= Rs_store("Company")
	First_name= Rs_store("First_name")
	Last_name= Rs_store("Last_name")
	Address1= Rs_store("Address1")
	Address2= Rs_store("Address2")
	City= Rs_store("City")
	State= Rs_store("State")
	Zip= Rs_store("Zip")
	Country= Rs_store("Country") 
	Phone= Rs_store("Phone")
	Fax= Rs_store("Fax")
	EMail= Rs_store("Email") 
	Budget_left = Rs_store("Budget_left")
	Is_Residential = Rs_store("Is_Residential")
else
	Company= ""
	First_name= ""
	Last_name= ""
	Address1= ""
	Address2= ""
	City= ""
	State= ""
	Zip= ""
	Country= ""
	Phone= ""
	Fax= ""
	EMail= ""
   Is_Residential = -1
end if
rs_Store.Close

if Is_Residential=0 then
   checked_residential=""
else
   checked_residential="checked"
end if

%>
<html>
<body>
<form method="POST" action="update_records_action.asp?cid=<%=cid%>"  onsubmit="return Validations1()">
<input type="Hidden" name="Page_id" value="<%= Page_id %>">
<input type="Hidden" name="Record_Type" value="<%= Record_Type %>">

	<tr>
	<td width="73" class="inputname">First Name</td>
		<td width="195" class="inputvalue">
			<input id="First_name1" name="First_name1" value="<%= First_name%>" size="40" >*
			<INPUT type="hidden"  name=First_name_C value="Re|String||||"></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Last Name</td>
		<td width="195" class="inputvalue">
			<input id="Last_Name1" name="Last_name1" value="<%= Last_name%>" size="40" >*
			<INPUT type="hidden"  name=Last_name_C value="Re|String||||"></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Company</td>
		<td width="195" class="inputvalue">
			<input name="Company" value="<%= Company%>" size="40" ></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Address Line 1</td>
		<td width="195" class="inputvalue">
			<input id="Address11" name="Address11" value="<%= Address1%>" size="40" >*
			<INPUT type="hidden"  name=Address1_C value="Re|String||||"></td>
	</tr>

	<tr>
	<td width="73" class="inputname">Address Line 2</td>
		<td width="195" class="inputvalue">
			<input	name="Address2" value="<%= Address2%>" size="40" ></td>
	</tr>

	<tr>
	<td width="73" class="inputname">City</td>
		<td width="195" class="inputvalue">
			<input id="City1" name="City1" value="<%= City%>"size="40" >*
			<INPUT type="hidden"  name=City_C value="Re|String||||"></td>
	</tr>

	<tr>	
	<td width="73" class="inputname">Country</td>
		<td width="195" class="inputvalue">
			 <select id='Country1' name='Country1' onChange="validateState1()">
			 							<%
										if Country <> "" then
										%>
										<option selected="selected" value='<%=Country %>'><%=Country %></option>
										<%
										end if
										%>
										<option><%=create_country_list ("Country1",Country1,1,"") %></option>												
							 </select>*	
			</td>			
	</tr>

     <tr>
	<td width="73" class="inputname">State</td>
		<td width="195" class="inputvalue">
			
			<%
											   
								   	sql_STMT_USA1 = "SELECT State, State_Long FROM Sys_States WHERE Country = 'United States' order by State_Long asc"
									sql_STMT_Canada1 = "SELECT State, State_Long FROM Sys_States WHERE Country = 'Canada' order by State_Long asc"

								   if Country = "United States" then
								    Div11="block"									
								    Div21="none"
								    Div31="none"
									rs_Store.open sql_STMT_USA1,conn_store,1,1
									if not(rs_Store.eof or rs_Store.bof) then
										do while not rs_Store.eof
											if rs_Store.fields("State") = State then
												USA_State_Value1 = rs_Store.fields("State")
												USA_State_Text1 = rs_Store.fields("State_Long")
												exit do												
											end if											
										rs_Store.MoveNext
										loop										
									end if
									rs_Store.close
									Canada_State1 = ""
									Other_State1 = ""
									
								   elseif Country = "Canada" then
								    Div11="none"
								    Div21="block"
								    Div31="none"
									USA_State1=""
									rs_Store.open sql_STMT_Canada1,conn_store,1,1
									if not(rs_Store.eof or rs_Store.bof) then
										do while not rs_Store.eof
											if rs_Store.fields("State") = State then
												Canada_State_Value1 = rs_Store.fields("State")
												Canada_State_Text1 = rs_Store.fields("State_Long")
												exit do												
											end if											
										rs_Store.MoveNext
										loop										
									end if
									rs_Store.close									
									Other_State=""
								   
								   else
								    Div11="none"
								    Div21="none"
								    Div31="block"	
									USA_State1=""
									Canada_State1 = ""
									Other_State1=State								   
								   end if
								   %>
								   
								   
								   <div id="UnitedStatesDiv1" style="display: <%= Div11 %>">								   
								   	<select id='State_UnitedStates1' name='State_UnitedStates1' onChange="storeStateData1()">
								   		<%
										if USA_State_Value1 <> "" then
										%>
										<option selected="selected" value='<%= USA_State_Value1 %>'><%= USA_State_Text1 %></option>
										<%
										end if
										rs_Store.open sql_STMT_USA1,conn_store,1,1
											if not (rs_Store.eof or rs_Store.bof) then
												do while not rs_Store.eof
										%>
												<option value="<%=rs_Store.fields("State")%>"><%=rs_Store.fields("State_Long")%></option>
										<%
												rs_Store.MoveNext
												loop
											end if
										rs_Store.close		
										%>		

									</select>*							
							</div>
						
									<div id="CanadaDiv1" style="display: <%= Div21 %>">
									<select id='State_Canada1' name='State_Canada1' onChange="storeStateData1()">
										<%
										if Canada_State_Value1 <> "" then
										%>
										<option selected="selected" value='<%= Canada_State_Value1 %>'><%= Canada_State_Text1 %></option>
										<%
										end if
										rs_Store.open sql_STMT_Canada1,conn_store,1,1
											if not (rs_Store.eof or rs_Store.bof) then
												do while not rs_Store.eof
										%>
												<option value="<%=rs_Store.fields("State")%>"><%=rs_Store.fields("State_Long")%></option>
										<%
												rs_Store.MoveNext
												loop
											end if
										rs_Store.close		
										%>																
								  </select>*
								  </div>

												
									<div id="optDiv1" style="display: <%= Div31 %>">
									<input type="text" id='txtStore_State1' name='txtStore_State1' value="<%= Other_State1 %>" size="40" maxlength=100 onBlur="storeStateData1()">
							        </div>
									
			<input type="hidden" id="State1" name="State1" value="<%=State%>" size="40" >
			</td>
	</tr>
	
	

	<tr>
	<td width="73" class="inputname">Zipcode</td>
		<td width="195" class="inputvalue">
			<input id="Zip1" name="Zip1" value="<%= Zip%>" size="15">*
			<INPUT type="hidden"  name=Zip_C value="Re|String||||"></td>
	</tr>			 

	<tr>
	<td width="73" class="inputname">Phone</td>
		<td width="195" class="inputvalue">
			<input id="Phone1" name="Phone1" value="<%= Phone%>" size="40" >*
			<INPUT type="hidden"  name=Phone_C value="Re|String||||"></td>
	</tr>

	<tr>
	<td width="73" class="inputname">Email</td>
		<td width="195" class="inputvalue">
			<input id="Email1" name="Email1" value="<%= Email%>" size="40" >*
			<INPUT type="hidden"  name=Email_C value="Re|String||||"></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Fax</td>
		<td width="195" class="inputvalue">
			<input id="text41" name="Fax" value="<%= Fax%>" size="40" ></td>
	</tr>
	<% if Record_Type<>0 then %>
   <tr>
	<td width="73" class="inputname">Residential</td>
		<td width="195" class="inputvalue">
			<input name="Is_Residential" value="-1" type=checkbox <%=checked_residential %>></td>
	</tr>
	<% else %>
   <tr>
	<td width="73" class="inputname">&nbsp;</td>
	</tr>
	<% end if %>
	<tr>
		<td colspan="2" align=center>
			<input type="reset" value="Reset" name="Reset"> 
			<% if Record_type=0 or Record_type=1 then %>
			<input type="submit" value="Update" name="Modify_my_<%= Record_Type %>">
			<% Else %>
				<input type="submit" value="Update" name="Modify_my_1">
			<% End If %>
	
			<% If allow_delete then %>
				<input type="hidden" name="Delete_Addr" value="<%= Record_type %>">
				<input type="submit" value="Delete" name="Delete_Ship_Addr">
			<% End If %>
			</td>
	 </tr> 

</form> 
</body>	
</html>