<!--#include file="pagedesign.asp"-->
<!--#include file="include/country_list.asp"-->
<!--#include file="include/List_of_countries.asp"-->


<script language="javascript" type="text/javascript" src="js/display_cust_form.js"></script>

<%
' RETRIEVE VALUES FROM THE DATABASE, USING
' RECORD_TYPE PARAM: 
'	0=BILLING ADDRESS
'	1=FIRST SHIPPING ADDRESS
'	2=SECOND SHIPPING ADDRESS, ETC
sql_select_cust =  "SELECT Store_Customers.User_id, Store_Customers.Password, Store_Customers.Budget_left,Store_Customers.Last_name, Store_Customers.First_name, Store_Customers.Company, Store_Customers.Address1, Store_Customers.Address2, Store_Customers.City, Store_Customers.State, Store_Customers.Zip, Store_Customers.Country, Store_Customers.EMail, Store_Customers.Phone,Store_Customers.FAX,Store_Customers.Is_Residential FROM Store_Customers WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type =	"&Record_type&" and store_id="&Store_Id
rs_Store.open sql_select_cust,conn_store,1,1
if not rs_Store.eof then
	rs_Store.MoveFirst 
	User_id= rs_Store("User_id")
	Password= rs_Store("Password")
	Company= rs_Store("Company")
	First_name= rs_Store("First_name")
	Last_name= rs_Store("Last_name")
	Address1= rs_Store("Address1")
	Address2= rs_Store("Address2")
	City= rs_Store("City")
	State= rs_Store("State")
	Zip= rs_Store("Zip")
	Country= rs_Store("Country") 
	Phone= rs_Store("Phone")
	Fax= rs_Store("Fax")
	EMail= rs_Store("Email") 
	Budget_left = rs_Store("Budget_left")
	Is_Residential = rs_Store("Is_Residential")
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
	
<form method="POST" action="update_records_action.asp?cid=<%=cid%>"  onsubmit="return Validations()">
<input type="Hidden" name="Page_id" value="<%= Page_id %>">
<input type="Hidden" name="Record_Type" value="<%= Record_Type %>">

	<tr>
	<td width="73" class="inputname">First Name</td>
		<td width="195" class="inputvalue">
			<input id="First_name" name="First_name" value="<%= First_name%>" size="40" >*
			<INPUT type="hidden"  name=First_name_C value="Re|String||||"></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Last Name</td>
		<td width="195" class="inputvalue">
			<input id="Last_name" name="Last_name" value="<%= Last_name%>" size="40" >*
			<INPUT type="hidden"  name="Last_name_C" value="Re|String||||"></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Company</td>
		<td width="195" class="inputvalue">
			<input id="Company" name="Company" value="<%= Company%>" size="40" ></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Address Line 1</td>
		<td width="195" class="inputvalue">
			<input id="Address1" name="Address1" value="<%= Address1%>" size="40" >*
			<INPUT type="hidden"  name=Address1_C value="Re|String||||"></td>
	</tr>

	<tr>
	<td width="73" class="inputname">Address Line 2</td>
		<td width="195" class="inputvalue">
			<input id="Address2" name="Address2" value="<%= Address2%>" size="40" ></td>
	</tr>

	<tr>
	<td width="73" class="inputname">City</td>
		<td width="195" class="inputvalue">
			<input id="City" name="City" value="<%= City%>"size="40" >*
			<INPUT type="hidden"  name=City_C value="Re|String||||"></td>
	</tr>

	<tr>

	<td width="73" class="inputname">Country</td>
		<td width="195" class="inputvalue">
			 <select id='Country' name='Country' onChange="validateState()">
										<%
											if Coutnry <> "" then
										%>
											<option selected="selected" value='<%=Country %>'><%=Country %></option>
										<%
											end if
										%>
											<option><%=create_country_list ("Country",Country,1,"") %></option>												
							 </select>*
			</td>			
	</tr>

     <tr>
	<td width="73" class="inputname">State</td>
		<td width="195" class="inputvalue">
			
			<%
								   
								   	sql_STMT_USA = "SELECT State, State_Long FROM Sys_States WHERE Country = 'United States' order by State_Long asc"
									sql_STMT_Canada = "SELECT State, State_Long FROM Sys_States WHERE Country = 'Canada' order by State_Long asc"
									
								   if Country = "United States" then
								    Div1="block"									
								    Div2="none"
								    Div3="none"	
									
									rs_Store.open sql_STMT_USA,conn_store,1,1
									if not(rs_Store.eof or rs_Store.bof) then
										do while not rs_Store.eof
											if rs_Store.fields("State") = State then
												USA_State_Value = rs_Store.fields("State")
												USA_State_Text = rs_Store.fields("State_Long")
												exit do												
											end if											
										rs_Store.MoveNext
										loop										
									end if
									rs_Store.close
										
									'arr_USA = Array("AL:Alabama","AK:Alaska","AZ:Arizona","AR:Arkansas","AA:Armed Forces - Americas","AE:Armed Forces - Europe","AP:Armed Forces - Pacific","CA:California","CO:Colorado","CT:Connecticut","DE:Delaware","FL:Florida","GA:Georgia","HI:Hawaii","ID:Idaho","IL:Illinois","IN:Indiana","IA:Iowa","KS:Kansas","KY:Kentucky","LA:Louisiana","ME:Maine","MD:Maryland","MA:Massachusetts","MI:Michigan","MN:Minnesota","MS:Mississippi","MO:Missouri","MT:Montana","NE:Nebraska","NV:Nevada","NH:New Hampshire","NJ:New Jersey","NM:New Mexico","NY:New York","NC:North Carolina","ND:North Dakota","OH:Ohio","OK:Oklahoma","OR:Oregon","PA:Pennsylvania","PR:Puerto Rico","RI:Rhode Island","SC:South Carolina","SD:South Dakota","TN:Tennessee","TX:Texas","UT:Utah","VT:Vermont","VI:Virgin Islands","VA:Virginia","WA:Washington","DC:Washington DC","WV:West Virginia","WI:Wisconsin","WY:Wyoming") 									
				
									'for a=0 to ubound(arr_USA) step 1
									'	if Left(arr_USA(a),2) = State then
									'		USA_State_Text = Right(arr_USA(a),len(arr_USA(a))-3)
									'	exit for	
									'	end if
									'next
									Canada_State = ""
									Other_State=""
									
								   elseif Country = "Canada" then
								    Div1="none"
								    Div2="block"
								    Div3="none"
									USA_State=""
									'Canada_State_Value = State
									
									rs_Store.open sql_STMT_Canada,conn_store,1,1
									if not(rs_Store.eof or rs_Store.bof) then
										do while not rs_Store.eof
											if rs_Store.fields("State") = State then
												Canada_State_Value = rs_Store.fields("State")
												Canada_State_Text = rs_Store.fields("State_Long")
												exit do												
											end if											
										rs_Store.MoveNext
										loop										
									end if
									rs_Store.close									
									Other_State=""
								   
								   else
								    Div1="none"
								    Div2="none"
								    Div3="block"	
									USA_State=""
									Canada_State = ""
									Other_State = State								   
								   end if
								   %>
								   
								   
								   <div id="UnitedStatesDiv" style="display: <%= Div1 %>">								   
								   	<select id='State_UnitedStates' name='State_UnitedStates' onChange="storeStateData()">
								   		<%
										if USA_State_Value <> "" then
										%>
										<option selected="selected" value='<%= USA_State_Value %>'><%= USA_State_Text %></option>
										<%
										end if
										rs_Store.open sql_STMT_USA,conn_store,1,1
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
						
									<div id="CanadaDiv" style="display: <%= Div2 %>">
									<select id='State_Canada' name='State_Canada' onChange="storeStateData()">
										<%
										if Canada_State_Value <> "" then
										%>
										<option selected="selected" value='<%= Canada_State_Value %>'><%= Canada_State_Text %></option>
										<%
										end if
										rs_Store.open sql_STMT_Canada,conn_store,1,1
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

												
									<div id="optDiv" style="display: <%= Div3 %>">
									<input type="text" id='txtStore_State' name='txtStore_State' value="<%= Other_State %>" size="40" maxlength=100 onBlur="storeStateData()">
							        </div>
									
			<input type="hidden" id="State" name="State" value="<%=State%>" size="40" >
			</td>
	</tr>
	
	

	<tr>
	<td width="73" class="inputname">Zipcode</td>
		<td width="195" class="inputvalue">
			<input id="Zip" name="Zip" value="<%= Zip%>" size="15">*
			<INPUT type="hidden"  name=Zip_C value="Re|String||||"></td>
	</tr>			 

	<tr>
	<td width="73" class="inputname">Phone</td>
		<td width="195" class="inputvalue">
			<input id="Phone" name="Phone" value="<%= Phone%>" size="40" >*
			<INPUT type="hidden"  name=Phone_C value="Re|String||||"></td>
	</tr>

	<tr>
	<td width="73" class="inputname">Email</td>
		<td width="195" class="inputvalue">
			<input id="Email" name="Email" value="<%= Email%>" size="40" >*
			<INPUT type="hidden"  name="Email_C" value="Re|String||||"></td>
	</tr>
	 
	<tr>
	<td width="73" class="inputname">Fax</td>
		<td width="195" class="inputvalue">
			<input id="Fax" name="Fax" value="<%= Fax%>" size="40" ></td>
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
