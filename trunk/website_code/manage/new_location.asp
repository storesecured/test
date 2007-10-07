<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/List_of_countries.asp"-->

<% 
op=Request.QueryString("op")
if op="edit" then
	Ship_Location_Id = Request.QueryString("Id")
	if Ship_Location_Id = "" then
	Ship_Location_Id = Request.QueryString("Ship_Location_Id")
	end if
	if not isNumeric(Ship_Location_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	if Ship_Location_Id=0 then
	   response.redirect "company.asp"
        end if
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sql_select="select * from store_ship_location where Ship_Location_Id=" & Ship_Location_Id & " and store_id=" & store_id
	rs_Store.open sql_select,conn_store,1,1
	if rs_Store.bof or rs_Store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	location_name=rs_store("location_name")
	location_state=rs_store("location_state")
	location_zip=rs_store("location_zip")
	location_country=rs_store("location_country")
  rs_Store.close
  
end if

sCommonName = "Shipping Location"
if op="edit" then
        sTitle = "Edit Shipping Location - "&Location_name
        sFullTitle = "Shipping > Locations > Edit - "&Location_name
else
        sTitle = "Add Shipping Location"
        sFullTitle = "Shipping > Locations > Add"

end if
sTextHelp="shipping/shipping-locations.doc"

sCancel = "location_Manager.asp"
sFormAction = "location_action.asp"
thisRedirect = "new_location.asp"
sFormName ="Create_Location"
sMenu = "shipping"
sQuestion_Path = "advanced/location_manager.htm"
createHead thisRedirect
%>


<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Ship_Location_Id" value="<%=Ship_Location_Id%>">


			
		
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Name</b></td>
				<td width="60%" class="inputvalue">
					<input type="text" name="Location_name" size="30" value="<%= Location_name %>" maxlength=50>
						<INPUT type="hidden"  name=Location_name_C value="Re|String|0|50|||Name">
                              <% small_help "Name" %></td>
			</tr>
    
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Zip</B></td>
				<td width="60%" class="inputvalue">
					<input type="text" name="Location_Zip" size="30" value="<%= Location_Zip %>" maxlength=50>
						<INPUT type="hidden"  name=Location_Zip_C value="Re|String|0|50|||Zip">
                              <% small_help "Zip" %></td>
			</tr>
			
            
                  <tr bgcolor='#FFFFFF'>
	<td class='inputname'><B>Country</b></td>
	<td class='inputvalue'>
							<select id='location_country' name='location_country' size="1" onChange="validateState()">
										<%
											if  location_country <> "" then
										%>
										<option selected="selected" value='<%=location_country%>'><%=location_country%></option>
										<%
											end if
										%>
										<option><%=create_country_list ("location_country",location_country,1,"") %></option>												
							 </select>
		
		 <% small_help "Country" %></td>
</tr>




								<%
							   
									set rsState = server.CreateObject("Adodb.Recordset")
									
								   	sql_STMT_USA = "SELECT State, State_Long FROM Sys_States WHERE Country = 'United States' order by State_Long asc"
									sql_STMT_Canada = "SELECT State, State_Long FROM Sys_States WHERE Country = 'Canada' order by State_Long asc"
								

								   if location_country = "United States" then
								    Div1="block"									
								    Div2="none"
								    Div3="none"	
									

								rsState.open sql_STMT_USA,conn_store,1,1									
									if not(rsState.eof or rsState.bof) then
										do while not rsState.eof
											if rsState.fields("State") = location_state then
												USA_State_Value = rsState.fields("State")
												USA_State_Text = rsState.fields("State_Long")
												exit do												
											end if											
										rsState.MoveNext
										loop										
									end if
									rsState.close
																		
									Canada_State_Value = ""
									Other_State=""
									
								   elseif location_country = "Canada" then
								    Div1="none"
								    Div2="block"
								    Div3="none"
									USA_State_Value=""
									
									
								rsState.open sql_STMT_Canada,conn_store,1,1									
									if not(rsState.eof or rsState.bof) then
										do while not rsState.eof
											if rsState.fields("State") = location_state then
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
									USA_State_Value=""
									Canada_State_Value = ""
									Other_State = location_state								   
								   end if
								   %>



            
            <tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>State/Province</B></td>
				<td width="60%" class="inputvalue">
								  
                        <div id="UnitedStatesDiv" style="display:<%= Div1 %>">								   
								   	<select id='State_UnitedStates' name='State_UnitedStates' onChange="storeStateData()">								   		
										<%
											if USA_State_Value <> "" then
										%>
										<option selected="selected" value='<%=USA_State_Value%>'><%=USA_State_Text%></option>
										<%
										end if
										%>
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
						
									<div id="CanadaDiv" style="display:<%= Div2 %>">
									<select id='State_Canada' name='State_Canada' onChange="storeStateData()">										
										<%
											if Canada_State_Value <> "" then
										%>
											<option selected="selected" value='<%=Canada_State_Value%>'><%=Canada_State_Text%></option>
										<%
											end if
										%>
										<%
										rsState.open sql_STMT_Canada,conn_store,1,1
											if not (rsState.eof or rsSTate.bof) then
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

												
									<div id="optDiv" style="display:<%= Div3 %>">
										<input type="text" id='txtStore_State' name='txtStore_State' value="<%= Other_State %>" size="30" maxlength=100 onBlur="storeStateData()">
							        </div>
                        
                        
                        <input type="hidden" id="Location_State" name="Location_State" size="30" value="<%= Location_State %>" maxlength=2>
						<INPUT type="hidden"  name="Location_State_C" value="Op|String|0|2|||State">
                              <% small_help "State" %></td>
			</tr>
      
      

<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Location_name","req","Please enter a name for this location.");
 frmvalidator.addValidation("Location_Zip","req","Please enter a zip code.");
 
 
 
  function validateState()
 {
 	if(window.document.getElementById("location_country").value=="United States")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="block"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="none" 	       
 	        window.document.getElementById("State_UnitedStates").options[0].selected=true
 	        window.document.getElementById("Location_State").value = window.document.getElementById("State_UnitedStates").value
 	    }
 	else if(window.document.getElementById("location_country").value=="Canada")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="block"
 	        window.document.getElementById("optDiv").style.display="none" 
 	        window.document.getElementById("State_Canada").options[0].selected=true
 	        window.document.getElementById("Location_State").value = window.document.getElementById("State_Canada").value	    	    
 	    }
 	else
 	   {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="block"
 	        window.document.getElementById("txtStore_State").value=""	
			window.document.getElementById("Location_State").value=""			    
 	    }
 	    
 }	
 
 
  function storeStateData()
    {
    if(window.document.getElementById("location_country").value=="United States")
 	    {
            window.document.getElementById("Location_State").value = window.document.getElementById("State_UnitedStates").value
        }   
    else if(window.document.getElementById("location_country").value=="Canada")
 	    {
            window.document.getElementById("Location_State").value = window.document.getElementById("State_Canada").value	    
        }
    else
        {    
            window.document.getElementById("Location_State").value = window.document.getElementById("txtStore_State").value	    
        }
    }      



</script>

