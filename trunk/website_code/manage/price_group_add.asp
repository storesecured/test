<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

Item_Id = Request.QueryString("Item_Id")

sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
rs_store.close

sCancel="price_group_list.asp?Item_Id="&Item_Id
sFullTitle ="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Price Group</a> > "
sCommonName="Item Price Group"
sInstructions="Select the Customer Group for which you would like to offer a special price for this item."
mode_operation = request.querystring("mode_operation")

		'ADD THE PRICE GROUP
		if Request.Form("Form_Name") = "Add_Price_Group" then
			'ERROR CHECKING
			If Form_Error_Handler(Request.Form) <> "" then 
				Error_Log = Form_Error_Handler(Request.Form)
				%><!--#include file="Include/Error_Template.asp"--><%
				response.end
			else
				'INSERT THE PRICE GROUP RECORD
				item_id = request.Form("item_id")
				Group_Price = Request.Form("Group_Price")
				Customer_Group = Request.Form("customer_group")
		
		
				sql_ins="insert into Store_items_Price_Group (Store_ID,Item_Id,Group_Price,customer_group ) values ("&Store_ID&","&Item_Id&","&Group_Price&",'" & customer_group &"')"
				response.Write sql_ins
				conn_store.Execute sql_ins
				response.redirect "Price_Group_List.asp?Item_Id="&Item_Id
			end if
		else
			response.Write "<h3>" & mode_operation & "</h3>"
		end if
		'END OF ADD ACTION CODE

		'***************************EDIT  THE PRICE GROUP
		if Request.Form("Form_Name") = "Edit_Price_Group" then
			'ERROR CHECKING
			If Form_Error_Handler(Request.Form) <> "" then 
				Error_Log = Form_Error_Handler(Request.Form)
				%><!--#include file="Include/Error_Template.asp"--><%
				response.end
			else
				'INSERT THE PRICE GROUP RECORD
				item_id = request.Form("item_id")
				Group_Price = Request.Form("Group_Price")
				Customer_Group = Request.Form("customer_group")
				price_group_id = request.form("price_group_id")
		
				sql_update="update Store_items_Price_Group set Group_Price = " & Group_Price & ", customer_group= " & Customer_Group & " where price_group_id = " & price_group_id & " and store_id =" & Store_ID
'				response.Write sql_update
				conn_store.Execute sql_update
				response.redirect "Price_Group_List.asp?Item_Id="&Item_Id
			end if
	
		end if
		'**********************************END OF EDIT ACTION CODE




		%>
		
		<%
if request.QueryString("OP")="edit" then
	'--------------EDIT MODE
	

		mode_operation = "E"
		sFormAction = "price_group_add.asp"
		sName = "Add_Class_1"
		sFormName = "Edit_Price_Group"
		sTitle = "Edit Price Group - "&Item_Name
		sFullTitle = sFullTitle&"Edit - "&Item_Name
		sSubmitName = "Edit_Price_Group"
		thisRedirect = "price_group_list.asp"
		sMenu="inventory"
		createHead thisRedirect
		
		if Service_Type < 7  then %>
			<tr bgcolor='#FFFFFF'>
			<td colspan=2>
				This feature is not available at your current level of service.  
                                Gold service or higher is needed. 
			</td></tr>
		   <% createFoot thisRedirect, 0%>
		<% else %>
		
		
		

			<%
			'--CREATE THE QUERY GET THE DATA
			'response.Write request.QueryString("ID")
			sql_search = "select * from Store_Items_Price_Group where price_group_id = " & request.QueryString("ID") & " and store_id =" & Store_ID
			set rs_edit_price_group = server.CreateObject("ADODB.Recordset")
			rs_edit_price_group.open sql_search, conn_store ,2,3
			if not rs_edit_price_group.eof then
'				response.Write rs_edit_price_group.fields(0) & " | " & rs_edit_price_group.fields(1) & " | " & rs_edit_price_group.fields(2) & " | " & rs_edit_price_group.fields(3) 
				group_Price = rs_edit_price_group.fields("group_price")
				var_customer_group = rs_edit_price_group.fields("customer_group")
				
			end if
			rs_edit_price_group.close
			set rs_edit_price_group=nothing

			'---
			%>
			<input name="price_group_id" type ="hidden" value="<%=request.querystring("id")%>">
			<input name="item_id" type ="hidden" value="<%=request.querystring("item_id")%>">

			
		
						<tr bgcolor='#FFFFFF'>
						<td class="inputname">Customer Group</td>
						<td class="inputvalue">

							<select name="customer_group">
								
		<!--BEGIN LISTING THE CUSTOMER GROUP-->
								<%
								
								sql_customer_group = "select group_name,group_id from Store_Customers_Groups where store_id =" & store_Id
								set rs_store1 = server.CreateObject("ADODB.Recordset")
								rs_store1.open sql_customer_group, conn_store, 2,3
								while not rs_store1.eof
									if rs_store1.fields("group_id") = cint(var_customer_group) then
								%>
									<option value="<%=rs_store1.fields("group_id")%>" selected><%=rs_store1.fields("group_name")%></option>
									
								<%									
									else
								%>
									<option value="<%=rs_store1.fields("group_id")%>"><%=rs_store1.fields("group_name")%></option>									
								<%
									end if
									rs_store1.movenext
								wend
								rs_store1.close
								set rs_store1 = nothing
								%>
		<!--END LISTING THE CUSTOMER GROUP-->						
							</select><% small_help "Customer Group" %></td>
						
						</tr>
		
						<tr bgcolor='#FFFFFF'>
							<td class="inputname">Group Price</td>
							<td class="inputvalue">
								<%= Store_Currency %><input type="text" name="Group_Price" size="10" value=<%=group_price%> onKeyPress="return goodchars(event,'0123456789.')">
								<input type="hidden" name="Group_Price_C" value="Re|Integer||||Group Price">
								<% small_help "Group Price" %></td>
						</tr>
			 
		
		
		<% createFoot thisRedirect, 1%>
		<SCRIPT language="JavaScript">
		 var frmvalidator  = new Validator(0);
			frmvalidator.addValidation("Group_Price","req","Please enter a item price.");
			frmvalidator.addValidation("customer_group","req","Please choose the Customer Group");
		</script>
		<% end if %>

<%
	'---------------END OF EDIT MODE		
else
	'=--------------------------ADD  MODE
%>
<%	
		sFormAction = "price_group_add.asp"
		sName = "Add_Class_1"
		sFormName = "Add_Price_Group"
		sTitle = "Add Price Group - "&Item_Name
		sFullTitle=sFullTitle&"Add - "&Item_Name
		sSubmitName = "Add_Price_Group"
		thisRedirect = "price_group_list.asp"
		sMenu="inventory"
		createHead thisRedirect
		
		if Service_Type < 7  then %>
			<tr bgcolor='#FFFFFF'>
			<td colspan=2>
				This feature is not available at your current level of service.  Gold service or higher is needed.
			</td></tr>
		   <% createFoot thisRedirect, 0%>
		<% else %>
		
		
		

			<input name="item_id" type ="hidden" value="<%=request.querystring("item_id")%>">

			
		
						<tr bgcolor='#FFFFFF'>
						<td class="inputname">Customer Group</td>
						<td class="inputvalue">
							<select name="customer_group">
								<option value="" >--Select Customer Group--</option>
		<!--BEGIN LISTING THE CUSTOMER GROUP-->
								<%
								sql_customer_group = "select group_name,group_id from Store_Customers_Groups where store_id =" & store_Id
								set rs_store1 = server.CreateObject("ADODB.Recordset")
								rs_store1.open sql_customer_group, conn_store, 2,3
								while not rs_store1.eof
								%>
									<option value="<%=rs_store1.fields(1)%>"><%=rs_store1.fields(0)%></option>
								<%
									rs_store1.movenext
								wend
								rs_store1.close
								set rs_store1 = nothing
								%>
		<!--END LISTING THE CUSTOMER GROUP-->						
							</select><% small_help "Customer Group" %></td>
						
						</tr>
		
						<tr bgcolor='#FFFFFF'>
							<td class="inputname">Group Price</td>
							<td class="inputvalue">
								<%= Store_Currency %><input type="text" name="Group_Price" size="10" onKeyPress="return goodchars(event,'0123456789.')">
								<input type="hidden" name="Group_Price_C" value="Re|Integer||||Group Price">
								<% small_help "Group Price" %></td>
						</tr>
			 
		
		
		<% createFoot thisRedirect, 1%>
		<SCRIPT language="JavaScript">
		 var frmvalidator  = new Validator(0);
			frmvalidator.addValidation("Group_Price","req","Please enter a item price.");
			frmvalidator.addValidation("customer_group","req","Please choose the Customer Group");
		</script>
		<% end if %>


<%
'-------------END OF MODE CHECKING
end if
%>


