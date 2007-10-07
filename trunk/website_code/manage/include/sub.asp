<!--#include virtual="common/common_functions.asp"-->
<%

function encodeHTML(sHTML)
	sHTML=replace(sHTML,"&","&amp;")
	sHTML=replace(sHTML,"<","&lt;")
	sHTML=replace(sHTML,">","&gt;")
	encodeHTML=sHTML
end function

' ================================================================
' FORMAT A NUMBER AS A Zip Code
Function Zipcode_Format_Function (Zip)
	if Zip < 10000 then
		if Zip < 1000 then
			if Zip < 100 then
				if Zip < 10 then
					Zip = "0000"&Zip
				else
					 Zip = "000"&Zip
				end if
			else
				 Zip = "00"&Zip
			end if
		else
			 Zip = "0"&Zip
		end if
	end if
	Zipcode_Format_Function = Zip
End Function	 


' ================================================================ 
' CHECK IF SPECIFIED COUSTOMER BELONGS TO SPECIFIED COUSTOMERS GROUP
Function  Is_Cid_In_Group(cid,Group_Id)
	Is_Cid_In_Group = false
	
	Customer_Buy_more_then_required = False
	Customer_have_Enough_Budget = False
	Country_Found = False
	Company_Found = False
	Cid_Found = False
	sIncluded=0
	

	'SELECT THE CID, COUNTRY, REGION AND OTHER ATTRIBUTES
	sql_cid_attr="SELECT Store_Customers.CCid, Store_Customers.Country, Store_Customers.Orders_Dept, Store_Customers.Orders_Total, Store_Customers.Company, Store_Customers.Budget_left FROM Store_Customers WITH (NOLOCK) WHERE Store_Customers.Cid="&Cid&" AND Store_Customers.Record_type=0"
	rs_store.open sql_cid_attr,conn_store,1,1
	rs_store.MoveFirst
		Country = rs_store("Country")
		Orders_Total = rs_store("Orders_Total")
		Company = rs_store("Company")
		Budget_left = rs_store("Budget_left")
		CCid = rs_store("CCid")
		Orders_Dept = rs_Store("Orders_Dept")
	rs_store.Close

	'CHECK IF CID BELONGS TO SPECIFIED GROUP
	sql_group_attr ="SELECT * FROM Store_Customers_Groups WITH (NOLOCK) WHERE Store_Customers_Groups.Group_id="&Group_id&" AND Store_Customers_Groups.Store_id="&Store_id
	rs_store.open sql_group_attr,conn_store,1,1
	If  rs_store.BOF = True Or rs_store.EOF = True THEN
		Is_Cid_In_Group = False
		Exit Function
	End If
	rs_store.MoveFirst
		Group_Country = rs_store("Group_Country")
		Group_Purchase_History = rs_store("Group_Purchase_History")
		Group_Company = rs_store("Group_Company")
		Group_budget_min = rs_store("Group_budget_min")
		Group_Cid = rs_store("Group_Cid")
		Group_Dept = rs_Store("Group_Dept")
	rs_store.Close

	'CHECK COUSTOMER BUDGET AND PURCHASE HISTORY
	if Orders_Total => Group_Purchase_History then
		Customer_Buy_more_then_required = true
	End if
	
	'CHECK BUDGET REQUIREMENT
	If Budget_left => Group_budget_min then 
		Customer_have_Enough_Budget = true
	End if

	if UCase(Group_Country) = "ALL" Then
		Country_Found = True
	else
		Group_Country_array = Split(Group_Country,",",-1,1)
		For Each one_country in Group_Country_array
			if UCase(Country) = UCase(one_country) then 
				Country_Found = True
			End if	
		Next
	End If


	If UCase(Group_Company) = "ALL" Then
		Company_Found = True
	Else	
		Group_Company_array = Split(Group_Company,",",-1,1)
		For Each one_Company in Group_Company_array
			if UCase(Company) = UCase(one_Company) then 
				Company_Found = True
			End if
		Next
	End if

	If Group_Cid <> "" Then
		Group_Cid_array = Split(Group_Cid,",",-1,1)
		For Each one_Cid in Group_Cid_array
			if int(CCid) = int(one_Cid) then
				Cid_Found = True
			End if
		Next
	End if
	
	if (Orders_Dept = "" or isnull(Orders_Dept)) and group_dept <> "" then
	  sIncluded = 0
	elseif group_dept = "" then
	  sIncluded = 1
	else
	  sIncluded=0

	  if not isnull(Orders_Dept) then
	  sArray = split(Orders_Dept, ",")
		for each one in sArray
		if not isnull(group_dept) then
		  sArrayCurrent = split(group_dept,",")
		  if sIncluded = 0 then
			  for each Dept in sArrayCurrent
					if Dept=one and sIncluded=0 then
						sIncluded = 1
					 end if
			  next
			end if
		end if
		next
		end if
	end if

	'CHECK IF ALL THE CONDITION ARE TRUE
	If (Customer_Buy_more_then_required = True And Customer_have_Enough_Budget = True And Country_Found = True And Company_Found = True and sIncluded=1) Or	Cid_Found = True Then
		Is_Cid_In_Group = True 
	else 
		Is_Cid_In_Group = false
	end if

End Function 
' ================================================================ 
' DISPLAYS BILLING OR SHIPPING INFORMATION FOR THE SPECIFIED
' COUSTOMER
 Sub Display_cust_info(cid,Record_type)
	sql_select_cust =  "SELECT Store_Customers.Last_name, Store_Customers.First_name, Store_Customers.Company, Store_Customers.Address1, Store_Customers.Address2, Store_Customers.City, Store_Customers.State, Store_Customers.Zip, Store_Customers.Country, Store_Customers.EMail, Store_Customers.Phone,Store_Customers.FAX FROM Store_Customers WITH (NOLOCK) WHERE Store_Customers.Cid="&cid&" AND Store_Customers.Record_type =	"&Record_type&" and Store_customers.store_id="&Store_Id
	rs_Store.open sql_select_cust,conn_store,1,1
	if not rs_store.eof and not rs_store.bof then
	rs_Store.MoveFirst %>	
		<blockquote> 
			
				<b><%= rs_Store("Company") %></b><br>
				<%= rs_Store("First_name")  %> &nbsp;<%= rs_Store("Last_name") %><br>
				<%= rs_Store("Address1") %><br>
				<%=  rs_Store("Address2") %><br>
				<%=  rs_Store("City") %>, <%= rs_Store("State") %>, <%= rs_Store("zip") %><br>
				<%= rs_Store("Country") %><br>
				Phone <%= rs_Store("Phone") %><br>
				Fax <%= rs_Store("Fax") %><br>
				<%= rs_Store("EMail") %>

		</blockquote><%
	end if
        rs_Store.Close
End Sub

' ================================================================ 
' DISPLAYS A DROP DOWN MENU CONTAINING AVAILABLE PAYMENT METHODS
Sub Disp_Payment
	sql_Payment_methods =  "SELECT Payment_method_id,Payment_name FROM Payment_methods  WHERE Store_id="&Store_id
	rs_Store.open sql_Payment_methods,conn_store,1,1
	rs_Store.MoveFirst %>	 
		<select name="Payment_Method" size="1">
			<option 
				<% if request.form("Payment_Method")="" then  %>
					selected 
				<% End If %>
				>All</option>
			<% Do While Not rs_Store.EOF %>
				<option	value="<%= rs_Store("Payment_name") %>"
					<% if request.form("Payment_Method")= rs_Store("Payment_name") then  %>
						selected 
					<% End If %>
					><%= rs_Store("Payment_name") %></option>
				<% rs_Store.MoveNext %>
			<% Loop %> 
			<% rs_Store.Close%>
		</select><%
End Sub


' ================================================================ 
' DISPLAYS INFORMATION FROM AN INVOICE
Sub Create_Invoice (Cart_Type,Soid)
	'LOAD VALUES FROM STORE_TRANSACTIONS TABLE
	'CHECK IF ORDER IS VERIFIED 
	Verified = Does_Order_Verified(Soid)
	sql_Show_big_cart = "Select * from Store_transactions WITH (NOLOCK) where Oid ="&Soid&"	AND Store_id ="&Store_id&" order by store_transactions.item_sku"

	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_Show_big_cart,mydata,myfields,noRecords)
	Total_weight=0
 %>
	<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
		<tr>	 
			<% if Cart_Type = "Big" then%>
			<td><STRONG></STRONG></td>
			<% End If %> 
		<td><STRONG>Qty Ordered</STRONG></td>
			<td><STRONG>SKU </STRONG></td>
		<td><STRONG>Name </STRONG></td>
		<% if Cart_Type = "Big" then %>

				<td><STRONG>Attribute SKU </STRONG></td>


				<td><STRONG>User Definable Fields</STRONG></td>
                 <% End If %>
                 <% if Cart_Type <>"Packing" then %>
		<td><STRONG>Price</STRONG></td>
		<td><STRONG>Total </STRONG> </td>
		<% if Recurring_Total <> 0 then %>
		<td><STRONG>Recurring</STRONG> </td>
		<% end if %>
			<% if Cart_Type = "Invoice" and Verified = true then%>
				<td>Electronic Delivery</a></td>
				<td>File Location</a></td>
			<% End If %>
		</tr>
                <% end if %>
		<% Line_id = 1
		Quantity_Total=0
			if noRecords = 0 then
				FOR rowcounter= 0 TO myfields("rowcount")
				Ext_Sale_Price = mydata(myfields("sale_price"),rowcounter)*mydata(myfields("quantity"),rowcounter)
				if Ext_Sale_Price = 0 then
					Ext_Sale_Price = "<i>Free</i>"
				else
					Ext_Sale_Price = Currency_Format_Function(mydata(myfields("sale_price"),rowcounter)*mydata(myfields("quantity"),rowcounter))
				End if
				Sale_Price = mydata(myfields("sale_price"),rowcounter)
				if Sale_Price = 0 then
					Sale_Price = "<i>Free</i>"
				else
					Sale_Price = Currency_Format_Function(mydata(myfields("sale_price"),rowcounter))
				End if
				Item_Id = mydata(myfields("item_id"),rowcounter)
				Quantity = mydata(myfields("quantity"),rowcounter)
				Item_name = mydata(myfields("item_name"),rowcounter)
				Item_Sku = mydata(myfields("item_sku"),rowcounter)
				Item_Attribute_skus = mydata(myfields("item_attribute_skus"),rowcounter)
				Item_Attribute_hiddens = mydata(myfields("item_attribute_hiddens"),rowcounter)
				Promotion_ids = mydata(myfields("promotion_ids"),rowcounter)
				Notes = mydata(myfields("notes"),rowcounter)
				U_d_1 = mydata(myfields("u_d_1"),rowcounter)
				U_d_2 = mydata(myfields("u_d_2"),rowcounter)
				U_d_3 = mydata(myfields("u_d_3"),rowcounter)
				U_d_4 = mydata(myfields("u_d_4"),rowcounter)
				U_d_5 = mydata(myfields("u_d_5"),rowcounter)
				Recurring_Fee = Currency_Format_Function(mydata(myfields("recurring_fee"),rowcounter))
                                item_weight = mydata(myfields("item_weight"),rowcounter)
                                Total_Weight=Total_Weight+(item_weight*quantity)

				sql_item = "select U_d_1_name, U_d_2_name, U_d_3_name, U_d_4_name, U_d_5_name, Description_S,File_Location, Item_Remarks from store_items WITH (NOLOCK) where Store_id = "&Store_id&" and item_id = "&item_id

				rs_store.open sql_item,conn_store,1,1
				if rs_store.Bof = false then
					Description_S = Rs_store("Description_S")
					File_Location = Rs_store("File_Location")
					Item_Remarks = rs_store("Item_Remarks")
					U_d_1_name = rs_store("U_d_1_name")
					U_d_2_name = rs_store("U_d_2_name")
					U_d_3_name = rs_store("U_d_3_name")
					U_d_4_name = rs_store("U_d_4_name")
					U_d_5_name = rs_store("U_d_5_name")
				else
					Description_S = ""
					File_Location = ""
					Item_Remarks = ""
					U_d_1_name = ""
					U_d_2_name = ""
					U_d_3_name = ""
					U_d_4_name = ""
					U_d_5_name = ""
				end if
			rs_store.close %>
			<tr>	
				<% if Cart_Type = "Big" then %>
				<td></td>	
				<% End If %>		 
			<% if Cart_Type = "Big" then %>
					<td align=center><%= Quantity %></td>
				<% Else	%> 
					<td><%= Quantity %></td>
				<% End If %>
				<td><%= Item_sku %></td> 
			<td>
				<% If Item_id > 0 then%>
					<a class=link href="http://<%= request.servervariables("HTTP_HOST") %>/item_edit.asp?item_id=<%= Item_id %>"><%= Item_name %></a>
				<% Else	%>
					<%= Item_name %>
				<% End If %>
				</td> 
				<% if Cart_Type = "Big" then %>
					<td><%= Item_Attribute_skus %>&nbsp;</td>
					
				<% end if %>
				<% if Cart_Type="Big" or Cart_Type="Packing" then %>
					<td>
						<table>

								<% If U_d_1 <> "" then %>
									<tr><td colspan="6"><%= U_d_1_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_1 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_2 <> "" then %>
									<tr><td colspan="6"><%= U_d_2_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_2 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_3 <> "" then %>
									<tr><td colspan="6"><%= U_d_3_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_3 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_4 <> "" then %>
									<tr><td colspan="6"><%= U_d_4_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_4 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_5 <> "" then %>
									<tr><td colspan="6"><%= U_d_5_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_5 %></td></tr></table></td></tr>
								<% End If %>

						</table>
					</td>
				<% End If %>
            <% if Cart_Type <>"Packing" then %>
			<td><%= Sale_Price %></td> 
			<td><%= Ext_Sale_Price %></td>
			<% if Recurring_Total <> 0 then %>
			<td><%= Recurring_Fee %></td>
			<% end if %>
				<% if Cart_Type = "Big" then %>
					<input type="hidden" name="Item_Id_<%= Line_id %>" value="<%= Item_Id %>">
				<% End If %>	 
			</td>

				<% if Promotion_ids <> "" and (Cart_Type = "Big" or Cart_Type = "Invoice" ) then %>
					<td><small><i>Promotion</i></small></td>
				<% End If %>
         <% if Notes <> "" then %>
					<td><small><%= Notes %></small></td>
				<% End If %>
			<% if Cart_Type = "Invoice" and Verified = true then %>
				<td></td>
				<td><a class=link href="esd_download.asp?<%= url_string %>&File_Location=<%= File_Location %>"><%= File_Location %></a></td>
			<% End If %>
                        <% end if %> 
		</tr>  
		<% Line_id = Line_id + 1 %>
		<% Quantity_Total = Quantity_Total+Quantity %>
	<% Next %>

	<% else

		sql_Show_big_cart = "select shopper_id from store_purchases WITH (NOLOCK) where store_id="&Store_id&" and oid="&Soid
		set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_Show_big_cart,mydata,myfields,noRecords)
		rowcounter=0
		 %>
		<tr><td colspan=3>Order appears to be an abandoned cart</td><td colspan=2><a class=link target=_blank href="https://<%= Secure_Name %>/include/process_manual.asp?Shopper_Id=<%=mydata(myfields("shopper_id"),rowcounter)%>&Store_Id=<%= Store_Id %>&Return_From=Admin&Oid=<%= Oid %>">Process Order</a></td><td colspan=2><a class=link target=_blank href="<%= Site_Name %>show_big_cart.asp?Shopper_Id=<%=mydata(myfields("shopper_id"),rowcounter)%>&Store_Id=<%= Store_Id %>&Page_Id=15">View Cart</a></td></tr>
	<% End If %>
	</table>
	<% if Cart_Type ="Packing" then %>
	   <table><TR><TD>Total # of items: <%= Quantity_Total %></td></tr></table>
	<% end if %>
	<br><%
       Session("Total_Weight")=formatnumber(Total_Weight,2)
End Sub

' FOR PICK LIST added on 06/02/2005

'---------------------
' ================================================================ 
' DISPLAYS PICKLIST
Sub Create_PickList(Cart_Type,Soid)

	'LOAD VALUES FROM STORE_TRANSACTIONS TABLE
	'CHECK IF ORDER IS VERIFIED 
	Soid = replace(Soid,"*",",")

	sql_Show_big_cart = " Select item_id,sum(quantity) as quantity,item_sku,item_name,item_attribute_skus,u_d_1,u_d_2,u_d_3,u_d_4,u_d_5 "&_
						" from Store_transactions WITH (NOLOCK) where Oid in ("&Soid&") AND Store_id = "&Store_id &_
						" group by item_sku,item_name,item_attribute_skus,u_d_1,u_d_2,u_d_3,u_d_4,u_d_5,item_id "&_
						" order by item_sku asc,item_attribute_skus asc,u_d_1 asc,u_d_2 asc,u_d_3 asc,u_d_4 asc,u_d_5 asc "
	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_Show_big_cart,mydata,myfields,noRecords)
 %>
	<table cellspacing="0" cellpadding="0" border="1"	width="100%" >
		<tr>	 
			<% if Cart_Type = "Big" then%>
			<td><STRONG></STRONG></td>
			<% End If %> 
		<td><STRONG>Qty Ordered</STRONG></td>
			<td><STRONG>SKU </STRONG></td>
		<td><STRONG>Name </STRONG></td>
		<% if Cart_Type = "Big" then %>

				<td><STRONG>Attribute SKU </STRONG></td>


				<td><STRONG>User Definable Fields</STRONG></td>
			<% End If %>		 
		
		</tr> 
		<% Line_id = 1
			if noRecords = 0 then
				FOR rowcounter= 0 TO myfields("rowcount")
				
				Item_Id = mydata(myfields("item_id"),rowcounter)
				Quantity = mydata(myfields("quantity"),rowcounter)
				Item_name = mydata(myfields("item_name"),rowcounter)
				Item_Sku = mydata(myfields("item_sku"),rowcounter)
				Item_Attribute_skus = mydata(myfields("item_attribute_skus"),rowcounter)
				
				U_d_1 = mydata(myfields("u_d_1"),rowcounter)
				U_d_2 = mydata(myfields("u_d_2"),rowcounter)
				U_d_3 = mydata(myfields("u_d_3"),rowcounter)
				U_d_4 = mydata(myfields("u_d_4"),rowcounter)
				U_d_5 = mydata(myfields("u_d_5"),rowcounter)

				sql_item = "select U_d_1_name, U_d_2_name, U_d_3_name, U_d_4_name, U_d_5_name, Description_S,File_Location, Item_Remarks from store_items WITH (NOLOCK) where Store_id = "&Store_id&" and item_id = "&item_id
				
				rs_store.open sql_item,conn_store,1,1
				if rs_store.Bof = false then
					Description_S = Rs_store("Description_S")
					File_Location = Rs_store("File_Location")
					Item_Remarks = rs_store("Item_Remarks")
					U_d_1_name = rs_store("U_d_1_name")
					U_d_2_name = rs_store("U_d_2_name")
					U_d_3_name = rs_store("U_d_3_name")
					U_d_4_name = rs_store("U_d_4_name")
					U_d_5_name = rs_store("U_d_5_name")
				else
					Description_S = ""
					File_Location = ""
					Item_Remarks = ""
					U_d_1_name = ""
					U_d_2_name = ""
					U_d_3_name = ""
					U_d_4_name = ""
					U_d_5_name = ""
				end if
			rs_store.close %>
			<tr>	
				<% if Cart_Type = "Big" then %>
				<td></td>	
				<% End If %>		 
			<% if Cart_Type = "Big" then %>
					<td align=center><%= Quantity %></td>
				<% Else	%> 
					<td><%= Quantity %></td>
				<% End If %>
				<td><%= Item_sku %></td> 
			<td>
				<% If Item_id > 0 then%>
					<a class=link href="item_edit.asp?item_id=<%= Item_id %>"><%= Item_name %></a>
				<% Else	%>
					<%= Item_name %>
				<% End If %>
				</td> 
				<% if Cart_Type = "Big" then %>
					<td><%= Item_Attribute_skus %>&nbsp;</td>
					
				 
					<td>
						<table>

								<% If U_d_1 <> "" then %>
									<tr><td colspan="6"><%= U_d_1_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_1 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_2 <> "" then %>
									<tr><td colspan="6"><%= U_d_2_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_2 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_3 <> "" then %>
									<tr><td colspan="6"><%= U_d_3_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_3 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_4 <> "" then %>
									<tr><td colspan="6"><%= U_d_4_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_4 %></td></tr></table></td></tr>
								<% End If %>
								<% If U_d_5 <> "" then %>
									<tr><td colspan="6"><%= U_d_5_name %><br>
										<table border=1 cellspacing=0 cellpadding=0 width=100%><TR><TD><%= U_d_5 %></td></tr></table></td></tr>
								<% End If %>

						</table>
					</td>
				<% End If %>		 
			
		</tr>  
		<% Line_id = Line_id + 1 %>
		<% Quantity_Total = Quantity_Total+Quantity %>
		
	<% Next %>
	<% else

		sql_Show_big_cart = "select shopper_id from store_purchases WITH (NOLOCK) where store_id="&Store_id&" and oid="&Soid
		set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_Show_big_cart,mydata,myfields,noRecords)
		rowcounter=0
		 %>
		<tr><td colspan=3>Order appears to be an abandoned cart</td><td colspan=2><a class=link target=_blank href="<%= Site_Name %>check_out_action.asp?Shopper_Id=<%=mydata(myfields("shopper_id"),rowcounter)%>&Store_Id=<%= Store_Id %>&Return_From=Admin">Process Order</a></td><td colspan=2><a class=link target=_blank href="<%= Site_Name %>show_big_cart.asp?Shopper_Id=<%=mydata(myfields("shopper_id"),rowcounter)%>&Store_Id=<%= Store_Id %>&Page_Id=15">View Cart</a></td></tr>
	<% End If %>
	
	
	

	</table>
	<br>
	<table border='0' width=100%>
	<TR>
	<TD><b>Total Qty Ordered : <%=Quantity_Total%></b></td>
	<TD></td>
	</tr>
	</table>
	

	<br>
	<br>
	<br>
	<br>

	<table border='1' width=100% cellspacing="0" cellpadding="0">
	<TR>
		<TD width=15%>&nbsp;<b>Order#</b></td>
		<TD width=85%>&nbsp;<b>Comments/Special Notes</b></td>
	</tr>
	
		<%
		
		sql_show_order_comments = "select Cust_Notes,OID from store_purchases WITH (NOLOCK) where store_id="&Store_id&" and oid in ("&Soid&") and cust_notes<>''"
		
		set myfields=server.createobject("scripting.dictionary")
		Call DataGetrows(conn_store,sql_show_order_comments,mydata,myfields,noRecords)
		
		if noRecords = 0 then
				
				FOR rowcounter= 0 TO myfields("rowcount")
				
				CustNotes = mydata(myfields("cust_notes"),rowcounter)
				Order_Id = mydata(myfields("oid"),rowcounter)
				
		%>	
		<tr>
		<td>&nbsp;<%=Order_Id%></td>
		<td>&nbsp;<%=CustNotes%></td>
		</tr>
		<% next %>
		<% end if%>
	</table>
	
	<br>
	
	<%

End Sub
'---------------------


' FOR PICK LIST



' ================================================================
' RETURNS NUMBER OF RECORDS IN A RECORDSET
function get_number_of_recordset (rs)
	i = 0
	if not rs.bof then		
		while not rs.eof
			i = i + 1		
			rs.MoveNext 
		wend
		rs.MoveFirst 
	end if	
	get_number_of_recordset = i
end function


' ================================================================ 
' END AM EMAIL USING

Sub Send_Mailsupport(SentFrom,SendTo,Subject,Body,Attachment )
		SendTo_array = split(SendTo,",")
		for each one_SendTo in SendTo_array
			Set Mail = Server.CreateObject("Persits.MailSender")
			Mail.From = SentFrom ' Specify sender's address
			Mail.AddAddress one_SendTo ' Name is optional
			Mail.Subject = Subject
			if isnull(Attachment)=false then
				Mail.AddAttachment Attachment
			end if
			Mail.Body = Body
			Mail.Helo="mail.storesecured.com"
		   Mail.Queue=True
			Mail.Send
		next 

end sub


' ================================================================ 
' CHECK IF SPECIFIED ORDER IS VERIFIED
Function Does_Order_Verified(Oid) 
	sql_purchase = "select oid from store_purchases WITH (NOLOCK) where oid = "&oid&" and verified = -1 and Store_id = "&Store_id
	rs_Store.open sql_purchase,conn_store,1,1
		if rs_Store.eof = false or rs_Store.bof = false then
			Does_Order_Verified = true
		else
			Does_Order_Verified = false
		end if
	rs_Store.Close 	
End Function
' ================================================================ 
' FORMAT A DATE TO US FORMAT
Function GetUSFormatDate(OtherDate)
	 Dim	sTime

	 sTime = ""

	 If (Month(OtherDate) < 10) Then
		sTime = sTime & 0 & Month(OtherDate)
	 Else
		sTime = sTime & Month(OtherDate)
	 End If
	 sTime = sTime & "/"
	 
	 If (Day(OtherDate) < 10) Then
		sTime = sTime & 0 & Day(OtherDate)
	 Else
		sTime = sTime & Day(OtherDate)
	 End If
	 sTime = sTime & "/" & Year(OtherDate) & " "

	 If (Hour(OtherDate) < 10) Then
		sTime = sTime & 0 & Hour(OtherDate)
	 Else
		sTime = sTime & Hour(OtherDate)
	 End If
	 sTime = sTime & ":"

	 If (Minute(OtherDate) < 10) Then
		sTime = sTime & 0 & Minute(OtherDate)
	 Else
		sTime = sTime & Minute(OtherDate)
	 End If
	 sTime = sTime & ":"

	If (Second(OtherDate) < 10) Then
		sTime = sTime & 0 & Second(OtherDate)
	Else
		sTime = sTime & Second(OtherDate)
	End If

	GetUSFormatDate = sTime
End Function

' ================================================================
function nullifyQ(theString)
	tmpResult = theString
	tmpResult = replace(tmpResult,"'","''")
	tmpResult = replace(tmpResult,"%7E","~")
	nullifyQ = tmpResult
end function
' ================================================================

Function CheckReferer()
strHost = Request.ServerVariables("HTTP_HOST")
strReferer = Request.ServerVariables("HTTP_REFERER")
if Len(strRefererer) > 0 then
	iLength = Len(strReferer) - (InStr(1, strReferer, "://") + 2)
  strReferer = Right(strReferer, iLength)
  strReferer = Left(strReferer, InStr(1, strReferer, "/") - 1)
  If strReferer = strHost Then
	  blnCheckReferer = True
  Else
		blnCheckReferer = False
  End If
  CheckReferer = blnCheckReferer
else
  CheckReferer = True
end if
End Function

Function SizeUsage()

Set fs = CreateObject("Scripting.FileSystemObject")

size_EDD = 0
Folder_Path1 = fn_get_sites_folder(Store_Id,"Download")
if fs.FolderExists(Folder_Path1) then
	set f = fs.getFolder(Folder_Path1)
	size_EDD = f.size
	Set f = nothing
end if

size_Exp = 0
Folder_Path1 = fn_get_sites_folder(Store_Id,"Export")
if fs.FolderExists(Folder_Path1) then
	set f = fs.getFolder(Folder_Path1)
	size_Exp = f.size
	Set f = nothing
end if

size_Upl = 0
Folder_Path1 = fn_get_sites_folder(Store_Id,"Upload")
if fs.FolderExists(Folder_Path1) then
	set f = fs.getFolder(Folder_Path1)
	size_Upl = f.size
	Set f = nothing
end if

size_Img = 0
Folder_Path1 = fn_get_sites_folder(Store_Id,"Images")
if fs.FolderExists(Folder_Path1) then
	set f = fs.getFolder(Folder_Path1)
	size_Img = f.size
	Set f = nothing
end if

size_Froogle = 0

size_Logs = 0
Folder_Path1 = fn_get_sites_folder(Store_Id,"Log")
if fs.FolderExists(Folder_Path1) then
	set f = fs.getFolder(Folder_Path1)
	size_Logs = f.size
	Set f = nothing
end if

size_Email = 0
if Trial_Version = 0 and Store_Domain <> "" then
	sDomain = Replace(Store_Domain,"www.","")
	Folder_Path1 = Mail_Path&sDomain
	if fs.FolderExists(Folder_Path1) then
		set f = fs.getFolder(Folder_Path1)
		size_Email = f.size
		Set f = nothing
	end if
end if

set fs = nothing
total_size = (size_EDD + size_Exp + size_Upl + size_Img+size_Froogle+size_Logs+size_Email) / 1024

on error resume next
sql = "Get_DB_Size "&Store_id
rs_store.open sql,conn_store,1,1
item_size=rs_store("items")
coupon_size=rs_store("coupons")
dept_size=rs_store("departments")
customer_size=rs_store("customers")
shipping_size=rs_store("shippers")
shopper_size=rs_store("shoppers")
ShoppingCart_size=rs_store("carts")
Transaction_size=rs_store("transactions")
Order_size=rs_store("purchases")
Attribute_size=rs_store("attributes")
rs_store.close

total_size_db = 0
total_size_db = ((Item_size+Coupon_size+Dept_size+Customer_size+Shipping_size+Shopper_size+ShoppingCart_size+Transaction_Size+Order_Size+Attribute_Size)*100/1024)
total_size = total_size + total_size_db

if Service_Type = 0 then
	available = 5000
elseif Service_Type = 3 then
	available = 50000
elseif Service_Type = 5 then
	available = 100000
elseif Service_Type = 7 then
	available = 250000
elseif Service_Type = 9 or Service_Type = 10 then
	available = 500000
else
	available = 10000000
end if


if Additional_Storage > 0 then
	available = available + (100000 * Additional_Storage)
end if

SizeUsage = FormatNumber(total_size/available*100, 1, 0, 0, 0)


End Function




SUB DataGetRowsPaged(parmConn,parmSQL,byref parmArray,byref parmDict,noRecords,StartRow,RowsPerPage)
	dim conntemp,rstemp,howmany,counter
	set conntemp=server.createobject("adodb.connection")
	conntemp.open parmConn
	session("sql") = parmSQL

	set rstemp=conntemp.execute(parmSQL)
	If rstemp.eof then
		noRecords=1
		parmdict.add "rowcount", -1
		rstemp.close
		set rstemp=nothing
		conntemp.close
		set conntemp=nothing
		exit sub
	end if

	' Now lets fill up dictionary with field names/numbers
	howmany=rstemp.fields.count
	for counter=0 to howmany-1
		if not parmdict.Exists(lcase(rstemp(counter).name)) then
			parmdict.add lcase(rstemp(counter).name), counter
		end if
	next

	noRecords=0

   on error resume next
   ' Now lets grab all the records

   rstemp.move StartRow
	parmArray=rstemp.getrows(RowsPerPage+1)
   
   rstemp.close
   set rstemp=nothing
	conntemp.close
	set conntemp=nothing

	if err.number=3021 then
	      response.redirect "error.asp?Message_Id=101&Message_Add="&server.urlencode("The records requested do not exist")
	end if

	if cint(ubound(parmarray,2))=cint(RowsPerPage) then
            parmdict.add "rowcount", -2
	else
	    parmdict.add "rowcount", ubound(parmarray,2)+StartRow
        end if
	parmdict.add "colcount", howmany-1

   on error goto 0

END SUB

function checkCookies()
       checkCookies = 1
end function

function checkHtmlEditorCompat()
    checkHtmlEditorCompat=1
end function

function checkSSL()
    checkSSL=1
    
end function

Function FormatDate( _
   byVal strDate, _
   byVal strFormat _
   )

   ' Accepts strDate as a valid date/time,
   ' strFormat as the output template.
   ' The function finds each item in the 
   ' template and replaces it with the 
   ' relevant information extracted from strDate.
   ' You are free to use this code provided the following line remains
   ' www.adopenstatic.com/resources/code/formatdate.asp

   ' Template items
   ' %m Month as a decimal no. 2
   ' %M Month as a padded decimal no. 02
   ' %B Full month name February
   ' %b Abbreviated month name Feb   
   ' %d Day of the month eg 23
   ' %D Padded day of the month eg 09
   ' %O Ordinal of day of month (eg st or rd or nd)
   ' %j Day of the year 54
   ' %Y Year with century 1998
   ' %y Year without century 98
   ' %w Weekday as integer (0 is Sunday)
   ' %a Abbreviated day name Fri
   ' %A Weekday Name Friday
   ' %H Hour in 24 hour format 24
   ' %h Hour in 12 hour format 12
   ' %N Minute as an integer 01
   ' %n Minute as optional if minute <> 00
   ' %S Second as an integer 55
   ' %P AM/PM Indicator PM

   On Error Resume Next

   Dim intPosItem
   Dim int12HourPart
   Dim str24HourPart
   Dim strMinutePart
   Dim strSecondPart
   Dim strAMPM

   ' Insert Month Numbers
   strFormat = Replace(strFormat, "%m", DatePart("m", strDate), 1, -1, vbBinaryCompare)

   ' Insert Padded Month Numbers
   strFormat = Replace(strFormat, "%M", Right("0" & DatePart("m", strDate), 2), 1, -1, vbBinaryCompare)

   ' Insert non-Abbreviated Month Names
   strFormat = Replace(strFormat, "%B", MonthName(DatePart("m", strDate), False), 1, -1, vbBinaryCompare)

   ' Insert Abbreviated Month Names
   strFormat = Replace(strFormat, "%b", MonthName(DatePart("m", strDate), True), 1, -1, vbBinaryCompare)

   ' Insert Day Of Month
   strFormat = Replace(strFormat, "%d", DatePart("d",strDate), 1, -1, vbBinaryCompare)

   ' Insert Padded Day Of Month
   strFormat = Replace(strFormat, "%D", Right ("0" & DatePart("d",strDate), 2), 1, -1, vbBinaryCompare)

   ' Insert Day of Month Ordinal (eg st, th, or rd)
   strFormat = Replace(strFormat, "%O", GetDayOrdinal(Day(strDate)), 1, -1, vbBinaryCompare)

   ' Insert Day of Year
   strFormat = Replace(strFormat, "%j", DatePart("y",strDate), 1, -1, vbBinaryCompare)

   ' Insert Long Year (4 digit)
   strFormat = Replace(strFormat, "%Y", DatePart("yyyy",strDate), 1, -1, vbBinaryCompare)

   ' Insert Short Year (2 digit)
   strFormat = Replace(strFormat, "%y", Right(DatePart("yyyy",strDate),2), 1, -1, vbBinaryCompare)

   ' Insert Weekday as Integer (eg 0 = Sunday)
   strFormat = Replace(strFormat, "%w", DatePart("w",strDate,1), 1, -1, vbBinaryCompare)

   ' Insert Abbreviated Weekday Name (eg Sun)
   strFormat = Replace(strFormat, "%a", WeekDayName(DatePart("w",strDate,1), True), 1, -1, vbBinaryCompare)

   ' Insert non-Abbreviated Weekday Name
   strFormat = Replace(strFormat, "%A", WeekDayName(DatePart("w",strDate,1), False), 1, -1, vbBinaryCompare)

   ' Insert Hour in 24hr format
   str24HourPart = DatePart("h",strDate)
   If Len(str24HourPart) < 2 then str24HourPart = "0" & str24HourPart
   strFormat = Replace(strFormat, "%H", str24HourPart, 1, -1, vbBinaryCompare)

   ' Insert Hour in 12hr format
   int12HourPart = DatePart("h",strDate) Mod 12
   If int12HourPart = 0 then int12HourPart = 12
   strFormat = Replace(strFormat, "%h", int12HourPart, 1, -1, vbBinaryCompare)

   ' Insert Minutes
   strMinutePart = DatePart("n",strDate)
   If Len(strMinutePart) < 2 then strMinutePart = "0" & strMinutePart
   strFormat = Replace(strFormat, "%N", strMinutePart, 1, -1, vbBinaryCompare)

   ' Insert Optional Minutes
   If CInt(strMinutePart) = 0 then
      strFormat = Replace(strFormat, "%n", "", 1, -1, vbBinaryCompare)
   Else
      If CInt(strMinutePart) < 10 then strMinutePart = "0" & strMinutePart
      strMinutePart = ":" & strMinutePart
      strFormat = Replace(strFormat, "%n", strMinutePart, 1, -1, vbBinaryCompare)
   End If

   ' Insert Seconds
   strSecondPart = DatePart("s",strDate)
   If Len(strSecondPart) < 2 then strSecondPart = "0" & strSecondPart
   strFormat = Replace(strFormat, "%S", strSecondPart, 1, -1, vbBinaryCompare)

   ' Insert AM/PM indicator
   If DatePart("h",strDate) >= 12 then
      strAMPM = "PM"
   Else
      strAMPM = "AM"
   End If

   strFormat = Replace(strFormat, "%P", strAMPM, 1, -1, vbBinaryCompare)

   FormatDate = strFormat

End Function



Function GetDayOrdinal( _
   byVal intDay _
   )

   ' Accepts a day of the month 
   ' as an integer and returns the 
   ' appropriate suffix
   On Error Resume Next

   Dim strOrd

   Select Case intDay
   Case 1, 21, 31
      strOrd = "st"
   Case 2, 22
      strOrd = "nd"
   Case 3, 23
      strOrd = "rd"
   Case Else
      strOrd = "th"
   End Select

   GetDayOrdinal = strOrd

End Function

function image_replace (sStart)
    sStart=replace(sStart,"images/images_themes/","%THEMEPATH%")
     sStart=replace(sStart,Site_Name&"images/images_"&Store_Id&"/","%IMAGEPATH%")
     sStart=replace(sStart,Site_Name&"images/","%IMAGEPATH%")
     sStart=replace(sStart,Site_Name2&"images/images_"&Store_Id&"/","%IMAGEPATH%")
     sStart=replace(sStart,Site_Name2&"images/","%IMAGEPATH%")
     sStart=replace(sStart,"images/images_"&Store_Id&"/","%IMAGEPATH%")
     sStart=replace(sStart,"=images/","=%IMAGEPATH%")
     sStart=replace(sStart,"='images/","='%IMAGEPATH%")
     sStart=replace(sStart,"=""images/","=""%IMAGEPATH%")
     
     sStart=replace(sStart,"%IMAGEPATH%","images/")
     sStart=replace(sStart,"%THEMEPATH%","images/images_themes/")

     image_replace=sStart
end function

sub sub_template_select(sTemplate_Id,sPreview_Flag)
    Session("Template_Id")=sTemplate_Id
	if sPreview_Flag = 1  then
        Session("Preview")=1
        sql_move = "exec wsp_template_preview " & Store_id & ", " & sTemplate_Id&";"
		conn_store.Execute(sql_move)
        server.execute "reset_design.asp"
        Session("Template_Id")=""
        Session("Preview")=""
		response.redirect Site_Name&"preview.asp?Preview=1"
	else
	    Session("Preview")=0
		sql_move = "exec wsp_design_apply " & Store_id & ", " & sTemplate_Id&";"
		conn_store.Execute(sql_move)
		server.execute "reset_design.asp"
		Session("Template_Id")=""
        Session("Preview")=""
	end if
end sub

function fn_get_dept_ids (sItem_Id)
	sDept_ids=""
	sql_select = "select department_id from store_items_dept where store_id="&store_id&" and item_id="&sItem_Id
	set getdeptfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select,getdeptdata,getdeptfields,noRecords)
	if noRecords = 0 then
	    FOR getdeptrowcounter= 0 TO getdeptfields("rowcount")
		    if sDept_ids="" then
		        sDept_ids=getdeptdata(getdeptfields("department_id"),getdeptrowcounter)
            else
                sDept_ids=sDept_ids&","&getdeptdata(getdeptfields("department_id"),getdeptrowcounter)
            end if
	    next
	end if
	set getdeptfields = Nothing
	if sDept_ids="" then
		sDept_ids="0"
	end if
	fn_get_dept_ids=sDept_ids
end function
%>
