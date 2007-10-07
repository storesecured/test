<!--#include file="Global_Settings.asp"-->

<%
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><%
	
else

	'ADD TAX RATE
	if Request.Form("Form_Name") = "Add_Tax" then
		if Request.Form("TaxRate") = "" OR Request.Form("Zip_Start") = "" Or Request.Form("Zip_End") = "" Or Request.Form("Department_IDs") = ""	then
			response.redirect "admin_error.asp?message_id=21"
		elseif not isnumeric(Request.Form("TaxRate")) then
			response.redirect "admin_error.asp?message_id=22"
		end if
	
		'RETRIEVE FORM DATA
		Zip_Name = Request.Form("Name")
		Zip_Start = Request.Form("Zip_Start")
		Zip_End = Request.Form("Zip_End")
		TaxRate = cdbl(Request.Form("TaxRate"))
		Department_IDs = Replace(Request.Form("Department_IDs"), " ", "")
		Zip_Range_ID = Request.Form("Zip_Range_ID")
		Tax_Shipping = Request.Form("Tax_Shipping")
		if Tax_Shipping <> "" then
			Tax_Shipping = 1
		else
			Tax_Shipping = 0
		end if
	
		if Request.Form("op")<>"edit" then
			'IF ADD THEN PREPARE AND RUN AN INSERT
			sql_insert = "insert into Store_Zips (Store_id,Zip_Start,Zip_End,Zip_Name,TaxRate,Department_IDs,Tax_Shipping) values ("&Store_id&","&Zip_Start&","&Zip_End&",'"&Zip_Name&"',"&TaxRate&",'"&Department_IDs&"',"&Tax_Shipping&")"
			conn_store.Execute sql_insert
			Response.Redirect "Tax_List.asp?"&sAddString
		else
			'IF EDIT THEN PREPARE AND RUN AN UPDATE
			Zip_Range_ID=Request.Form("Zip_Range_ID")
			sqlUpdate="update Store_Zips set Zip_Start=" & Zip_Start & ", " & _
						"Zip_End=" & Zip_End & ", " & _
						"Zip_Name='" & Zip_Name & "', " & _
						"TaxRate=" & TaxRate & ", " & _ 
						"Tax_Shipping=" & Tax_Shipping & ", " &_
						"Department_IDs='" & Department_IDs & "' " & _	
						"where store_id=" & store_id & " and Zip_Range_ID=" & Zip_Range_ID						
			response.write sqlUpdate
			conn_store.Execute sqlUpdate
			Response.Redirect "Tax_List.asp?"&sAddString
		end if
	end if
	
	if Request.Form("Form_Name") = "Add_Tax_State" then
		if Request.Form("TaxRate") = "" OR Request.Form("State") = "" Or Request.Form("Department_IDs") = ""	then
			response.redirect "admin_error.asp?message_id=21"
		elseif not isnumeric(Request.Form("TaxRate")) then
			response.redirect "admin_error.asp?message_id=22"
		end if
	
		'RETRIEVE FORM DATA
		State = Request.Form("State")
		TaxRate = cdbl(Request.Form("TaxRate"))
		Department_IDs = Replace(Request.Form("Department_IDs"), " ", "")
		State_Tax_ID = Request.Form("State_Tax_ID")
		Tax_Shipping = Request.Form("Tax_Shipping")
		if Tax_Shipping <> "" then
			Tax_Shipping = 1
		else
			Tax_Shipping = 0
		end if
		if Request.Form("op")<>"edit" then
			'IF ADD THEN PREPARE AND RUN AN INSERT
			sql_insert = "insert into Store_State_Tax (Store_id,State,TaxRate,Department_IDs,Tax_Shipping) values ("&Store_id&",'"&State&"',"&TaxRate&",'"&Department_IDs&"',"&Tax_Shipping&")"
			 conn_store.Execute sql_insert

			Response.Redirect "Tax_List.asp?"&sAddString
		else
			'IF EDIT THEN PREPARE AND RUN AN UPDATE
			State_Tax_ID=Request.Form("State_Tax_ID")
			sqlUpdate="update Store_State_Tax set State='" & State & "', " & _
						"TaxRate=" & TaxRate & ", " & _
						"Department_IDs='" & Department_IDs & "', " & _
						"Tax_Shipping=" & Tax_Shipping & " " &_
						"where store_id=" & store_id & " and State_Tax_ID=" & State_Tax_ID
			conn_store.Execute sqlUpdate
			'response.write sqlUpdate
			Response.Redirect "Tax_List.asp?"&sAddString
		end if
	end if
	
	if Request.Form("Form_Name") = "Add_Tax_Country" then

		if Request.Form("TaxRate") = "" OR Request.Form("Country") = "" Or Request.Form("Department_IDs") = ""	then
			response.redirect "admin_error.asp?message_id=21"
		elseif not isnumeric(Request.Form("TaxRate")) then
			response.redirect "admin_error.asp?message_id=22"
		end if
	
		'RETRIEVE FORM DATA
		Country = Request.Form("Country")
		TaxRate = cdbl(Request.Form("TaxRate"))
		Department_IDs = Replace(Request.Form("Department_IDs"), " ", "")
		Country_Tax_ID = Request.Form("Country_Tax_ID")
		Tax_Shipping = Request.Form("Tax_Shipping")
		if Tax_Shipping <> "" then
			Tax_Shipping = 1
		else
			Tax_Shipping = 0
		end if
		if Request.Form("op")<>"edit" then
			'IF ADD THEN PREPARE AND RUN AN INSERT
			sql_insert = "insert into Store_Country_Tax (Store_id,Country,TaxRate,Department_IDs,Tax_Shipping) values ("&Store_id&",'"&Country&"',"&TaxRate&",'"&Department_IDs&"',"&Tax_Shipping&")"
			 
			 conn_store.Execute sql_insert

			Response.Redirect "Tax_List.asp?"&sAddString
		else
			'IF EDIT THEN PREPARE AND RUN AN UPDATE
			Country_Tax_ID=Request.Form("Country_Tax_ID")
			sqlUpdate="update Store_Country_Tax set Country='" & Country & "', " & _
						"TaxRate=" & TaxRate & ", " & _
						"Department_IDs='" & Department_IDs & "', " & _
						"Tax_Shipping=" & Tax_Shipping & " " &_
						"where store_id=" & store_id & " and Country_Tax_ID=" & Country_Tax_ID
			conn_store.Execute sqlUpdate
			Response.Redirect "Tax_List.asp?"&sAddString
		end if
	end if
	
	'DELETE TAX RATE
	if Request.QueryString("Delete_Zip_Range_ID") <> "" then
		Zip_Range_ID = Request.QueryString("Delete_Zip_Range_ID")
		sql_delete = "delete from Store_Zips where Zip_Range_ID = "&Zip_Range_ID
		conn_store.Execute sql_delete
		Response.Redirect "Tax_List.asp?"&sAddString
	end if
	  
	  if Request.QueryString("Delete_State_Tax_ID") <> "" then
		Delete_State_Tax_ID = Request.QueryString("Delete_State_Tax_ID")
		sql_delete = "delete from Store_State_Tax where State_Tax_ID = "&Delete_State_Tax_ID
		conn_store.Execute sql_delete
		Response.Redirect "Tax_List.asp?"&sAddString
	end if
	
	if Request.QueryString("Delete_Country_Tax_ID") <> "" then
		Delete_Country_Tax_ID = Request.QueryString("Delete_Country_Tax_ID")
		sql_delete = "delete from Store_Country_Tax where Country_Tax_ID = "&Delete_Country_Tax_ID
		conn_store.Execute sql_delete
		Response.Redirect "Tax_List.asp?"&sAddString
	end if
end if
 %>
