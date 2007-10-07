<!--#include file="Global_Settings.asp"-->


<%

If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

 'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
%> 
	<!--#include virtual="common/Error_Template.asp"-->
<% 

else

	'CREATE / UPDATE A COUPON
	if Request.Form("Form_Name") = "Create_Location" then
		Location_Name = checkStringForQ(Request.Form("Location_Name"))
		Location_Zip = checkStringForQ(Request.Form("Location_Zip"))
		Location_State = checkStringForQ(Request.Form("Location_State"))
		Ship_Location_Id = checkStringForQ(Request.Form("Ship_Location_Id"))
        Location_Country = checkStringForQ(Request.Form("Location_Country"))
		
		if Location_Country = "" then
			response.Redirect("admin_error.asp?message_add=Please Select Country")
		else

			' INSERTING / UPDATING COUPON IN STORE_COUPONS TABLE
			if Request.Form("op")<>"edit" then
				' now we are ready to insert to table ...
				sql_insert = "insert into store_ship_location (Store_id,Location_name,location_zip,location_country,Location_State) values ("&Store_id&",'"&Location_name&"','"&Location_Zip&"','"&Location_Country&"','"&Location_State&"')"
				conn_store.Execute sql_insert
			else
				sqlUpdate="update store_ship_location set location_name='" & location_name & "', " & _
						"Location_Zip='" & Location_Zip & "', " & _
						"Location_State='" & Location_State & "', " & _
						"Location_Country='" & Location_Country & "' " & _
						"where store_id=" & store_id & " and Ship_Location_Id=" & Ship_Location_Id
				conn_store.Execute sqlUpdate
			end if
		end if
		Response.Redirect "location_manager.asp"
	end if 

	' DELETE A COUPON
	if Request.QueryString("Delete_Ship_Location_Id") <> "" then
		Ship_Location_Id = Request.QueryString("Delete_Ship_Location_Id")
		if not isNumeric(Ship_Location_Id) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if
		sql_delete = "Delete from store_ship_location where Ship_Location_Id = "&Ship_Location_Id
		conn_store.Execute sql_delete
		Response.Redirect "location_manager.asp"
	end if	

End If 

%>
