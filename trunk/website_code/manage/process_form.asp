<!--#include file="Global_Settings.asp"-->
<!--#include file="include/sub.asp"-->

<%
If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include virtual="common/Error_Template.asp"--><%
	
else
	 returnTo = request.form("redirect")
	'RETRIEVE FORM DATA
	if Request.Form("Form_Name") = "bck_font" then

		Template_Id=Request.Form("Template_Id")
	
		Global_Bck_Color = trim(Request.Form("Global_Bck_Color"))
		if Request.Form("Global_Text_Font_Color") <> "" then
			Global_Text_Font_Color = checkStringForQ(Request.Form("Global_Text_Font_Color"))
		else
			Global_Text_Font_Color = "000000"
		end if
	
		if Request.Form("Global_Text_Font_Face") <> "" then
			Global_Text_Font_Face = Request.Form("Global_Text_Font_Face")
		else
			Global_Text_Font_Face = "Arial"
		end if
	
		if Request.Form("Global_Text_Font_Size") <> "" then
			Global_Text_Font_Size = Request.Form("Global_Text_Font_Size")
		else
			Global_Text_Font_Size = "2"
		end if

		if Request.Form("Global_Link_Font_Color") <> "" then
			Global_Link_Font_Color = checkStringForQ(Request.Form("Global_Link_Font_Color"))
		else
			Global_Link_Font_Color = "000000"
		end if
	
		if Request.Form("Global_Link_Font_Face") <> "" then
			Global_Link_Font_Face = checkStringForQ(Request.Form("Global_Link_Font_Face"))
		else
			Global_Link_Font_Face = "Arial"
		end if
	
		if Request.Form("Global_Link_Font_Size") <> "" then
			Global_Link_Font_Size = Request.Form("Global_Link_Font_Size")
		else
			Global_Link_Font_Size = "2"
		end if
	
		if Request.Form("Global_Nav_Font_Color") <> "" then
			Global_Nav_Font_Color = checkStringForQ(Request.Form("Global_Nav_Font_Color"))
		else
			Global_Nav_Font_Color = "000000"
		end if

		if Request.Form("Global_Nav_Font_Face") <> "" then
			Global_Nav_Font_Face = checkStringForQ(Request.Form("Global_Nav_Font_Face"))
		else
			Global_Nav_Font_Face = "Arial"
		end if

		if Request.Form("Global_Nav_Font_Size") <> "" then
			Global_Nav_Font_Size = Request.Form("Global_Nav_Font_Size")
		else
			Global_Nav_Font_Size = "2"
		end if

		if not isNumeric(Global_Text_Font_Size) or not isNumeric(Global_Link_Font_Size) or not isNumeric(Global_Nav_Font_Size) then
			Response.Redirect "admin_error.asp?message_id=1"
		end if

		'UPDATE STORE_DESIGN TABLE
		
		sql_Update_Store_Activation = "update store_design_template set  Global_Bck_Color = '"&Global_Bck_Color&"', Global_Text_Font_Color = '"&Global_Text_Font_Color&"',  Global_Text_Font_Face = '"&Global_Text_Font_Face&"', Global_Text_Font_Size = "&Global_Text_Font_Size&", Global_Link_Font_Color = '"&Global_Link_Font_Color&"', Global_Link_Font_Size = "&Global_Link_Font_Size&", Global_Link_Font_Face = '"&Global_Link_Font_Face&"',Global_Nav_Font_Color = '"&Global_Nav_Font_Color&"', Global_Nav_Font_Size = "&Global_Nav_Font_Size&", Global_Nav_Font_Face = '"&Global_Nav_Font_Face&"' where Store_id = "&Store_id&" and Template_Id="&Template_Id
		
		conn_store.Execute sql_Update_Store_Activation
		response.redirect "bck_font.asp?Id="&Template_Id
	

	elseif Request.Form("Form_Name") = "buttons" then

		Template_Id=Request.Form("Template_Id")

		Button_image_Login = checkStringForQ(Request.Form("Button_image_Login"))
		Button_image_Continue = checkStringForQ(Request.Form("Button_image_Continue"))
		Button_image_Reset = checkStringForQ(Request.Form("Button_image_Reset"))
		Button_image_Yes = checkStringForQ(Request.Form("Button_image_Yes"))
		Button_image_No = checkStringForQ(Request.Form("Button_image_No"))
		Button_image_Order = checkStringForQ(Request.Form("Button_image_Order"))
		Button_image_Search = checkStringForQ(Request.Form("Button_image_Search"))
		Button_image_UpdateCart = checkStringForQ(Request.Form("Button_image_UpdateCart"))
		Button_image_ContinueShopping = checkStringForQ(Request.Form("Button_image_ContinueShopping"))
		Button_image_Checkout = checkStringForQ(Request.Form("Button_image_Checkout"))
		Button_image_Cancel = checkStringForQ(Request.Form("Button_image_Cancel"))
		Button_image_SaveCart = checkStringForQ(Request.Form("Button_image_SaveCart"))
		Button_image_RetrieveCart = checkStringForQ(Request.Form("Button_image_RetrieveCart"))
		Button_image_View = checkStringForQ(Request.Form("Button_image_View"))
		Button_image_Update = checkStringForQ(Request.Form("Button_image_Update"))
		Button_image_CancelOrder = checkStringForQ(Request.Form("Button_image_CancelOrder"))
		Button_image_ProcessOrder = checkStringForQ(Request.Form("Button_image_ProcessOrder"))
		Button_image_Next = checkStringForQ(Request.Form("Button_image_Next"))
		Button_image_Prev = checkStringForQ(Request.Form("Button_image_Prev"))
	    Button_image_Delete = checkStringForQ(Request.Form("Button_image_Delete"))
        Button_image_Up = checkStringForQ(Request.Form("Button_image_Up"))

		sql_Update_Store_Activation = "Update store_design_template set Button_image_Delete='"&Button_image_Delete&"', Button_image_Up='"&Button_image_Up&"', Button_image_Prev='"&Button_image_Prev&"', Button_image_Next='"&Button_image_Next&"', Button_image_Login = '"&Button_image_Login&"', Button_image_Continue = '"&Button_image_Continue&"', Button_image_Reset = '"&Button_image_Reset&"', Button_image_Yes = '"&Button_image_Yes&"', Button_image_No = '"&Button_image_No&"', Button_image_Order = '"&Button_image_Order&"', Button_image_Search = '"&Button_image_Search&"', Button_image_UpdateCart = '"&Button_image_UpdateCart&"', Button_image_ContinueShopping = '"&Button_image_ContinueShopping&"', Button_image_Checkout = '"&Button_image_Checkout&"', Button_image_Cancel = '"&Button_image_Cancel&"', Button_image_SaveCart = '"&Button_image_SaveCart&"', Button_image_RetrieveCart = '"&Button_image_RetrieveCart&"', Button_image_View = '"&Button_image_View&"', Button_image_Update = '"&Button_image_Update&"', Button_image_CancelOrder = '"&Button_image_CancelOrder&"', Button_image_ProcessOrder = '"&Button_image_ProcessOrder&"'	where Store_id = "&Store_id&" and Template_Id="&Template_Id

		conn_store.Execute sql_Update_Store_Activation
	   
	    response.redirect "buttons.asp?Id="&Template_Id
		'Response.redirect returnTo
	elseif Request.Form("Form_Name") = "Store_Nav_Layout" then

		Template_Id=Request.Form("Template_Id")

		Nav_Button_HTML = replace(Request.Form("Nav_Button_HTML"),"'","''")
			Nav_Link_HTML = replace(Request.Form("Nav_Link_HTML"),"'","''")
			sql_update = "UPDATE Store_Design_template SET Nav_Button_HTML='"&Nav_Button_HTML&"', Nav_Link_HTML='"&Nav_Link_HTML&"' WHERE Store_id="&Store_id&" and Template_Id="&Template_Id
			
			conn_store.Execute sql_update

		
	    response.redirect "nav_layout.asp?Id="&Template_Id
	end if
	
	if request.form("redirect") <> "" then
      response.redirect request.form("redirect")
   else
      response.redirect "admin_home.asp"
   end if
end if
%>
