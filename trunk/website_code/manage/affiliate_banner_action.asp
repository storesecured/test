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
	
	'RETRIEVE FORM DATA
		
		if(Request.Form("Banner_Image")="" and Request.Form("Banner_Text")="" )  then
			Response.Redirect "admin_error.asp?message_id=2"
		end if

		Banner_Image = checkstringforQ(trim(Request.Form("Banner_Image")))

		if Request.Form("Banner_Text") <> "" then
			Banner_Text = checkstringforQ(Request.Form("Banner_Text"))
		end if
	
		View_Order = Request.Form("View_Order")

		
		 if request.form("op") = "edit" then
			Banner_Id = request.form("Banner_Id")
		 end if
		
		'UPDATE STORE_AFFILIATE_BANNERS TABLE
   	    if request.form("op") <> "edit" then
			sql_Update_Affiliate_Banner = "Insert into Store_Affiliate_Banners(Store_Id, Banner_Text, Banner_Image, View_Order) values ("&Store_id&",'"&Banner_Text&"','"&Banner_Image&"',"&View_Order&")"
			conn_store.Execute sql_Update_Affiliate_Banner
		else
			 if View_Order = "" then
				 sql_Update_Affiliate_Banner = "update Store_Affiliate_Banners set Banner_Text='"&Banner_Text&"',Banner_Image='"&Banner_Image&"' where Store_Id="&Store_Id & " and Banner_Id= " & Banner_Id
			 else
				 sql_Update_Affiliate_Banner = "update Store_Affiliate_Banners set Banner_Text='"&Banner_Text&"',Banner_Image='"&Banner_Image&"',View_Order="&View_Order&" where Store_Id="&Store_Id & " and Banner_Id= " & Banner_Id
	  	     end if
	
		end if
		
		response.redirect "banners_list.asp"

end if
%>
