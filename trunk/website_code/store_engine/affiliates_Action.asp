<!--#include file="include/header.asp"-->
<%

' ERROR CHECKING

If not CheckReferer then
	response.redirect Site_Name&"error.asp?message_id=1"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	response.redirect Site_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
else

	' ADD A NEW AFFILIATE

		' RETRIEVING FORM DATA
		Code = checkStringForQ(request.form("Code"))
		Contact_Name = checkStringForQ(request.form("Contact_Name"))
		Email = checkStringForQ(request.form("Email"))
		URL = checkStringForQ(request.form("URL"))
		Password = checkStringForQ(request.form("Password"))
		Password_Confirm = checkStringForQ(request.form("Password_Confirm"))
		Zip = checkStringForQ(Request.Form("Zip"))
    	State = checkStringForQ(Request.Form("State"))
    	State_Opt = checkStringForQ(Request.Form("State_Opt"))
	   Company = checkStringForQ(Request.Form("Company"))
    	Address = checkStringForQ(Request.Form("Address"))
   	City = checkStringForQ(Request.Form("City"))
    	Country = checkStringForQ(Request.Form("Country"))
      if (Country<>"United States" and Country<>"Canada") and State_Opt<>"" then
		     State=State_Opt
	    end if
      Phone = checkStringForQ(Request.Form("Phone"))

		' CHECK PASSWORD AND PASSWORD CONFIRMATION
		if (Password<>Password_Confirm) then
			response.redirect Site_Name&"error.asp?message_id=64"
		end if
		
		' CHECK IF ADDRESS STARTS WITH HTTP://
		if (len(URL)<7) then
			response.redirect Site_Name&"error.asp?message_id=71"
		else
			if ucase((mid(URL,1,7)))<>"HTTP://" then
				response.redirect Site_Name&"error.asp?message_id=71"
			end if
		end if
		
		' CHECK EMAIL VALIDITY
		if Instr(1, Email, "@") = 0 or	Instr(1, Email, ".") = 0 then
			response.redirect Site_Name&"error.asp?message_id=72"
		end if

		' CHECK IF AFFILIATE CODE EXIST
		sql_check = "select Affiliate_ID from Store_Affiliates where Code='"&Code&"' and store_id="&Store_id

		rs_Store.open sql_check,conn_store,1,1
		if not rs_store.eof then
			rs_store.close
			response.redirect Site_Name&"error.asp?message_id=65"
		end if
		rs_store.close

		if Screen_affiliates then
			Approved = 0
		Else
			 Approved = -1
		End if
		'INSERT INTO THE DATABASE
		sql_ins = "insert into Store_Affiliates (Store_ID, Code, Contact_Name, Email, URL, Password, Email_Notification,Approved,Zip,State,Company,Address,City,Country,Phone) values ("&store_id&", '"&Code&"', '"&Contact_Name&"', '"&Email&"', '"&URL&"', '"&Password&"', -1,"&Approved&",'"&Zip&"','"&State&"','"&Company&"','"&Address&"','"&City&"','"&Country&"','"&Phone&"')"
		conn_store.execute sql_ins
		
		%>
		<table><tr><td class='normal'>Congratulations you are now an affiliate.  Click <a href="http://affiliates.easystorecreator.com/affiliates_login.asp?Store_id=<%= store_id %>" class='link'>here</a> to login and get links etc.
		</td></tr></table><%


end if



%> 
<!--#include file="include/footer.asp"-->

