<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->


<% 
serverfree=freeserver

' ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><% 
else
on error resume next
	' RETRIVING FORM DATA
	if request.form("Accept_Terms") = "" then
		fn_error "You must accept the terms of service to create a store."
	end if
	Store_name = Replace(Request.Form("Store_name"), "'", "''")
	Site_Host = request.form("Site_Host")
   Site_name1 = Replace(Request.Form("Site_name"), "'", "")
	Site_name1 = LCase(Replace(Site_name1, " ", ""))
	Site_name = Site_name1 & "." & Site_Host
	if Site_name1 = "secure" or Site_name1 = "manage" or Site_name1 = "forum" or Site_name1 = "www" then
		fn_error "The Web Address, http://"&Site_name&", is already in use."
	end if
	if inStr(Site_Name1,",") > 0 then
		fn_error "The Web Address, http://"&Site_name&", cannot contain a comma."
	end if
	Secure_Name = Site_name1 & ".storesecured.com"
	Store_User_id = Replace(Request.Form("Store_User_id"), "'", "''")
	Store_Password = Replace(Request.Form("Store_Password"), "'", "''")
	Store_password_confirm = Replace(Request.Form("Store_password_confirm"), "'", "''")
	Store_Email = Replace(Request.Form("Email"), "'", "''")
	
	Service = Request.form("Service")
	if Service = "executive" and request.form("new_merch_acct") = "" then
		response.redirect "admin_error.asp?Message_Id=105"
	end if

	if Request.Form("Store_Password") <> Request.Form("Store_password_confirm") then
		response.redirect "admin_error.asp?Message_id=101"
	end if

	' CHECKING IF USER ID AlLREADY EXISTS IN STORE LOGINS TABLE
	sql_check = "select Login_Id from store_logins where Store_User_id = '"&Store_User_id&"' and Store_Password='"&Request.Form("Store_Password")&"'"
	rs_Store.open sql_check,conn_store,1,1
	if rs_store.bof = false then
		rs_Store.Close
		fn_error "The username you have choosen is already being used.  Please choose a new username."
	end if
	rs_Store.Close
	
	' CHECKING IF STORE NAME ALREADY EXISTS IN store_settings TABLE
	sql_check = "select Secure_Name from store_settings where Secure_Name = '"&Secure_Name&"' "
	rs_Store.open sql_check,conn_store,1,1
	if rs_store.bof = false then
		rs_Store.Close
		fn_error "The Web Address, http://"&Site_Name&", is already in use."
	end if
	rs_Store.Close

	' SET LOGIN PRIVILEGES TO 1 (ADMIN)
	Login_Privilege = 1

	if isNumeric(request.form("Affiliate")) then
		Affiliate = request.form("Affiliate")
	else
		Affiliate = 0
	end if
	
	Referrer = request.form("Referrer")
	
	resellerID = trim(Request.Form("hidResellerID"))

	sql_insert = "exec wsp_admin_store_create '"&Store_name&"','"&Store_User_id&"','"&Store_Password&"','"&Site_Name&"','"&Store_Email&"',"&Affiliate&",'"&Referrer&"',1,"&serverfree&","&resellerID&",'"&Secure_Name&"'"
	conn_store.Execute sql_insert

	sql_login = "select login_id, Store_id, Login_Privilege from store_logins where Store_User_Id = '"&Store_User_Id&"' and Store_Password = '"&Store_Password&"'"
	rs_Store.open sql_login,conn_store,1,1
	Store_id = rs_Store("Store_id")
	rs_Store.close
	
	sql_update = "exec wsp_design_copy "&store_Id&",-1,1;"
	conn_store.Execute(sql_update)
        
     sql_update = "exec wsp_design_apply "&store_Id&",-1;"
     conn_store.Execute(sql_update)


	server_id=freeserver

	if resellerID<> "" then 
		'*********************************************ends here**************************************************************
		sql =	" insert into  tbl_Reseller_Customer_Master(fld_Reseller_id,fld_Customer_ID,fld_plan_user_Reseller_id,fld_plan_Transaction_Date,fld_Amount_Pay_to_reseller,fld_status_flag)"&_ 
				" VALUES ("&resellerID&","&Store_id&",0,'"& now() &"',0,1)"
		conn_store.execute (sql)
	end if
	
	' CREATE STORE FOLDERS
	Set fso = CreateObject("Scripting.FileSystemObject")
    
    sPath = fn_get_sites_folder(Store_Id,"Base")
    if not fso.folderExists(sPath) then
        Set fldr = fso.CreateFolder(sPath)
        set copy_folder = fso.GetFolder(fn_get_sites_folder(101,"Base"))
        copy_folder.copy(sPath)
    end if

	' SET THE SESSION VARIABLE AND STATE
	Session("Store_id") = Store_id
	Session("Login_Privilege") = Login_Privilege

	Email_Text = "Congratulations, your store has been created and is now available at http://"&Site_Name&vbcrlf&vbcrlf&_
		"You can login to manage your store at http://manage."&Site_Host&" or from our homepage at "&_
		"http://www."&Site_Host&vbcrlf&vbcrlf&"Login: "&Store_User_id&vbcrlf&"Password: "&Store_Password&vbcrlf&vbcrlf&_
		"Should you have any questions regarding the features available, "&_
		"choosing a merchant accounts, upgrading your service or need any help at all please submit a support "&_
		"request.  Your store id is "&Store_id&", please reference this number when requesting support."&vbcrlf&vbcrlf&_
		"Thank you for choosing "&Site_Host&" for your ecommerce needs,"&vbcrlf&_
		"The "&Site_Host&" Team"
	Send_Mail sNoReply_email,Store_Email,Site_Host&" Store",Email_Text

	sType = "New"
	%>
    <script src="https://ssl.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-1343888-1";
_udn="none";
_ulink=1;
urchinTracker();
</script>

<%
end if
%>


