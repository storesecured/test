<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
if Session("Login_Privilege") <> 1 then
   response.redirect "noaccess.asp"
end if

if Request.QueryString("Delete_Id") <> "" then
	Delete_Id = request.querystring("Delete_Id")
elseif Request.Form("Delete_Id") ="SEL" and request.form("DELETE_IDS") <> "" then
	Delete_Id = request.form("DELETE_IDS")
else
	Delete_Id = ""
end if

if Delete_Id<>"" then

        sql_select = "select login_privilege from store_logins where store_id="&store_id&" and login_id in ("&delete_id&") and login_privilege=1"
        rs_Store.open sql_select,conn_store,1,1
        if not rs_store.EOF then
                login_privilege=rs_store("login_privilege")
                if login_privilege="1" then
                        fn_error "The admin login cannot be removed."
                end if

        end if

end if

sFlashHelp="logins/logins.htm"
sMediaHelp="logins/logins.wmv"
sZipHelp="logins/logins.zip"

sQuestion_Path = "advanced/login_1.htm"

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Logins"
myStructure("ColumnList") = "login_id,store_user_id,store_first_name,store_last_name,store_email,login_privilege"
myStructure("HeaderList") = "store_user_id,store_first_name,store_last_name,store_email,login_privilege"
myStructure("DefaultSort") = "store_user_id"
myStructure("PrimaryKey") = "login_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "account"
myStructure("FileName") = "security.asp"
myStructure("FormAction") = "store_settings.asp"
myStructure("Title") = "Admin Logins"
myStructure("FullTitle") = "My Account > Admin Logins"
myStructure("CommonName") = "Logins"
myStructure("NewRecord") = "security_add.asp"
myStructure("Heading:login_id") = "PK"
myStructure("Heading:store_first_name") = "First"
myStructure("Heading:store_last_name") = "Last"
myStructure("Heading:store_email") = "Email"
myStructure("Heading:store_user_id") = "Username"
myStructure("Heading:login_privilege") = "Priviledge"
myStructure("Format:store_first_name") = "STRING"
myStructure("Format:store_last_name") = "STRING"
myStructure("Format:store_email") = "STRING"
myStructure("Format:store_user_id") = "STRING"
myStructure("Format:login_privilege") = "LOOKUP"
myStructure("Length:coupon_amount") = 2
myStructure("Length:coupon_total_left") = 0
myStructure("Lookup:login_privilege") = "1:Admin^2:Regular"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->


<% createFoot thisRedirect, 0%>
