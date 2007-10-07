<!--#include file="include/header.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/common_functions.asp"-->
<%

sFormName="Send_Friend"

q_Item_Page_Name = fn_get_querystring("Item_Page_Name")
q_Dept_Page_Name = fn_get_querystring("Dept_Page_Name")

if q_Item_Page_Name="" then
    Item_ID = fn_get_querystring("Item_ID")
    if isNumeric(Item_ID) and Item_ID<>"" then
        sql_select_items="exec wsp_item_transform "&store_id&","&Item_ID&",0;"
        session("sql")=sql_select_items
        fn_print_debug sql_select_items
        rs_store.open sql_select_items,conn_store,1,1
        if not rs_store.eof and not rs_store.bof then
            item_page_name=rs_store("item_page_name")
            full_name=rs_store("full_name")
            dept_page_name=replace(full_name," > ","/")
            sUrl=replace(fn_item_url(full_name,item_page_name),"-detail.htm","-friend.htm")
            fn_redirect_perm sUrl
        end if
    end if
    Item_Id=""
    sUrl=Site_Name
else
	full_name=fn_transform_deptstring (q_Dept_Page_Name)
	sUrl=fn_item_url(full_name,q_Item_Page_Name)

end if

if Item_Id<>"" then
    sMessage = "<a href='"&sUrl&"' class=link>Return to the item</a>"
end if

if request.form("With_message") <> "" then
        If Form_Error_Handler(Request.Form) <> "" then
	       Error_Log = Form_Error_Handler(Request.Form)
	       fn_error Site_Name&"form_error.asp?Error_Log="&server.urlencode(Error_Log)
        else
            Email_Address = request.form("Email_Address")
            Send_To = request.form("Send_To")
            With_Message = request.form("With_Message")
            Email_Copy = request.form("Email_Copy")
            Subject = request.form("Subject")
            
    	   With_Message = With_Message & vbcrlf&vbcrlf&sUrl
    
    	   Call Send_Mail(Email_Address,Send_To,Subject,With_Message)
    	   if Email_Copy <> "" then
    		Call Send_Mail(Email_Address,Email_Address,Subject,With_Message)
    	   end if
    	
    	   sMessage = "Your message has been successfully sent.<BR><BR>"&sMessage
    	end if
    	
    	%>
    	<table>
	<tr><td colspan=2><%= sMessage %></td></tr>
	</table>
<%
else
%>
	<form method=post action='' name="<%=sFormName %>">
        
	<table>
	<tr><td colspan=2><%= sMessage %></td></tr>
	<tr><td>Your Email</td><td>
        <input type=text size=60 name=Email_Address value='<%=Email %>'>
        <INPUT type="hidden"  name=Email_Address_C value="Re|Email|0|100|||Your Email">
        </td></tr>
	<tr><td>Friends Email</td><td>
        <input type=text size=60 name=Send_To>
        <INPUT type="hidden"  name=Send_To_C value="Re|Email|0|100|||Friends Email">
        </td></tr>
	<tr><td>Send yourself a copy</td><td>
        <input type=checkbox name=Email_Copy value=1>
        </td></tr>
     <tr><td>Subject</td><td>
        <input type=text size=60 name=Subject value='Check this out'>
        <INPUT type="hidden"  name=Subject_C value="Re|Subject|0|100|||Subject">
        </td></tr>
	<tr><td>Message</td><td>
        <textarea rows=5 cols=45 name=With_message>Type your message here</textarea>
        <INPUT type="hidden"  name=With_message_C value="Re|String|||||With Message">
        </td></tr>
	<tr><td colspan=2><%= fn_create_action_button ("Button_image_Continue", "Continue", "Send Now") %></td></tr>
	</table></form>
<%
end if

%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("Email_Address","req","Please enter your email address.");
 frmvalidator.addValidation("Email_Address","email","Your email address must be a valid email.");
 frmvalidator.addValidation("Send_To","req","Please enter your friends email.");
 frmvalidator.addValidation("Send_To","email","Your friends email must be a valid email.");
 frmvalidator.addValidation("With_message","req","You must type a message in order to continue.");
</script>
<!--#include file="include/footer.asp"-->
