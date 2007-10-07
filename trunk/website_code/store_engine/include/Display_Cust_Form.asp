<!--#include file="user_register_info.asp"-->

<%
display_register_jstop

'BILLING / SHIPPING ADDRESS TO BE DISPLAYED IS SELECTED USING
'THE RECORD_TYPE PARAM
sql_select_cust = "exec wsp_customer_lookup "&store_id&","&cid&","&Record_type&";"
fn_print_debug sql_select_cust
rs_Store.open sql_select_cust,conn_store,1,1
if  not rs_store.eof then
	rs_Store.MoveFirst
	User_id= Rs_store("User_id")
	Password= Rs_store("Password")
	Company= Rs_store("Company")
	First_name= Rs_store("First_name")
	Last_name= Rs_store("Last_name")
	Address1= Rs_store("Address1")
	Address2= Rs_store("Address2")
	City= Rs_store("City")
	State = Rs_Store("State")
	Zip= Rs_store("Zip")
	Store_Country= Rs_store("Country")
	Phone= Rs_store("Phone")
	Fax= Rs_store("Fax")
	EMail= Rs_store("Email")
	Is_Residential= Rs_store("Is_Residential")
	rs_Store.Close
else
	rs_Store.Close
	Company= ""
	First_name= ""
	Last_name= ""
	Address1= ""
	Address2= ""
	City= ""
	Zip= ""
	Phone= ""
	Fax= ""
	EMail= ""
	Is_Residential= -1
end if

if Is_Residential<>0 then
   residential_checked="checked"
end if


%>
	
<form method="POST" action="include/update_records_action.asp" name=modify_customer>
<input type="Hidden" name="Page_id" value="<%= Page_id %>">
<input type="Hidden" name="Record_Type" value="<%= Record_Type %>">
<input type=hidden name=form_name value=modify_customer>


<%
display_register_info

if Record_Type<>0 then 
   if show_residential or show_residential=Null then %>
   <tr>
	<td class='normal'>Residential Location</td>
	<td class='normal'>
		<input type=checkbox <%=residential_checked %> name="Is_Residential" value="-1"></td>
  </tr>
     <% else %>
		       <INPUT name=Is_Residential type=hidden value="-1">
		<% end if %>
<% end if %>
<tr>
	<td colspan="2" class='normal' align=center>

		<% if Record_type=0 then %>
			<%= fn_create_action_button ("Button_image_Update", "Modify_my_0", "Update") %>
			<% sSubmitName = "Modify_my_0" %>
		<% Else %>
			<%= fn_create_action_button ("Button_image_Update", "Modify_my_1", "Update") %>
			<% sSubmitName = "Modify_my_1" %>
		<% End If %>
		</form>

		<% If allow_delete then %>
		<% response.write "<form name='Delete_Addr' action='include/update_records_action.asp' method=post>" %>
         <input type=hidden name=Form_Name value=Delete_Addr>
			<input type="hidden" name="Delete_Addr" value="<%= Record_type %>">
			<%= fn_create_action_button ("Button_image_Delete", "Delete_Ship_Addr", "Delete") %>
		 </form>
      <% End If %>
	</td>
</tr>


<%
      sFormName = "modify_customer"

      %>
      <SCRIPT language="JavaScript">
       var frmvalidator  = new Validator("<%= sFormName %>");
       <% display_register_jsbottom %>

      </script>

