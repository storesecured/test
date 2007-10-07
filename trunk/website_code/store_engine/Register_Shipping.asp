<!--#include file="include/header.asp"-->
<!--#include file="include/user_register_info.asp"-->


<%
first_name=""
last_name=""
address1=""
address2=""
city=""
zip=""
phone=""
email=""
fax=""
state=""
company=""

display_register_jstop

%>
<form method="POST" action="<%= Switch_Name %>include/register_Shipping_Action.asp" name="Registration">
<input type="Hidden" name="Record_type" value="1">
<input type="Hidden" name="Redirect" value=<%= fn_get_querystring("ReturnTo") %>>

<TABLE WIDTH="311" BORDER=0 CELLSPACING=1 CELLPADDING=1 style="height: 114px">
	<TR>
		<TD width="268" colspan="3" class='normal'><b>Enter <%= Shipping %> Information</b></TD>
	</TR>
	
	<% display_register_info %>
	<% if show_residential or show_residential=Null then %>
	<TR>
		<TD class='normal'>Residential Location</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Is_Residential type=checkbox value="-1" checked></TD>
	</TR>
	<% else %>
           <INPUT name=Is_Residential type=hidden value="-1">
	<% end if %>
	<TR>
		<TD width="131" class='normal'>	</TD>
		<TD width="164" class='normal' colspan=2>

		
			<%= fn_create_action_button ("Button_image_Continue", "Register_Shipping_User", "Continue") %>
			</TD>
	</TR>

</TABLE>
</form>
<%
sSubmitName = "Register_Shipping_User"
sFormName = "Registration" 
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 <% display_register_jsbottom %>

</script>

<!--#include file="include/footer.asp"-->
