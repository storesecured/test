

<!--#include file="include\header.asp"-->
<!--#include file="include\user_register_info.asp"-->


<%
display_register_jstop

%>
<form method="POST" action="<%= Switch_Name %>include/register_Action.asp" name="Registration">
<% if "http://"&Request.ServerVariables("HTTP_HOST")&"/"=Site_Name then %>
        <Input type="hidden" name="Shopper_Id" value="<%=Shopper_Id%>">
<% end if %>
<input type="Hidden" name="Record_type" value="0">
<input type="Hidden" name="Redirect" value=<%= fn_get_querystring("ReturnTo") %>>

<TABLE WIDTH="100%" BORDER=0 CELLSPACING=1 CELLPADDING=1>

	<% If not (fn_get_querystring("ExpressCheckout")<>"" and ExpressCheckout) then %>
		<TR>
			<TD	colspan=3 width="400" class='normal'><b>Create an Account</b></TD>
		</TR>
		
		<TR>
			<TD class='normal' width='40%'>Login</TD>
			<TD colspan="2" class='normal' nowrap width='60%'>
				<INPUT	name=User_Id size=30 maxlength=50>
				<INPUT type="hidden"  name=User_Id_C value="Re|String|0|50|||Login">*</TD>
		</TR>
		
		<TR> 
			<TD class='normal'>Password</TD>
			<TD colspan="2" class='normal'>
				<INPUT name=Password type=password size=30 maxlength=50>
				<INPUT name=Password_C type=hidden value="Re|String|0|50|||Password" size=30>*</TD>
		</TR>
		
		<TR>
			<TD class='normal'>Confirm Password</TD>
			<TD colspan="2" class='normal'>
				<INPUT name=Password_Confirm type=password size=30 maxlength=50>
				<INPUT name=Password_Confirm_C type=hidden value="Re|String|0|50|||Confirm Password" size=30>*</TD>
		</TR>
		
		<% if AllowCookies=-1 then %>
			<tr>
				<td colspan="3" class='normal'>

					<input type="checkbox" name="SaveCookie" value="TRUE">
					Save login and password on this computer
				</td>
			</tr>
		<% end if %>
		
		<TR>
			<TD colspan="3"><hr></TD>
		</TR>
	
	<% Else %>
		<input type="hidden" name="ExpressCheckout" value="true">
	<% End If %>

	<TR>
		<TD width="400" colspan="3" class='normal'><b>Billing Information</b></TD>
	</TR> 
	
	
	<% display_register_info %>
	<% if show_tax_exempt then %>
	<TR>
		<TD class='normal'>Tax Exempt?</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=tax_exempt type=checkbox value="-1"></TD>
	</TR>
	<TR>
		<TD class='normal'>Federal Tax Id<BR><font class=small>(Required if tax exempt)</font></TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=tax_id type=textbox size=20></TD>
	</TR>
	<% else %>
           <INPUT name=tax_id type=hidden value="0">
           <INPUT name=tax_exempt type=hidden value="">
	<% end if %>
	<% if show_special_offers or show_special_offers=Null then %>
	<TR>
		<TD class='normal'><%= Email_Me %></TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Spam type=checkbox value="-1"></TD>
	</TR>
	<% else %>
           <INPUT name=Spam type=hidden value="0">
	<% end if %>
       
	<% if show_shipping then  %>
        <TR>
		<TD class='normal'><%= Shipping %> address same as billing address?</TD>
		<TD colspan="2" class='normal'>
			<INPUT	name=Shipping_Same type=checkbox value="-1" checked></TD>
	    </TR>
	
		<% if show_residential or show_residential=Null then %>
		<TR>
			<TD class='normal'>Residential Location</TD>
			<TD colspan="2" class='normal'>
				<INPUT	name=Is_Residential type=checkbox value="-1" checked></TD>
		</TR>
		<% else %>
		       <INPUT name=Is_Residential type=hidden value="-1">
		<% end if %>
	
	<%else %>
	<INPUT	name=Shipping_Same type=hidden value="-1">
	<INPUT name=Is_Residential type=hidden value="-1">
	<%end if%>
	
	<tr>
		<td>&nbsp;</td>
	</tr>

	<TR>
		<TD> </TD>
		<TD width="164" class='normal' colspan=2>

			<%= fn_create_action_button ("Button_image_Continue", "Register_User", "Continue") %>
			</TD>
	</TR>
</TABLE>

</form>
<% 
sSubmitName = "Register_User"
sFormName = "Registration"
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 <% If not (fn_get_querystring("ExpressCheckout")<>"" and ExpressCheckout) then %>
 frmvalidator.addValidation("User_Id","req","Please enter a login.");
 frmvalidator.addValidation("User_Id","alnumhyphen","Login can only contain letters and numbers");
 frmvalidator.addValidation("Password","req","Please enter a password.");
 frmvalidator.addValidation("Password","alnumhyphen","Password can only contain letters and numbers.");
 frmvalidator.addValidation("Password_Confirm","req","Please confirm the password.");

 <% end if %>

 <% display_register_jsbottom %>
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");
  function DoCustomValidation()
{
  var frm = document.forms["<%= sFormName %>"];
  <% If not (fn_get_querystring("ExpressCheckout")<>"" and ExpressCheckout) then %>
  if (frm.Password.value != frm.Password_Confirm.value)
    {
      alert('The 2 passwords must match.');
      return false
    }
  <% end if %>
	if (frm.State_UnitedStates.options[frm.State_UnitedStates.selectedIndex].value == "" && frm.Country.options[frm.Country.selectedIndex].value == "United States")
		{ 
		alert('Please choose a state.');
		return false
		}
	else if (frm.State_Canada.value == "" && frm.Country.options[frm.Country.selectedIndex].value == "Canada")
		{
		  alert('Please choose a province.');
		  return false
		}
  return true

}

</script>

<!--#include file="include/footer.asp"-->

