<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="help/coupon_settings.asp"-->
<%

if Show_Coupon <> 0 then
	 checkedenable = "checked"
else 
	 checkedenable = ""
end if

sFormAction = "Store_Settings.asp"
sName = "Coupon"
sFormName = "coupon_settings"
sTitle = "Coupon Settings"
sFullTitle = "Marketing > Coupons > Settings"
sCommonName="Coupon Settings"
sCancel="coupon_manager.asp"
sSubmitName = "Coupon"
thisRedirect = "coupon_settings.asp"
sMenu = "marketing"
createHead thisRedirect

if Service_Type < 7	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>

<% else %>

			 
				<TR bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Accept Coupons</td>
				<td width="70%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="Show_Coupon" value="-1" <%= checkedenable %>>&nbsp;
				<% small_help "Show Coupon" %></td>
				</tr>
         

<% createFoot thisRedirect,1 %>


<% end if %>
