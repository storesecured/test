<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
if Enable_Rewards <> 0 then
	 checkedenable = "checked"
	 sDiv="block"
else 
	 checkedenable = ""
	 sDiv="none"
end if

sInstructions="A rewards program enables you to reward customers with a percentage of each purchase being placed into a rewards account which can be used for future purchases in your store.<BR><BR>Please note that rewards cannot be used for shipping or taxes."

sFormAction = "Store_Settings.asp"
sName = "Affiliates"
sFormName = "reward_settings"
sTitle = "Rewards"
sFullTitle = "Marketing > Rewards"
sCommonName = "Rewards"
sSubmitName = "Rewards"
thisRedirect = "reward_settings.asp"
sMenu = "marketing"
sQuestion_Path = "marketing/rewards.htm"
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

			 
				<tr bgcolor='#FFFFFF'><td width="30%" class="inputname"><b>Enable Rewards</b></td>
					<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Enable_Rewards" value="-1" <%= checkedenable %>>&nbsp;
					<% small_help "Enable Rewards" %></td>
				</tr>


				
				<tr bgcolor='#FFFFFF'>
				<td class="inputname" width="30%"><b>Minimum Use Balance</b>

					<td width="70%" class="inputvalue">
						<%= Store_Currency %><input name="Rewards_Minimum" type="input" value="<%= Rewards_Minimum %>" size=5 onKeyPress="return goodchars(event,'0123456789.')">
						<input type="hidden" name="Rewards_Minimum_C" value="Re|Integer||||Minimum Use Balance">
		  <% small_help "Rewards Minimum" %></td>
						</tr>

				<tr bgcolor='#FFFFFF'>
				<td class="inputname" width="30%"><b>Rewards Percent</b></td>

					<td width="70%" class="inputvalue"><input name="Rewards_Percent" type="input" value="<%= Rewards_Percent %>" size=5 onKeyPress="return goodchars(event,'0123456789.')">%
					<input type="hidden" name="Rewards_Percent_C" value="Re|Integer|||||Rewards Percent">
						 <% small_help "Rewards Percent" %></td>
						</tr>
         

<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Rewards_Minimum","req","Please enter a minimum use balance.");
 frmvalidator.addValidation("Rewards_Percent","req","Please enter a rewards percent.");
</script>

<% end if %>
