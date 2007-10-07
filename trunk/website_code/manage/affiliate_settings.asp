<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

if Enable_affiliates <> 0 then
	 checkedenable = "checked"
	 sDiv="block"
else 
	 checkedenable = ""
	 sDiv="none"
end if

if Screen_affiliates <> 0 then
	 checkedscreen = "checked"
else 
	 checkedscreen = ""
end if

if Affiliate_type = 1 then
	 checked1 = "checked"
	 payPercent = Affiliate_amount
else
	 checked0 = "checked"
	 payAmount = Affiliate_amount
end if

sArticleHelp="affiliates.htm"
sInstructions="Affiliates are independent marketers who are paid only for performance."

sFormAction = "Store_Settings.asp"
sName = "Affiliates"
sFormName = "Affiliate_settings"
sTitle = "Marketing > Affiliates > Settings"
sSubmitName = "Affiliates"
thisRedirect = "affiliate_settings.asp"
sMenu = "marketing"
sQuestion_Path = "advanced/affiliate_program.htm"
createHead thisRedirect

if Service_type < 7 then %>
	<tr>
	<td colspan=2>
This feature is not available at your current level of service.<BR><BR>
		GOLD Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect,0 %>

<% else %>

				 <tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Enable Affiliates</b></td>
					<td width="60%" class="inputvalue" colspan=2><input class="image" type="checkbox" name="Enable_affiliates" value="-1" <%= checkedenable %>>&nbsp;
					<% small_help "Enable Affiliates" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname" rowspan=2><B>Type</B></td>
					<td class="inputvalue" width="20%"><input class="image" type="radio" name="Affiliate_type" value="1" <%= checked1 %>>Percent of sale</td><td class=inputvalue><input name="payPercent" type="input" value="<%= payPercent %>" size=5 onKeyPress="return goodchars(event,'0123456789.')">%
				  <input type="hidden" name="payPercent_C" value="Op|Integer|||||Payment Percent">
				  <% small_help "Pay %" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>

					<td class="inputvalue" width="20%"><input class="image" type="radio" name="Affiliate_type" value="0" <%= checked0 %>>Per click</td><td class=inputvalue><%= Store_Currency %><input name="payAmount" type="input" value="<%= payAmount %>" size=5 onKeyPress="return goodchars(event,'0123456789.')">
				  <input type="hidden" name="payAmount_C" value="Op|Integer|||||Payment Amount">
				  <% small_help "Pay Amount" %></td>
				</tr>


				<tr bgcolor='#FFFFFF'>
				<td class="inputname" width="40%"><b>Payment threshold</b></td>

					<td width="60%" class="inputvalue" colspan=2><%= Store_Currency %><input name="Affiliate_payout" type="input" value="<%= Affiliate_payout %>" size=5  onKeyPress="return goodchars(event,'0123456789.')">
					<input type="hidden" name="Affiliate_payout_C" value="Re|Integer|||||Payment Threshold">
						 <% small_help "Payment threshold" %></td>
						 </tr>
				 <tr bgcolor='#FFFFFF'>
				<td class="inputname" width="40%"><b>Return Duration</b></td>

					<td width="60%" class="inputvalue" colspan=2><input name="Affiliate_cookie" type="input" value="<%= Affiliate_cookie %>" size=5  onKeyPress="return goodchars(event,'0123456789')"> days
					<input type="hidden" name="Affiliate_cookie_C" value="Re|Integer|||||Return Duration">
						 <% small_help "Affiliate Cookie" %></td>
						 </tr>

<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Affiliate_payout","req","Please enter a payment threshold.");
 frmvalidator.addValidation("Affiliate_cookie","req","Please enter a return duration.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");

function DoCustomValidation()
{
  var frm = document.forms[0];
  if (document.forms[0].Affiliate_type[0].checked == true)
  {
    if (document.forms[0].payPercent.value == "")
    {
    alert('Please enter a percentage to pay affiliates.');
    return false;
    }
    else
    {
    return true;
    }
  }
  else
  if (document.forms[0].Affiliate_type[1].checked == true)
  { 
    if (document.forms[0].payAmount.value == "")
    {
    alert('Please enter a pay per click amount.');
    return false;
    }
    else
    {
    return true;
    }
  }
  else
  {

    return true;
  }
}


</script>


<% end if %>
