<!--#include file="Global_settings.asp"-->
<!--#include file = "pagedesign.asp"-->
<%
	sFormAction = "customer_mass_edit_action.asp"
	sTitle="Mass Edit Customers"
	sSubmitName = "Mass_Edit_Update"
	sFormName = "Mass_Edit_Update"
	thisRedirect = "customer_Mass_Edit.asp"
	sMenu = "customers"
	createHead thisRedirect

	Special_start_date=FormatDateTime(now(),2)
	Special_end_date=Dateadd("m",1,Special_start_date)


if Service_Type < 7 then %>
	<tr bgcolor='FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.
	</td></tr>

<% createFoot thisRedirect, 0%>
<% else %>
	<!-- getting the values of the ids -->
		<input type="hidden" name="DELETE_IDS" value="<%= request.form("DELETE_IDS") %>">
		
		<tr bgcolor='FFFFFF'>
			<td width="20%" class="inputname">Budget</td>
			<td width="80%" class="inputvalue">
						<input name="Budget" size="10">
						<INPUT type="hidden"  name="Budget_C" value="Op|Integer|||||Budget">
						<% small_help "Original Budget Amount" %></td>
		</tr>


		<tr bgcolor='FFFFFF'>
			<td width="20%" class="inputname">Budget Left</td>
			<td width="80%" class="inputvalue">
						<input name="Budget_Left" size="10" >
						<INPUT type="hidden"  name="Budget_Left_C" value="Op|Integer|||||Budget Left ">
						<% small_help "Budget Amount Left" %></td>
		</tr>

				
		<tr bgcolor='FFFFFF'>
			<td width="20%" class="inputname">Promo Emails</td>
			<td width="80%" class="inputvalue">
			<input class="image" type="radio" name="spam" value="-1">Yes&nbsp;
				<input class="image" type="radio" name="spam" value="0">No&nbsp;
				<input class="image" type="radio" name="spam" value="" checked>don't update&nbsp;
<!--				<input type="checkbox" name="spam" >-->
				<% small_help "Promotional Emails" %>						
			</td>
		</tr>


		<tr bgcolor='FFFFFF'>
			<td width="20%" class="inputname">Tax Exempt </td>
			<td width="80%" class="inputvalue">
				<input class="image" type="radio" name="tax_exempt" value="-1">Yes&nbsp;
				<input class="image" type="radio" name="tax_exempt" value="0">No&nbsp;
				<input class="image" type="radio" name="tax_exempt" value="" checked>don't update&nbsp;
<!--				<input type="checkbox" name="tax_exempt" >-->
				<% small_help "Tax Exempt" %>
			</td>
		</tr>

		<tr bgcolor='FFFFFF'>
			<td width="20%" class="inputname">Protected Page Access</td>
			<td width="80%" class="inputvalue">
				<input class="image" type="radio" name="Protected_Page_Access" value="-1">Yes&nbsp;
				<input class="image" type="radio" name="Protected_Page_Access" value="0">No&nbsp;
				<input class="image" type="radio" name="Protected_Page_Access" value="" checked>don't update&nbsp;
<!--				<input type="checkbox" name="Protected_Page_Access" >-->
				<% small_help "Access to Protected Pages" %>
			</td>
		</tr>


		
		<% createFoot thisRedirect,1%>
<% end if %>
