<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="editor_include.asp"-->

<%
payment_id=request.querystring("id")
if payment_id<>"" then
	sql_select = "select special_instructions,sys_payment_methods.Payment_method_id,payment_name,accept,store_payment_id, accept, payment_method_message from store_payment_methods left outer join sys_payment_methods on store_payment_methods.payment_method_id=sys_payment_methods.payment_method_id where store_id="&Store_Id&" and store_payment_id="&payment_id
	rs_Store.open sql_select,conn_store,1,1
	payment_name = rs_store("payment_name")
	store_payment_id = payment_id
	accept = rs_store("accept")
	Payment_method_id=rs_store("Payment_method_id")
	special_instructions=rs_store("special_instructions")
	payment_method_message = rs_store("payment_method_message")
	rs_Store.close
end if

if accept <> 0 then
	accept_Checked = "checked"
else
	accept_Checked = ""
end if

sTextHelp="paymentmethods/payment_methods.doc"

sFormAction = "Store_Settings.asp"
sName = "Store_Payment"
sFormName = "Store_Payment"
sCancel = "payment_manager.asp"
sCommonName = "Payment Method"
sTitle = "Payment Method"
sFullTitle = "General > Payments > <a href=payment_manager.asp class=white>Methods</a> > "
if payment_id<>"" then
	sFullTitle=sFullTitle &"Edit - "&payment_name
	sTitle="Edit "&sTitle&" - "&payment_name
else
	sFullTitle=sFullTitle &"Add"
	sTitle="Add "&sTitle
end if
sSubmitName = "Payment Methods"
thisRedirect = "payment_manager.asp"
sTopic = "Payment Method"
sMenu = "general"
createHead thisRedirect
%>


	<TR bgcolor='#FFFFFF'>
	<td width="25%" class="inputname">Payment Method Name</td>

	<td width="75%" class="inputvalue">
	<% if Payment_method_id="" then %>
		<input type="text" name="payment_name" value="<%= payment_name %>" size=60 maxlength=50>
		<input type="hidden" name="payment_name_C" value="Re|String|0|50|||Payment Name">
	<% else %>
		<%= payment_name %><input type="hidden" name="payment_name" value="<%= payment_name %>" size=60 maxlength=20>
	<% end if %>
	<% small_help "payment_name" %></td></tr>
	<% if special_instructions<>"" then %>
		<tr bgcolor='#FFFFFF'><td width="25%" class="inputname">Special Usage Applies</td>
		<td colspan=2 width="75%"><%=special_instructions %></td></tr>
	<% end if %>
	<tr bgcolor='#FFFFFF'><td width="25%" class="inputname">Enabled</td>
	<td width="75%"><input class="image" type="checkbox" name="Accept" value="-1" <%= accept_Checked %>><% small_help "accept" %></td></tr>

	<TR bgcolor=FFFFFF>
	<td class="inputvalue" colspan=2 width="100%"><b>Payment Message</b><BR>
		<%= create_editor ("payment_method_message",payment_method_message,"") %>
		<input type="hidden" name="payment_method_message_C" value="Op|String|||||payment_method_message">
	<% small_help "payment_method_message" %></td>
	</tr>
	<input type=hidden name=store_payment_id value="<%= store_payment_id %>">

<% createFoot thisRedirect, 1%>
