<!--#include file="include/header.asp"-->
	<font class='normal'>
	<form method="POST" action="<%= Site_Name %>rma_Action.asp" name="RMA">
     <table><tr><td colspan=2>Enter your last name and order id to get an RMA number.
     If your order was placed over <%= Auto_RMA_Days %> days ago your RMA request will be sent to the store admin.</td></tr>
     <tr><td>Order Id</td><TD><input type="text" name="Invoice" size="20" value="<%= fn_get_querystring("oid")%>">
							<INPUT name=Invoice_C type=hidden value="Re|Integer|||||Invoice"></td></tr>
     <tr><td>Last Name</td><TD><input type="text" name="Last_Name" size="20" value="<%= last_name %>">
							<INPUT name=Last_Name_C type=hidden value="Re|String|0|50|||Last Name"></td></tr>
     <tr><td>Reason for Return</td><TD><textarea name="Notes" rows=5 cols=30></textarea></td></tr>
     <tr><td><%= fn_create_action_button ("Button_image_Continue", "Continue", "Update") %></td></tr></table>
	</form>
     </font>


<!--#include file="include/footer.asp"-->
