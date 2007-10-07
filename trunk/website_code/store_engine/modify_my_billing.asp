<!--#include file="include/header.asp"--> 
	 
<table border="0" height="372" width="400">
	
	<tr>
		<td height="366" width="400">
			<table border="0" width="100%">
				<tr>
					<td colSpan="2" width="400" class='big'><b>Modify Billing Info</b></td>
				</tr>

				<tr>
					<td colSpan="2" width="400"></td>
				</tr>

				<% Record_Type = 0 %>
				<!--#include file="include/Display_Cust_Form.asp"-->
			</table>  
		</td> 

		
	</tr>
	<% response.write "<form name='Modify_Shipping' action='Cart_Action.asp?Return_To="&server.urlencode(fn_get_querystring("Return_To"))&"' method=post>" %>
	<tr><td align=center><%= fn_create_action_button ("Button_image_Checkout", "Check_Out", "Check Out")%></td></tr>
	</form>
</table>


<!--#include file="include/footer.asp"-->
