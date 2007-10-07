<!--#include file="include/header.asp"-->

<form method="POST" action="<%=Switch_Name %>include/form_action.asp" name=Retrieve_cart>

<table border="0" width="77%">
	<tr>
		<td width="59%" class='normal'>Enter the wish list/cart id</td>
		<td width="60%"  class='normal'>
			<input type="text" name="Shopper_Id_Retrieve" size="20">
			<%= fn_create_action_button ("Button_image_RetrieveCart", "Retrieve_Cart", "Retrieve Cart") %>
			</td>
	</tr>
</table>

</form>
<%
sFormName = "Retrieve_cart"
sSubmitName = "Retrieve_Cart"
%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator("<%= sFormName %>");
 frmvalidator.addValidation("Shopper_Id_Retrieve","req","Please enter the id of the cart you wish to retrieve.");

</script>


<!--#include file="include/footer.asp"-->
