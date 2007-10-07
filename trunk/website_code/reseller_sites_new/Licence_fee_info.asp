

<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="Licence_fee_info"

%>
<!--#include file = "header.asp"-->
<script language=javascript>

</script>

<table border="0" width=525 align=left bordercolor="#000066">

			<tr>
			<b>Reseller Program</b>
					
				</tr>

			<form method="POST" action="" name="frm">
			 <TR> </tr>

			<tr>
	<%if request.querystring("Accepted") <> "" then 
	Billing_Type = request.querystring("Billing_Type")
	if Billing_Type = "Normal" then %>
		<td colspan=2><BR><B>Your Licence Fee has been accepted and your service is now activated.</B>
	<% elseif Billing_Type = "Custom" then %>
		<td colspan=2><BR><B>Your  Licence Fee has been accepted we will contact you if any further information is required before starting your custom work.</B>
	<% end if %>
	<td colspan=2 align=left><br>
		<input type="submit" border="0" value="Continue" onclick="javascript:show()" id=submit1 name=submit1></td>
	</tr>
	<input type="text" value="<%=payment_method%>" >
	
	</form>


		</table>

<!--#include file="footer.asp"-->




