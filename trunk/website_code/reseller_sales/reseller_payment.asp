

<%
if trim(Session("tempresellerid")) ="" then 
	Response.redirect "../manage/error.asp?Message_id=94"
end if


title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="reseller_payment"
includejs=1


%>
<!--#include file = "header.asp"-->
<script language=javascript>
function show()
{
	document.frmtype.action ="reseller_billing_info.asp"
	document.frmtype.submit();
}
</script>

<table border="0" width=525 align=left bordercolor="#000066">

			<tr>
			<b>Reseller Program</b>
					
				</tr>

			<form method="POST" action="reseller_billing_info.asp" name="frmtype">
			 <TR>
                <TD class=inputname><B>Type</B></TD>
                <TD class=inputvalue><SELECT name=payment_method size=1>
                                  <OPTION selected value=Visa>Visa</OPTION> <OPTION 
                    value=Mastercard>Mastercard</OPTION> <OPTION 
                    value="American Express">American Express</OPTION> <OPTION 
                    value=Discover>Discover</OPTION> <OPTION value=eCheck>eCheck 
                    (US Only)</OPTION> </SELECT></TD>
                <TD class=inputvalue><IMG
                  src="images/icon_visa.gif"><IMG 
                  src="images/icon_mastercard.gif"><IMG 
                  src="images/icon_discover.gif"><IMG 
                  src="images/icon_amex.gif"><IMG 
                  src="images/echeck.jpg"></B> 
                </TD>
                </tr>
                

			<tr>

					<td colspan=2 align=left><br>
						<!--<input type="submit" border="0" value="Continue" onclick="javascript:show()"></td>-->
						<input type="submit" border="0" value="Continue" ></td>
				</tr>
						
			</form>


		</table>

<!--#include file="footer.asp"-->
