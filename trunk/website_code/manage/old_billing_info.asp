<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
server.execute "reset_design.asp"

sFormAction = "old_Store_payment.asp"
sFormName = "payment"
sTitle = "Billing Info"
thisRedirect = "billing_info.asp"
sMenu = "account"
createHead thisRedirect

	Billing_Type = request.querystring("Billing_Type")
	if Billing_Type = "Normal" then %>
		<TR bgcolor='#FFFFFF'><td colspan=2><BR><B>Your payment has been accepted and your service is now activated.</B>
	<% elseif Billing_Type = "Custom" then %>
		<TR bgcolor='#FFFFFF'><td colspan=2><BR><B>Your payment has been accepted and your service is now activated.</B>
	<% else %>
		<TR bgcolor='#FFFFFF'><td colspan=2><BR><B>Your payment information has been updated and will be used for all future billings.</B>

	<% end if
	
	sql_select = "select payment_id,sys_payments.amount,payment_description,payment_type from sys_payments inner join sys_billing on sys_payments.store_id=sys_billing.store_id where sys_billing.Store_Id="&Store_Id&" and datediff(hh,transdate,'"&now()&"')<1 order by transdate desc"
	rs_store.open sql_select, conn_store, 1, 1
	if not rs_store.eof then
	    sPayment_Id=rs_store("Payment_Id")
	    sAmount=rs_store("Amount")
	    sPayment_Description=rs_store("Payment_Description")
	    sPayment_type=rs_store("Payment_type")
	    sCity=rs_store("city")
	    sState=rs_store("state")
	    sCountry=rs_store("country")
	%>
	</form><form style="display:none;" name="utmform">
    <textarea id="utmtrans">
    <%  
    response.write "UTM:T|"&Store_Id&"||"&sAmount&"|0|0|"&sCity&"|"&sState&"|"&sCountry&" UTM:I|"&Store_Id&"|"&sPayment_Id&"|"&sPayment_Description&"|"&sPayment_type&"|"&sAmount&"|1" 
    %>
    </textarea>
    </form>
    <script type="text/javascript">
    __utmSetTrans();
    </script>
    <% end if
    rs_store.close %>


<BR><BR>Click here to return to the <a href=admin_home.asp class=link>home page</a>.
<BR><BR>Click here to view <a href=my_payments.asp class=link>prior invoices</a>.</td></tr>

<TR bgcolor='#FFFFFF'><td colspan=2><BR>**********************************************************
  </td></tr>
<script src="https://ssl.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-1343888-1";
_udn="none";
_ulink=1;
urchinTracker();
</script>

<!-- Google Code for purchase Conversion Page -->
<script language="JavaScript" type="text/javascript">
<!--
var google_conversion_id = 1062080517;
var google_conversion_language = "en_US";
var google_conversion_format = "1";
var google_conversion_color = "666666";
if (<%=sAmount %>) {
  var google_conversion_value = <%=sAmount %>;
}
var google_conversion_label = "purchase";
//-->
</script>
<script language="JavaScript" src="https://www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<img height=1 width=1 border=0 src="https://www.googleadservices.com/pagead/conversion/1062080517/?value=<%=sAmount %>&label=purchase&script=0">
</noscript>

                                    


<% createFoot thisRedirect, 0%>

