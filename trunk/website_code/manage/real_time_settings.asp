<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="help/real_time_settings.asp"-->

<script language="Javascript">
	var paymtds;
	function fnPaymtMethods(val1){
		if (val1=="G"){
			var paymtds = document.getElementById("paymtds").value;
		}else{
			paymtds = val1;
			if (paymtds=='undefined' || paymtds != 0){
				document.getElementById(val1).style.visibility = "visible";
				document.getElementById(val1).style.display="block";
			}
		}

		switch (paymtds){
			case "-1":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 4){
												document.getElementById(i).style.display="none";
												document.getElementById(i).style.visibility="hidden";
											}else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "0":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 0 && i != 4 && i != 38){
												document.getElementById(i).style.display="none";
												document.getElementById(i).style.visibility="hidden";
											}else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "1":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 1 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "2":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 2 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "4":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "5":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){									
											if (i != 5 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}	
									}	break;
			case "6":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){									
											if (i != 6 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "7":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){									
											if (i != 7 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "8":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){									
											if (i != 8 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "9":for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){									
											if (i != 9 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "10":	for(var i=0;i<39;i++){
										if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){										
											if (i != 10 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
										}
									}	break;
			case "11":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){									
											if (i != 11 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "12":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 12 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "13":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 13 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "14":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 14 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "15":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 15 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "16":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 16 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "17":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 17 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "18":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 18 && i != 4 && i != 38) {
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "19":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 19 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "20":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 20 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "22":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 22 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "23":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 23 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "24":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 24 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "26":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 26 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "28":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 28 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "29":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 29 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "31":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 31 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "32":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 32 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "35":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 35 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "36":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 36 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;			
			case "37":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 37 && i != 4 && i != 38){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;
			case "38":	for(var i=0;i<39;i++){
											if (i !=3 && i != 25 && i != 27 && i != 30 && i != 33 && i != 34){
											if (i != 38 && i != 4){
												document.getElementById(i).style.visibility="hidden";
												document.getElementById(i).style.display="none";
											}	else{
												document.getElementById(i).style.visibility = "visible";
												document.getElementById(i).style.display="block";
											}
											}
										}	break;							
		}	
	}
</script>
<%

if Real_Time_Processor = 0 then
	pchecked0 = "selected"
elseif Real_Time_Processor = 1 then
	pchecked_plugnpay = "selected"
elseif Real_Time_Processor = 2 then
	pchecked_authorize = "selected"
elseif Real_Time_Processor = 4 then
	pchecked_paypal = "selected"
elseif Real_Time_Processor = 5 then
	pchecked_psigate = "selected"
elseif Real_Time_Processor = 7 then
	pchecked_echo = "selected"
elseif Real_Time_Processor = 8 then
	pchecked_payflowlink = "selected"
elseif Real_Time_Processor = 9 then
	pchecked_payflowpro = "selected"
elseif Real_Time_Processor = 10 then
	pchecked_linkpoint = "selected"
elseif Real_Time_Processor = 11 then
	pchecked_2checkout = "selected"
elseif Real_Time_Processor = 12 then
	pchecked_internet = "selected"
elseif Real_Time_Processor = 14 then
	pchecked_bluepay = "selected"
elseif Real_Time_Processor = 15 then
	pchecked_electronic = "selected"
elseif Real_Time_Processor = 17 then
	pchecked_worldpay = "selected"
elseif Real_Time_Processor = 19 then
	pchecked_protx= "selected"
elseif Real_Time_Processor = 20 then
	pchecked_moneybookers= "selected"
elseif Real_Time_Processor = 22 then
	pchecked_nochex= "selected"
elseif Real_Time_Processor = 24 then
	pchecked_pri= "selected"
elseif Real_Time_Processor = 26 then
	pchecked_eft= "selected"
elseif Real_Time_Processor = 28 then
	pchecked_cybersource= "selected"
elseif Real_Time_Processor = 29 then
	pchecked_xor= "selected"
elseif Real_Time_Processor = 31 then
	pchecked_propay = "selected"
elseif Real_Time_Processor = 36 then
	pchecked_PayPal_Pro = "selected"
elseif Real_Time_Processor = 38 then
	pchecked_Google_Checkout = "selected"
end if

if Use_CVV2 then
	checked_cvv = "checked"
end if
if Auth_Capture then
	checked_auth = "checked"
end if
if Show_SecureLogo <>	0	then
	 Show_SecureLogo_Checked = "Checked"
else
	 Show_SecureLogo_Checked = ""
end if
if Paypal_Express <>	0	then
	 Paypal_Express_Checked = "Checked"
else
	 Paypal_Express_Checked = ""
end if
if Google_Checkout <>	0	then
	 Google_Checkout = "Checked"
else
	 Google_Checkout = ""
end if


sFormAction = "Store_Settings.asp"
sCommonName = "Payment Gateway"
sFormName = "Real_Time_Settings"
sTitle = "Payment Gateway"
sFullTitle = "General > Payments > Gateway"
sSubmitName = "Update_Store_Real_Time"
thisRedirect = "real_time_settings.asp"
sTopic = "real_time_settings"
AddPicker=1
sMenu="general"
sQuestion_Path = "advanced/payment_processor.htm"
createHead thisRedirect

sql_real_time = "exec wsp_real_time_property "&Store_Id&",-1;"
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_real_time,mydata,myfields,noRecords)
%>
<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="22">
          <table><tr><td width='70%'><B>Special PayPal Express Checkout Offer for StoreSecured Merchants</b><BR>
Free transaction processing on Express Checkout transactions until 1/31/2008.  Plus receive a $100 credit in a Yahoo! Sponsored Search account.
<a class=link target=_blank href="http://www.paypal-promo.com/searchmarketing/tracking/media_source/easystore.html">
Click here to Signup with Paypal Express Checkout</a></td><td width='30%' align=right><a href="http://www.paypal-promo.com/searchmarketing/tracking/media_source/easystore.html"><img src="http://www.paypal-promo.com/searchmarketing/images/banner/150x100_a.gif" border=0></a>
</td></tr></table>
		</td>
	 </tr>
<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="3" height="22">
			<input type="button" class="Buttons" value="Fraud Control" name="Back_To_Item_List" OnClick='JavaScript:self.location="maxmind_settings.asp"'>
			<input type="button" class="Buttons" value="Payment Methods" name="Add" OnClick='JavaScript:self.location="payment_manager.asp"'>
			<input type="button" class="Buttons" value="Find a Merchant Account" name="Add" OnClick='JavaScript:self.location="merchantacct.asp"'>
		</td>
	 </tr>
<TR bgcolor='#FFFFFF'><td class="inputname"><B>Req Card Code</B></td><td class="inputvalue"><input class="image" type="checkbox" name="Use_CVV2" value="-1" <%= checked_cvv %>>
<% small_help "Require Card Code" %></td></tr>
<TR bgcolor='#FFFFFF'><td class="inputname"><B>Auth and Deposit</B></td><td class="inputvalue"><input class="image" type="checkbox" name="Auth_Capture" value="-1" <%= checked_auth %>>
<% small_help "Authorize and Deposit" %></td></tr>
<TR bgcolor='#FFFFFF'>
<td width="30%" class="inputname"><B>Show Secure Seal</b></td>
<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Show_SecureLogo" value="-1" <%= Show_SecureLogo_Checked %>>&nbsp;
<% small_help "Show_SecureLogo" %></td>
</tr>

		<TR bgcolor='#FFFFFF'>
					<td width="100%" colspan="4" height="19" class=instructions>&nbsp;You must have an account with one of the following providers before enabling real-time processing.<BR>
					<BR>Read <a class="link" href="javascript:goArticle('creditcards.htm')">this</a>
				article about accepting credit cards</td>
				</tr>

				<TR bgcolor='#FFFFFF'><td colspan=3>&nbsp;</td></tr>
				<TR bgcolor='#FFFFFF'>
				<td class="inputname"><B>Payment Gateways</B></td>
				<td class="inputvalue">
				<select id='paymtds' name="Real_Time_Processor" onchange="fnPaymtMethods('G');">
				<option value = "-1">Select Payment Gateway</option>
				<option value = "0" <%=pchecked0%>>No Processor</option>
				<% if pchecked_2checkout="selected" then %>
				<option value = "11" <%=pchecked_2checkout%>>2Checkout</option>
				<% end if %>
				<option value = "2" <%=pchecked_authorize%>>Authorize.net</option>
				<option value = "14" <%=pchecked_bluepay%>>BluePay</option>
				<option value = "28" <%=pchecked_cybersource%>>Cybersource</option>
				<% if pchecked_echo="selected" then %>
				<option value = "7" <%=pchecked_echo%>>Echo</option>
				<% end if %>
				<% if pchecked_eft="selected" then %>
				<option value = "26" <%=pchecked_eft%>>EftNet</option>
				<% end if %>
				<% if pchecked_electronic="selected" then %>
				<option value = "15" <%=pchecked_electronic%>>eProcessing </option>
				<% end if %>
				<option value = "12" <%=pchecked_internet%>>Internet Secure</option>
				<option value = "10" <%=pchecked_linkpoint%>>LinkPoint</option>
				<% if pchecked_moneybookers="selected" then %>
				<option value = "20" <%=pchecked_moneybookers%>>MoneyBookers</option>
				<% end if %>
				<% if pchecked_nochex="selected" then %>
				<option value = "22" <%=pchecked_nochex%>>NoChex</option>
				<% end if %>
				<% if pchecked_payflowlink="selected" then %>
				<option value = "8" <%=pchecked_payflowlink%>>Payflow Link</option>
				<% end if %>
				<option value = "9" <%=pchecked_payflowpro%>>Payflow Pro</option>
				<% if pchecked_pri="selected" then %>
				<option value = "24" <%=pchecked_pri%>>Payment Resource</option>
				<% end if %>
				<option value = "4" <%=pchecked_paypal%>>Paypal Standard</option>
				<option value = "36" <%=pchecked_PayPal_Pro%>>PayPal Website Payment Pro</option>
				<option value = "1" <%=pchecked_plugnpay%>>PlugnPay</option>
				<% if pchecked_propay="selected" then %>
				<option value = "31" <%=pchecked_propay%>>Propay</option>
				<% end if %>
				<option value = "19" <%=pchecked_protx%>>Protx</option>
				<% if pchecked_psigate="selected" then %>
				<option value = "5" <%=pchecked_psigate%>>PSI Gate</option>
				<% end if %>
				<% if pchecked_worldpay="selected" then %>
				<option value = "17" <%=pchecked_worldpay%>>WorldPay</option>
				<% end if %>
				<% if pchecked_xor="selected" then %>
				<option value = "29" <%=pchecked_xor%>>Xor(Ashrait)</option>
				<% end if %>
				</select>
<!--				<option value = "" <%=pchecked_checksbynet%>>ChecksByNet</option> -->
				<% small_help "Payment Methods" %></td>
				  </tr>
				</table>

				<div id="0" style="display:none;visibility:hidden;">
					<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
						<TR bgcolor='#FFFFFF'>
						<td width="30%" class="inputname"><B>No Processor</B></td>
						<td width="70%" class="inputvalue">*<a class="link" href=noprocessor.asp<%= sAddString %>>Instructions about using No Processor</a>
						<% small_help "No Processor" %></td>
						</tr>
					</table>
				</div>

							<div id="1" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
									<TR bgcolor='#FFFFFF'>
									<td width="30%" class="inputname"><B>PlugnPay Settings</B></TD>
									<TD width="70%" class="inputvalue"><% small_help "Use_PnP" %></td></tr>
									<% 
									FOR rowcounter= 0 TO myfields("rowcount")
										if mydata(myfields("processor_id"),rowcounter) = 1 then %>
										<TR bgcolor='#FFFFFF'>
										<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
										<td width="82%" class="inputvalue">
												<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
										<% small_help mydata(myfields("property"),rowcounter) %></td>
										</tr>
										<% 
										else
											currRow = rowcounter
											exit for 
										end if
									Next %>
									<TR bgcolor='#FFFFFF'>
									<td height="21" colspan=4>&nbsp;</td>
									</tr>
							</table>
							</div>


							<div id="2" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
										<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B><a href=http://www.easystorecreator.com/authorize.asp target=blank class=link>Authorize.net</a> Settings</b></td>
										<td width="70%" class="inputvalue"><a class="link" href=authorizesetup.asp<%= sAddString %>>Instructions for setting up Authorize.Net</a>
                                                                                <% small_help "Use Authorize.net" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 2 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
							<%
								currRow = 5
							%>

							<div id="5" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=https://www.psigate.com/wizard/page0.asp?partner=1102703 class=link target=_blank
										onMouseOver="status='http://www.psigate.com'; return true" onMouseOut="status=''; return true">PSI Gate</a> Settings</B></td>
										<TD width="70%" class="inputvalue">* <a class="link" href=psi_certificate.asp<%= sAddString %>>Instructions for setting up PSI Gate</a>
										<% small_help "Use PSI Gate" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 5 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>

							<div id="6" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
											<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>iTransact Settings</B></td>
											<TD width="70%" class="inputvalue"><% small_help "Use iTransact" %></td></tr>
											<% 
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 6 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<% 
												else
													currRow = rowcounter
													exit for 
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
								</table>																																
								</div>

							<div id="7" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																									
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=http://www.echo-inc.com/echoapp.php?xecho=easyst class=link target=_blank>Echo</a> Settings</B></TD>
										<TD width="70%" class="inputvalue"><% small_help "Use_Echo" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 7 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
																																
							<div id="8" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
										<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B><a href=https://www.verisign.com/cgi-bin/go.cgi?a=p38230122366012000 class=link target=_blank
										onMouseOver="status='http://www.verisign.com'; return true" onMouseOut="status=''; return true">Payflow</a> Link Settings</B></td>
										<TD width="70%" class="inputvalue"><a class="link" href=payflowlink.asp<%= sAddString %>>Instructions for setting up Payflow Link</a>
                                                                                <% small_help "Use Payflow Link" %></td>
										</tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 8 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
								</table>
								</div>

							<div id="9" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
											<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B><a href=https://www.verisign.com/cgi-bin/go.cgi?a=p38230122366012000 class=link target=_blank
											onMouseOver="status='http://www.verisign.com'; return true" onMouseOut="status=''; return true">Payflow</a> Pro Settings</b></td>
											<TD width="70%" class="inputvalue"><% small_help "Use Payflow Pro" %></td></tr>
											<% 
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 9 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<% 
												else
													currRow = rowcounter
													exit for 
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
							</table>
							</div>

							<div id="10" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=http://www.easystorecreator.com/formmail_cardservice.asp class=link target=_blank
										onMouseOver="status='http://www.linkpoint.com'; return true" onMouseOut="status=''; return true">LinkPoint</a> Settings</B> </td><TD width="70%" class="inputvalue"><a class="link" href=linkpoint_certificate.asp<%= sAddString %>>Instructions for setting up LinkPoint</a><% small_help "Use Linkpoint" %></td></tr>
										<%
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 10 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
								</table>
								</div>
																																
							<div id="11" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>												
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=http://www.2checkout.com/cgi-bin/aff.2c?affid=74224 class=link target=_blank
										onMouseOver="status='http://www.2checkout.com'; return true" onMouseOut="status=''; return true">2Checkout</a> Settings</b></td>
										<TD width="70%" class="inputvalue"><a class="link" href=2checkout.asp<%= sAddString %>>Instructions for setting up 2Checkout</a>
                                                                                <% small_help "Use 2Checkout" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 11 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr> 
							</table>
							</div>

							<div id="12" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>											
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=https://www.internetsecure.com/prices.asp?ReferID=745 target=_blank class=link
										onMouseOver="status='http://www.internetsecure.com'; return true" onMouseOut="status=''; return true">Internet Secure</a> Settings</B></td>
										<TD width="70%" class="inputvalue"><% small_help "Use Internet Secure" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 12 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>

							<div id="13" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=https://aw.paysystems.com/aw.asp?B=35&A=144&Task=Click class=link target=_blank
										onMouseOver="status='http://www.paysystems.com'; return true" onMouseOut="status=''; return true">PaySystems</a> Settings</B></td>
										<TD width="70%" class="inputvalue"><a class="link" href=paysystems.asp<%= sAddString %>>Instructions for setting up PaySystems</a>
                                                                                <% small_help "Use PaySystems" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 13 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
								</table>
								</div>

							<div id="14" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=https://www.onlinedatacorp.com/onlineapp/entry.asp?hostID=504911 class=link target=_blank
										onMouseOver="status='http://www.onlinedatacorp.com'; return true" onMouseOut="status=''; return true">BluePay</a> Settings</B></td>
										<TD width="70%" class="inputvalue"><a class="link" href=bluepay.asp<%= sAddString %>>Instructions for setting up BluePay</a>
                                                                                <% small_help "Use BluePay" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 14 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
																																
							<div id="15" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>											
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href="http://qksrv.net/click-1334306-12796?url=http%3A%2F%2Fwww.electronictransfer.com" class=link target=_blank>eProcessing Network</a> Settings</B></td>
										<TD width="70%" class="inputvalue"><% small_help "Use Electronic Transfer" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 15 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
																																				
							<div id="16" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B>PayReady Settings</B></td>
										<TD width="70%" class="inputvalue"><% small_help "Use PayReady" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 16 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										 <TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>																																
							</div>

							<div id="17" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=https://secure.worldpay.com/app/splash.pl?Pid=62347 class=link target=_blank
										onMouseOver="status='http://www.worldpay.com'; return true" onMouseOut="status=''; return true">WorldPay</a> Settings</B></td>
										<TD width="70%" class="inputvalue"><a class="link" href=worldpay.asp<%= sAddString %>>Instructions for setting up WorldPay</a>
                                                                                <% small_help "Use WorldPay" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 17 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
																																
							<div id="18" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B>Settle Up</B> Settings</td>
										<TD width="70%" class="inputvalue"><a class="link" href=boa.asp<%= sAddString %>>Instructions for setting up Settle Up</a>
                                                                                <% small_help "Use BOA" %></td></tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 18 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
								</table>
								</div>
																																
								<div id="19" style="display:none;visibility:hidden;">
								<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
											<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>Protx Settings</B></td>
											<TD width="70%" class="inputvalue"><a class="link" href=protx.asp<%= sAddString %>>Instructions for setting up Protx</a>
                                                                                        <% small_help "Use Protx" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 19 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
								</table>
								</div>
																																
							<div id="20" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=https://www.moneybookers.com/app/?rid=291696 class=link
											onMouseOver="status='http://www.moneybookers.com'; return true" onMouseOut="status=''; return true">MoneyBookers</a> Settings</B></td>
											<TD width="70%" class="inputvalue"><% small_help "Use MoneyBookers" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 20 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
							</table>
							</div>
																																
								<div id="21" style="display:none;visibility:hidden;">
								<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
											<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>eWay Settings</B></td>
											<TD width="70%" class="inputvalue"><% small_help "Use eway" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 21 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
								</table>
								</div>

							<div id="22" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																																
											<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>NoChex Settings</B></td>
											<TD width="70%" class="inputvalue"><a class="link" href=nochex.asp<%= sAddString %>>Instructions for setting up NoChex</a><% small_help "Use NoChex" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 22 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
							</table>
							</div>
																																
							<div id="23" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>											
											<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>SecPay Settings</B></td>
											<TD width="70%" class="inputvalue"><% small_help "Use SecPay" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 23 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
							</table>
							</div>
																																
							<div id="24" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>											
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B>Use Payment Resource Settings</B></td>
										<TD width="70%" class="inputvalue"><% small_help "Use PRI" %></td></tr>
										<%
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 24 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<%
											else
												currRow = rowcounter
												exit for
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
																																
							<div id="26" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>											
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B>EftNet Settings</B></td>
										<TD width="70%" class="inputvalue"	<% small_help "Use EFT" %></td></tr>
										<%
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 26 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<%
											else
												currRow = rowcounter
												exit for
											end if
										Next %>
										<%
										on error goto 0
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 27 then
											else
												currRow = rowcounter
												exit for
											end if
										Next %>
											<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
																																
							<div id="28" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>													
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B>Cybersource Settings</B></td>
											<TD width="70%" class="inputvalue"><a class="link" href=cybersource_key.asp<%= sAddString %>>Instructions for setting up CyberSource </a>
                                                                                        <% small_help "Use Cybersource" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 28 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
										<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
								</table>
								</div>
																																
							<div id="29" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>												
											<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>Xor(Ashrait) Settings</B></td>
											<TD width="70%" class="inputvalue"><% small_help "Use Xor" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 29 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 30 then %>
												<% 
												else
													currRow = rowcounter
													exit for 
												end if
											Next %>
										<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr>
							</table>
							</div>
																																
							<div id="31" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>												
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href=http://epay.propay.com/cgi/appProcess1.exe/signup?NVXBIAFC class=link
											onMouseOver="status='http://www.propay.com'; return true" onMouseOut="status=''; return true">Propay</a> Settings</B></td>
											<TD width="70%" class="inputvalue"><% small_help "Use Propay" %></td></tr>
											<%
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 31 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<%
												else
													currRow = rowcounter
													exit for
												end if
											Next %>
										<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>

							<div id="32" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>												
										<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>Transecute Settings</B></td>
												<TD width="70%" class="inputvalue"><% small_help "Use transecutex" %></td>
										</tr>
											<% 
											FOR rowcounter= currRow TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 32 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<% 
												else
													currRow = rowcounter
													exit for 
												end if
											Next %>
										<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
							<%currRow = 54%>																			
							<div id="35" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
								<!-- SWISS NET -->
										<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B>SwissNet Settings</B></td>
												<TD width="70%" class="inputvalue"><% small_help "Use SwissNet" %></td>
												</tr>
												<% 
												FOR rowcounter= currRow TO myfields("rowcount")
													if mydata(myfields("processor_id"),rowcounter) = 35 then %>
													<TR bgcolor='#FFFFFF'>
													<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
													<td width="82%" class="inputvalue">
															<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
													<% small_help mydata(myfields("property"),rowcounter) %></td>
													</tr>
													<% 
													else
														currRow = rowcounter
														exit for 
													end if
												Next %>
								<!-- END SWISS NET-->
												<TR bgcolor='#FFFFFF'>
												<td height="21" colspan=4>&nbsp;</td>
											</tr>
							</table>
							</div>
																												
							
																												
					<div id="37" style="display:none;visibility:hidden;">
					<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>																												
						<!--PAYPAL WEBISTE PAYMENT PRO-->
								<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B>PayStation Settings</B></td>
										<TD width="70%" class="inputvalue"><a class="link" href=paystation.asp>Instructions for setting up PayStation </a>
                                                                                <% small_help "Use Paystation" %></td>
											</tr>
										<% 
										FOR rowcounter= currRow TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 37 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
						</table>
						</div>
						
						<div id="4" style="display:block;visibility:visible;">
						
						<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
						<TR bgcolor='#FFFFFF'><td colspan=3><HR></td></tr>
						<TR bgcolor='#FFFFFF'><td colspan=3>The below gateways are special payment method types that can be enabled in addition to your gateway if appropriate.
                                                To enable these payment types in addition to your gateway please select the Payment Methods button above and enable the respective payment method, ie Paypal</td></tr>
										<TR bgcolor='#FFFFFF'>
										<td width="30%" class="inputname"><B><a href="http://altfarm.mediaplex.com/ad/ck/3484-23890-3840-75" class=link target=_blank
										onMouseOver="status='http://www.paypal.com'; return true" onMouseOut="status=''; return true">Paypal Standard</a> Settings</B></td>
										<TD width="70%" class="inputvalue" colspan=3>
										</td></tr>
										<%
										FOR rowcounter= 3 TO myfields("rowcount")
											if mydata(myfields("processor_id"),rowcounter) = 4 then %>
											<TR bgcolor='#FFFFFF'>
											<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
											<td width="82%" class="inputvalue">
													<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
											<% small_help mydata(myfields("property"),rowcounter) %></td>
											</tr>
											<% 
											else
												currRow = rowcounter
												exit for 
											end if
										Next %>
										<TR bgcolor='#FFFFFF'>
										<td height="21" colspan=4>&nbsp;</td>
										</tr>
							</table>
							</div>
							
							<div style="display:block;visibility:visible;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
									<!--PAYPAL WEBISTE PAYMENT PRO-->
											<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B><a href="http://paypal.promotionexpert.com/PayPalMerchantGift/lead_gen_form.jsp?source=EASYSTORE" class=link>PayPal Pro</a> Settings</B></td>
													<TD width="70%" class="inputvalue"><a class=link href=http://paypal.promotionexpert.com/PayPalMerchantGift/lead_gen_form.jsp?source=EASYSTORE>Click here to signup</a><BR><a href=paypal_pro.asp class=link>Click here for Instructions</a>
                                                                                                        <% small_help "Use PaypalPro" %></td>
														</tr>
														<TR bgcolor='#FFFFFF'>
<td width="30%" class="inputname">Enable Express Checkout</td>
<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Paypal_Express" value="-1" <%= Paypal_Express_Checked %>>&nbsp;
<% small_help "Show_SecureLogo" %></td>
</tr>
													<%
													FOR rowcounter= 0 TO myfields("rowcount")
														if mydata(myfields("processor_id"),rowcounter) = 36 then %>
														<TR bgcolor='#FFFFFF'>
														<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
														<td width="82%" class="inputvalue">
																<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
														<% small_help mydata(myfields("property"),rowcounter) %></td>
														</tr>
														<% 
														else
															
														end if
													Next %>
													<!--<TR bgcolor='#FFFFFF'><td  class="inputvalue">&nbsp;</td>
													<td class="inputvalue" colspan=3 ><a class="link" href="paypalpro_certificate.asp">Upload PayPalPro certificate</a></td></tr>
													<TR bgcolor='#FFFFFF'>
													<td height="21" colspan=4>&nbsp;</td>
													</tr>-->
									<!--END PAYPAL PRO-->
													<TR bgcolor='#FFFFFF'>
													<td height="21" colspan=4>&nbsp;</td>
													</tr>
							</table>
							</div>
							<div style="display:block;visibility:show;">
                                                         <table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
									<!--Google Checkout-->
											<TR bgcolor='#FFFFFF'><td width="30%" class="inputname"><B>

                                              Google Checkout Settings</B></td>
													<TD width="70%" class="inputvalue">
                                                    <a class="link" href="http://checkout.google.com/sell?promo=sestoresecured">Click here to signup</a><BR>
                                                    <a class="link" href="google_setup.asp">Click here for instructions</a><BR>
                                                    
                                                                                                                                                            <% small_help "Use Google Checkout" %></td>
														</tr>
														<TR bgcolor='#FFFFFF'>
<td width="30%" class="inputname">Enable Google Checkout</td>
<td width="70%" class="inputvalue"><input class="image" type="checkbox" name="Google_Checkout" value="-1" <%= Google_Checkout%>>&nbsp;
<% small_help "Google_Checkout" %></td>
</tr>
	<TR bgcolor='#FFFFFF'>
<td width="30%" class="inputname">Google Checkout Button Style</td>
<td width="70%" class="inputvalue">
<% if GoogleCheckout_ButtonStyle=0 then %>
<input class="image" type="radio" name="GoogleCheckout_ButtonStyle" value="0" checked><b>White&nbsp;&nbsp;<input class="image" type="radio" name="GoogleCheckout_ButtonStyle" value="1">Transparent</b>
<%
else
%>
<input class="image" type="radio" name="GoogleCheckout_ButtonStyle" value="0"><b>White&nbsp;&nbsp;<input class="image" type="radio" name="GoogleCheckout_ButtonStyle" value="1" checked>Transparent</b>
	<% end if %>
<% small_help "Google_Checkout" %></td>
</tr>
													<%
													FOR rowcounter= 0 TO myfields("rowcount")
														if mydata(myfields("processor_id"),rowcounter) = 38 then %>
														<TR bgcolor='#FFFFFF'>
														<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
														<td width="82%" class="inputvalue">
																<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
														<% small_help mydata(myfields("property"),rowcounter) %></td>
														</tr>
														<% 

													else

														end if
													Next 
                                                                                                        %>

									<!--END Google Checkout-->
													<TR bgcolor='#FFFFFF'>
													<td height="21" colspan=4>&nbsp;</td>
													</tr>
							</table>
							</div>

							<div id="" style="display:none;visibility:hidden;">
							<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>											
										<TR bgcolor='#FFFFFF'>
											<td width="30%" class="inputname"><B>ChecksByNet Settings</B></td>
												<TD width="70%" class="inputvalue"></a>
											<% small_help "Use ChecksByNet" %></td>
										</tr>
											<% 
											FOR rowcounter= 53 TO myfields("rowcount")
												if mydata(myfields("processor_id"),rowcounter) = 33 then %>
												<TR bgcolor='#FFFFFF'>
												<td width="9%" class="inputname"><%= mydata(myfields("property"),rowcounter) %></td>
												<td width="82%" class="inputvalue">
														<input type="text" value="<%= decrypt(mydata(myfields("value"),rowcounter)) %>" name="<%= mydata(myfields("property"),rowcounter) %>" size="40">
												<% small_help mydata(myfields("property"),rowcounter) %></td>
												</tr>
												<% 
												else
													currRow = rowcounter
													exit for 
												end if
											Next %>
											<TR bgcolor='#FFFFFF'>
											<td height="21" colspan=4>&nbsp;</td>
											</tr> 
							</table>
							</div>

				<table border="0" width="100%" border='0' cellpadding=3 cellspacing=1>
					<TR bgcolor='#FFFFFF'><td colspan=3>&nbsp;</td></tr>
					<TR bgcolor='#FFFFFF'><td colspan=3><a class="link" href="merchantacct.asp" target=_blank>Click here for help choosing a merchant account.</a></td></tr>
			</table>

		<table border="0">	
<% createFoot thisRedirect,1
set myfields = Nothing %>
<script>fnPaymtMethods(<%= Real_Time_Processor %>);</script>
<Form Action="http://www.authorizenet.com/launch/ips_main.php" target="_blank" Name="EMS" method=post>
<input type="hidden" name="anUserID" value="easystorecreateapp"></Form>
