<!--#include file="include/header.asp"-->
<!--#include file="include/cart_display.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/crypt.asp"-->
<%
Server.ScriptTimeout =1200

GGrand_Total = 0

'RETRIVE INFORMATION FROM STORE_PURCHASES TABLE
response.Write fn_display_invoice (oid,shopper_id,1,0)

sql_select = "select max(payment_method) as payment_method,sum(grand_total) as grand_total from store_purchases WITH (NOLOCK) where store_id="&store_id&" and (oid="&oid&" or masteroid="&oid&")"
session("sql")=sql_select
fn_print_debug sql_select
rs_Store.open sql_select,conn_store,1,1
if not rs_store.eof and not rs_store.bof then
    Payment_Method=rs_store("Payment_Method")
    Grand_Total=rs_store("Grand_Total")
end if
rs_store.close
sPaymentText=""
iLocalPost=1

if fn_is_realtime(Payment_Method)=1 and (fn_is_third_party_gateway(Real_Time_Processor)=1 or fn_is_paypal_express(Payment_Method,Real_Time_Processor) or fn_is_paypal(Payment_Method,Real_Time_Processor)) then
    fn_print_debug "in include files if"
    iLocalPost=0
    if fn_is_paypal(Payment_Method,Real_Time_Processor)=1 then
		%><!--#include file="include/paypal/paypaldisplay.asp"--><%
		call create_form_post_paypal
   	elseif Real_Time_Processor = 8 then
		%><!--#include file="include/verisign/payflowlinkdisplay.asp"--><%
		call create_form_post_payflowlink
	elseif Real_Time_Processor = 11 then
		%><!--#include file="include/2checkout/2checkoutdisplay.asp"--><%
		call create_form_post_2checkout
	elseif Real_Time_Processor = 12 then
		%><!--#include file="include/internetsecure/internetsecuredisplay.asp"--><%
		call create_form_post_internet
	elseif Real_Time_Processor = 17 then
		%><!--#include file="include/worldpay/worldpaydisplay.asp"--><%
		call create_form_post_worldpay
	elseif Real_Time_Processor = 19 then
		%><!--#include file="include/protx/protxdisplay.asp"--><%
		call create_form_post_protx
	elseif Real_Time_Processor = 20 then
		%><!--#include file="include/moneybookers/moneybookersdisplay.asp"--><%
		call create_form_post_moneybookers
	elseif Real_Time_Processor = 22 then
		%><!--#include file="include/nochex/nochexdisplay.asp"--><%
		call create_form_post_nochex
	elseif Real_Time_Processor = 24 then
		%><!--#include file="include/pri/pridisplay.asp"--><%
		call create_form_post_pri
	elseif fn_is_paypal_express(Payment_Method,Real_Time_Processor)=1 then
		%><!--#include file="include/paypalpro/processPaypalExpress.asp"--><%
		call create_Form_Paypal
	end if
else
	sPaymentText = sPaymentText & ("<form method='POST' action='"&Switch_Name&"include/check_out_payment_action.asp' name='Payment' onSubmit=""return checkFields();"">")
end if
sPaymentText = sPaymentText & ("<input type=hidden name=Form_Name value=Payment_Process_Order>")

sCoupon = Coupon
if isNull(sCoupon) or sCoupon="" then
	sCoupon = "Coupon"
end if

sPaymentText = sPaymentText & ("<TABLE WIDTH='100%' BORDER=0 CELLSPACING=1 CELLPADDING=1><TR>"&_
    "<TD colspan='2' class='normal'><input type='hidden' name='Payment_Method' value='"&Payment_Method&"'>"&_
	"Payment Method <b>"&Payment_Method&"</b></font><br>")
if Payment_Method = "Charge my account" then
    sPaymentText = sPaymentText & ("Current Outstanding Budget <b>"&Currency_Format_Function(Budget_left)&_
        "</b></font><br>")
End If
sPaymentText = sPaymentText & ("</TD>")
if strServerPort= 443 and Show_SecureLogo then
    sPaymentText = sPaymentText & ("<td rowspan=6><!-- Secure ID - Do Not Remove! -->"&_
        "<div id='digicertsitesealcode' style='width: 81px; margin: 0px auto 0px 0px;' align='center'>"&_
        "<script language=""javascript"" type=""text/javascript"" "&_
        "src=""https://www.digicert.com/custsupport/sealtable.php?order_id="&ssl_oid&_
        "&seal_type=a&seal_size=large""></script><script language=""javascript"" type=""text/javascript"">"&_
        "coderz();</script></div><!-- END Secure ID --></td>")
end if

response.Write sPaymentText
sPaymentText=""

sPayments="Visa,Mastercard,American Express,Discover,Diners Club,eCheck,Paypal,Debit Card,Solo,Switch,Maestro"
if Is_In_Collection(sPayments,Payment_Method,",") then
    if fn_is_paypal(Payment_Method,Real_Time_Processor)=1 then
		fn_print_debug "in is paypal"
		call create_form_content_paypal
    elseif Real_Time_Processor = 8 then
		call create_form_content_payflowlink
	elseif Real_Time_Processor = 11 then
		call create_form_content_2checkout
	elseif Real_Time_Processor = 12 then
		 call create_form_content_internet
	elseif Real_Time_Processor = 17 then
		 call create_form_content_worldpay
	elseif Real_Time_Processor = 19 then
		 call create_form_content_protx
	elseif Real_Time_Processor = 20 then
		 call create_form_content_moneybookers
	elseif Real_Time_Processor = 22 then
		 call create_form_content_nochex
	elseif Real_Time_Processor = 24 then
		 call create_form_content_pri
	else
	    if Payment_Method = "eCheck" then
		    sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>&nbsp;</TD>"&_
			    "<TD width='219' class='normal'>&nbsp;</TD></TR><TR>")
		    if Real_Time_Processor<>14 then
			    sPaymentText = sPaymentText & ("<TD width='176' class='normal'>Bank Name</TD>"&_
			        "<TD width='219' class='normal'><INPUT type='text' name='BankName' size='28'>"&_
			        "<INPUT name='BankName_C' type='hidden' value='Re|String|0|50|||Bank Name'></TD>")
		    else
		        sPaymentText = sPaymentText & ("<TD width='176' class='normal'>Name on Account</TD>"&_
			        "<TD width='219' class='normal'><INPUT type='text' name='CardName' size='28'>"&_
			        "<INPUT name='CardName_C' type='hidden' value='Re|String|0|50|||Card Name'></TD>")
            end if
            sPaymentText = sPaymentText & ("</TR><TR><TD width='176' class='normal'>Bank Routing Number</TD>"&_
			    "<TD width='219' class='normal'><INPUT type='text' name='BankABA' size='28'>"&_
			    "<INPUT name='BankABA_C' type='hidden' value='Re|String|0|50|||Bank Routing Number'></TD></TR>"&_
		        "<TR><TD colspan=2 valign=top><font size=1>Bank Routing Number is the first 9 numbers on the "&_
		        "bottom left hand corner of a check</font></td></tr><TR>"&_
		        "<TD width='176' class='normal'>Account Num</TD><TD width='219' class='normal'>"&_
		        "<INPUT type='text' name='BankAccount' size='28'>"&_
				"<INPUT name='BankAccount_C' type='hidden' value='Re|String|0|252|||Account Num'></TD></TR>")
			if Real_Time_Processor<>14 then
				sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>Account Type</TD>"&_
				    "<TD width='219' class='normal'><select name='acct_type' size='1'>"&_
					"<option selected value='CHECKING'>Checking</option>"&_
					"<option value='SAVINGS'>Savings</option></select>"&_
					"<select name='org_type' size='1'><option selected value='I'>Individual</option>"&_
					"<option value='B'>Business</option></select></TD></TR>")
			    if Real_Time_Processor = "2" then
				    sPaymentText = sPaymentText & ("<TR><TD class='normal' colspan=2>"&_
				        "Enter either Tax id or drivers license info</TD></TR>"&_
                        "<TR><TD width='176' class='normal'>Tax Id or SSN</TD>"&_
					    "<TD width='219' class='normal'><INPUT type='text' name='TaxID' size='28'>"&_
					    "<INPUT name='TaxID_C' type='hidden' value='Op|String|0|252|||Tax Id'></TD></TR>")
				else
					sPaymentText = sPaymentText & ("<TD width='176' class='normal'>Check Serial #</TD>"&_
					    "<TD width='219' class='normal'><INPUT type='text' name='CheckSerial' size='28'>"&_
					    "<INPUT name='CheckSerial_C' type='hidden' value='Op|String|0|252|||Check Serial #'>"&_
					    "</TD></TR>")
				end if
			    sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>Drivers License State/Num</TD>"&_
				    "<TD width='219' class='normal'><INPUT type='text' name='DrvState' size='2'>"&_
				    "<INPUT name='DrvState_C' type='hidden' value='Op|String|0|252|||License State' size='2'>"&_
				    "<INPUT type='text' name='DrvNumber' size='15'>"&_
				    "<INPUT name='DrvNumber_C' type='hidden' value='Op|String|0|252|||License Number'></TD></TR>")
				Current_year = year(now())
				if Real_Time_Processor = "7" then
					sTitle = "Expiration Date"
					iYearStart = Current_year
					iYearEnd = Current_year + 20
				else
					sTitle = "Birth Date"
					iYearStart = Current_year - 100
					iYearEnd = Current_year - 5
				end if
				sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>"&sTitle&"</TD>"&_
					"<TD width='219' class='normal'><select name='dobm' size='1'>"&_
					"<option selected value='01'>January</option><option value='02'>February</option>"&_
					"<option value='03'>March</option><option value='04'>April</option>"&_
					"<option value='05'>May</option><option value='06'>June</option>"&_
					"<option value='07'>July</option><option value='08'>August</option>"&_
					"<option value='09'>September</option><option value='10'>October</option>"&_
					"<option value='11'>November</option><option value='12'>December</option>"&_
					"</select><INPUT type='text' name='dobd' size='2' maxlength='2'>"&_
					"<INPUT name='dob_d_C' type='hidden' value='Op|integer|1|31|||Day'>"&_
					"<select name='doby' size='1'>")
				for go_year = iYearStart to iYearEnd
				    sPaymentText = sPaymentText & ("<option value='"&go_year&"'>"&go_year&"</option>")
				next
				sPaymentText = sPaymentText & ("</select></TD></TR>")
			end if
		else
		    sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>&nbsp;</TD>"&_
			    "<TD width='219' class='normal'>&nbsp;</TD></TR><TR>"&_ 
				"<TD width='176' class='normal'>Name on Card</TD><TD width='219' class='normal'>"&_
				"<INPUT type='text' name='CardName' size='28' maxlength='50'>"&_
				"<INPUT name='CardName_C' type='hidden' value='Re|String|0|50|||Name on Card'>"&_
				"</TD></TR><TR> <TD width='176' class='normal'>Card Number</TD>"&_
				"<TD width='219' class='normal'><INPUT type='text' name='CardNumber' size='28' "&_
				"maxlength=30 onKeyPress=""return goodchars(event,'0123456789')"">"&_
				"<INPUT name='CardNumber_C' type='hidden' value='Op|String|0|252|||Card Number'>"&_
				"</TD></TR>")
            sJava = "frmvalidator.addValidation(""CardName"",""req"",""Please enter the name on the credit card."");"
            sJava = sJava & "frmvalidator.addValidation(""CardNumber"",""req"",""Please enter the credit card number."");"
            if Use_CVV2 then
				sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>Card Code</TD>"&_
					"<TD width='219' class='normal'><table><tr><td>"&_
                    "<INPUT type='text' name='CardCode' size='4' "&_
                    "onKeyPress=""return goodchars(event,'0123456789')"">"&_
					"<INPUT name='CardCode_C' type='hidden' value='Re|String|0|4|||Card Code'>"&_
					"</td><td><a href=""JavaScript:goCardCode();"" class='small'>"&_
					"<img src='images/images_themes/mini_cvv2.gif' border=0></a></td><td>"&_
					"<a href=""JavaScript:goCardCode();"" class='link'>What's This?</a></td>"&_
					"</tr></table></TD></TR>")
				sJava = sJava & "frmvalidator.addValidation(""CardCode"",""req"",""Please enter the card code."");"
            end if
			sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>Expiration Date</TD>"&_
			    "<TD width='219' class='normal'><select name='mm' size='1'>"&_
				"<option selected value='01'>January</option><option value='02'>February</option>"&_
				"<option value='03'>March</option><option value='04'>April</option>"&_
				"<option value='05'>May</option><option value='06'>June</option>"&_
				"<option value='07'>July</option><option value='08'>August</option>"&_
				"<option value='09'>September</option><option value='10'>October</option>"&_
				"<option value='11'>November</option><option value='12'>December</option></select>") 
			Current_year = year(now())
			sPaymentText = sPaymentText & ("<select name='yy' size='1'>")
			for go_year = Current_year to Current_year + 10
			    go_year_two_digit = right(go_year,2)
				sPaymentText = sPaymentText & ("<option value='"&go_year_two_digit&"'>"&go_year&"</option>")
			next
			sPaymentText = sPaymentText & ("</select></TD></TR>"&_
			    "<TR><TD width='176' class='normal'>&nbsp;</TD>"&_
			    "<TD width='219' class='normal'>&nbsp;</TD></TR>")
			if Payment_Method="Solo" or Payment_Method="Switch" or Payment_Method="Maestro" then
				sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>Issue Date</TD>"&_
                    	"<TD width='219' class='normal'>"&_
                    	"<INPUT type='text' name='Issue_Date' size='5' maxlength=5 "&_
                    	"onKeyPress=""return goodchars(event,'0123456789/')""> MM/YY"&_
					"<INPUT name='Issue_Date_C' type='hidden' value='Op|String|5|5|||Issue Date'>"&_
					"</td></tr>")
				sPaymentText = sPaymentText & ("<TR><TD width='176' class='normal'>Issue No.</TD>"&_
					"<TD width='219' class='normal'>"&_
                    	"<INPUT type='text' name='Issue_No' size='2' maxlength=2 "&_
                    	"onKeyPress=""return goodchars(event,'0123456789')"">"&_
					"<INPUT name='Issue_No_C' type='hidden' value='Op|String|0|2|||Issue No'>"&_
					"</td></tr>")
			end if
		end if
	end if
elseif Payment_Method="Purchase Order" then
    sPaymentText = sPaymentText & ("<tr><td>Please enter your Purchase Order #</td>"&_
        "<td align='left'><input type='text' name='Cust_PO' size='30'>"&_
        "<INPUT name='Cust_PO_C' type='hidden' value='Re|String|||||Purchase Order'></td></tr>")
End If
sPaymentText = sPaymentText & ("</TR>")
if strServerPort= 443 then
    sPaymentText = sPaymentText & ("<TR><TD colspan=2 class='normal'><B>Secure Page</B>"&_
    " all information will be encrypted before transfer.</TD></TR>")
elseif (Payment_Method = "Visa" or Payment_Method = "Mastercard" or Payment_Method = "Discover" or Payment_Method = "American Express" or Payment_Method = "Diners club" or Payment_Method="eCheck" or Payment_Method="Debit Card") then
    sPaymentText = sPaymentText & ("<TR><TD colspan=2 class='normal' bgcolor='red'>"&_
    "<font color=white><B>WARNING: This page is NOT secure</B><BR>We are unable to transfer "&_
    "you to a secure connection either because your browser does not support secure pages or "&_
    "you have disabled secure pages.  Please enable secure pages or choose a new browser "&_
    "before continuing.</font></TD></TR>")
end if
if iLocalPost=1 then
    sPaymentText = sPaymentText&("<input type=hidden name='Local_Post' value=1>")
end if
sPaymentText = sPaymentText & ("</TABLE><TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1>"&_
    "<TR><TD class='normal' colspan=3 align=center>"&_
	fn_create_action_button ("Button_image_ProcessOrder", "Payment_Process_Order", "Process Order")&_
	"</TD></TR></TABLE></form>")

sFormName = "Payment"
sSubmitName = "Payment_Process_Order"

sPaymentText = sPaymentText & ("<SCRIPT language=""JavaScript"">"&_
    "var submitcount = 0;"&_
    "function checkFields(theform) {"&_
    "if (submitcount == 0) {submitcount ++;return true;}"&_
    "else { alert(""The transaction has already been submitted please wait."&_
    "\n\nIf you are receiving this error after using the back button to return "&_
    "to the payment page, please hit F5 to refresh the screen."");return false;}}"&_
    "function goCardCode() {var w = 300;var h = 445;var winl=(screen.width - w) / 2;"&_
    "var wint =(screen.height - h) / 2;"&_
    "winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl"&_
    ";reWin = window.open('"&Switch_Name&"include/cvcard.asp','cardcode','toolbar=no,"&_
    "location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}"&_
    "var frmvalidator = new Validator("""&sFormName&""");"&_
    sJava&"</script>")

response.write sPaymentText
%>

<!--#include file="include/footer.asp"-->
