<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sql_select = "SELECT * FROM Store_external WHERE Store_id="&Store_id
rs_Store.open sql_select,conn_store,1,1
rs_Store.MoveFirst

preCharge_Enable = rs_Store("Precharge_Enable")
preCharge_MerchantID = rs_Store("Precharge_MerchantID")
preCharge_Security1 = rs_Store("Precharge_Security1")
preCharge_Security2 = rs_Store("Precharge_Security2")
if preCharge_Enable then
	preCharge_checked = "checked"
else
	preCharge_checked = ""
end if
rs_Store.close

sFormAction = "Store_Settings.asp"
sName = "Store_preCharge"
sFormName = "preCharge"
sTitle = "Guaranteed Transactions"
sSubmitName = "Update_preCharge"
thisRedirect = "preCharge_settings.asp"
sTopic = "preCharge"
AddPicker=1
sMenu = "general"
createHead thisRedirect

%>

  <TR bgcolor='#FFFFFF'>
  <td width="100%" colspan="5">Save time and money with guaranteed protection
using preCharge Certified Payments. Replace your
current time-consuming methods of online order
screening with preCharge Certified Payments, and
never let fraud hinder your business again.
preCharge offers merchants a proven method of
preventing fraud which comes complete with
guaranteed protection. No order screening process
can prevent 100% of fraud, however, while using
preCharge, all credit card transactions are guaranteed
against chargebacks. Unlike any other solution,
preCharge stands behind its technology with a full
guarantee, meaning your company will be reimbursed
when a chargeback does occur.
As any seasoned merchant knows, accepting credit
cards online is a double-edged sword.
<BR><BR>
With preCharge Certified Payments, your business no
longer has the stress and uncertainty accompanied
with card-not-present transactions. Peace of mind is
only a few steps away.
<BR><BR>
<B>How does it work?</b><BR>
The preCharge system looks at each transaction,
evaluating over 100 points of any given transaction—
at the moment a consumer submits their information.
Completely transparent to the consumer. Authorization
and screening occurs on the back-end, so the user
experience is uninterrupted.
<BR><BR>
Every transaction is processed using SSL 3.0 Strong
Encryption. All data is encrypted using a patentpending
reverse MD5 256bit encryption scheme which
cannot be broken. At no time is any information
communicated or used for outside purposes.
<BR><BR>
<B>The process is as follows:</b><BR>
1. The consumer submits their billing information to
your system.<BR>
2. The billing information is encrypted and sent to the
preCharge server for review without user intervention.<BR>
3. The preCharge system returns a fraud score and
response code to your system in less than 3 seconds.<BR>
4. The merchant system recognizes the fraud score,
and based on your settings, accepts or declines the
transaction. <BR>
5. You rest easy, assured that any payment that gets
approved is guaranteed to you.<BR><BR>
<B>Features</b><BR>
• Combines over 100 different cross-checks including IP,
geo-targeting, credit card number, bin number,
telephone and zip code validation for each transaction
all calculated and returned in a simple fraud score.<BR>
• Utilizes multiple activity databases updated in real-time
and throughout the day for maximum protection.<BR>
• Performs both negative and positive file checks. <BR>
• Authorizes transactions in 3 seconds or less.<BR>
• Includes comprehensive free email provider checking
and real-time email validation.<BR>
• Provides frequency checks to identify suspicious activity
across all preCharge merchants.<BR>
• Validates transactions in over 230 countries around the
globe each of which can be blocked.<BR>
• Offers complete consumer transparency - no user
confirmation, no user contact, and no user registration.<BR><BR>
<B>Benefits</b><BR><BR>
• Eliminates financial loss due to bank service fees and
chargeback losses.<BR>
• Compatible with any system, any gateway and any
merchant account. <BR>
• Saves processing time by automating the scrubbing
process.<BR>
• Ensures all transactions approved will be paid.<BR>
• Fully configurable to suit all of your business needs.<BR>
<BR>
<B>Cost</b><BR>
Each transaction is guaranteed at a rate of 2%-5.5% per transaction.  The actual rate depends on
the fraud score assigned to the transaction.<BR>
2.0% for score 0-250<BR>
3.5% for score 250-500<BR>
5.5% for score 500-750<BR>
Transactions scored > 750 cannot be guaranteed and will be rejected before reaching the gateway.

  </td>
  </tr>
  <TR bgcolor='#FFFFFF'>
  <td width="100%" colspan="5"><BR><BR>
  <a href=precharge.doc class=link target=new>Click here</A> to download the application and 
  get started immediately with your own guaranteed payment service.
  <BR><BR>
  <font size=1><B>Important Note:</b> Precharge is currently setup for use with the following gateways: Authorize.net, PlugnPay,
  Payflow Pro, Linkpoint, CyberSource.<BR>
  If your gateway is not in the list and you wish to use Precharge please contact support for further 
  information on limitations and use.</font>
  <BR><BR>
  </td>
  </tr>
  <TR bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Merchant ID</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= preCharge_MerchantID %>" name="preCharge_MerchantID" size="25" maxlength=50>
		  <INPUT type="hidden"	name="preCharge_MerchantID_C" value="Re|String|0|50|||Merchant ID">
		 <% small_help "MerchantID" %></td>
  </td>
  </tr>
  <TR bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Security Code1</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= preCharge_Security1 %>" name="preCharge_Security1" size="25" maxlength=50>
		  <INPUT type="hidden"	name="preCharge_Seccurity1_C" value="Re|String|0|50|||Security Code 1">
		 <% small_help "Secyrity1" %></td>
  </td>
  </tr>

  <TR bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Security Code2</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= preCharge_Security2 %>" name="preCharge_Security2" size="25" maxlength=50>
		  <INPUT type="hidden"	name="preCharge_Security2_C" value="Re|String|0|50|||Security Code">
		 <% small_help "Security2" %></td>
  </td>
  </tr>

  <TR bgcolor='#FFFFFF'>
  <td width="100%" class="inputname"><B>Enable Certification</B>
		</td><td class="inputvalue"><input class="image" type="checkbox" <%= preCharge_checked %> name="preCharge_Enable">
	  <% small_help "Enable Precharge" %></td>
  </tr>

	
<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("preCharge_MerchantID","req","Please enter your preCharge MerchantID.");
 frmvalidator.addValidation("preCharge_Security1","req","Please Security Code 1 supplied by preCharge.");
 frmvalidator.addValidation("preCharge_Security2","req","Please Security Code 2 supplied by preCharge.");

</script>

