<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sql_select = "SELECT * FROM Store_external WHERE Store_id="&Store_id
rs_Store.open sql_select,conn_store,1,1
rs_Store.MoveFirst

maxmind_enable = rs_Store("maxmind_enable")
maxmind_reject = rs_Store("maxmind_reject")
maxmind_license = rs_Store("maxmind_license")

if maxmind_enable then
	maxmind_enable_checked = "checked"
else
	maxmind_enable_checked = ""
end if
rs_Store.close

sFormAction = "Store_Settings.asp"
sName = "Store_maxmind_enable"
sFormName = "maxmind"
sCommonName="Fraud Control"
sTitle = "Payment Fraud Control"
sFullTitle = "General > Payments > Fraud Control"
sSubmitName = "Update_maxmind_enable"
thisRedirect = "maxmind_settings.asp"
sTopic = "maxmind"
AddPicker=1
sMenu = "general"
createHead thisRedirect

Processor_id=Real_Time_Processor

%>

  <TR bgcolor='#FFFFFF'>
  <td width="100%" colspan="5">
  FREE fraud transaction monitoring for up to 1000 transactions from <a class=link href=http://www.maxmind.com/app/ccfd_promo?promo=STORESEC1452 target=_blank>MaxMind</a>
  <BR><BR>Maxminds fraud
  detection services take your transaction data and run it through a series of tests based on the credit card number, ip location,
  customer phone, zip, city, state, country, customer email address and much more to determine the likelihood of fraud.  You can then 
  decide at what level of fraud likelihood to stop the transaction before even going to the gateway or set your reject score high to simply 
  use the fraud information to make an informed decision.  StoreSecured merchants receive 20% off retail pricing and 1000 free standard transactions per month from Maxmind.  
  Additional and/or premium transactions have 2 different levels and cost $.004 and $.015 respectively.
  <BR><BR>Maxmind recommends rejecting transactions with a fraud score of 2.5 or greater.
  <BR><BR>StoreSecured recommends initially setting your rejection score at 11 (Maxminds max score returned is 10) to approve all transactions.  This would allow you to view
  the fraud scores but not actually set a rejection point until you are familiar with which scores you are comfortable accepting and 
  which scores you feel are to risky.  Many current merchants have found that 2.5 score recommended for rejection is very restrictive so we recommend setting it higher initially
  and going from there.  Remember you can always see a high fraud score and decide to cancel the order manually.
  <BR><BR>Depending on the type of merchandise and level of risk you are willing to assume you can move the rejection value up or down based on what is best for your business.
  <BR><BR>You MUST have an account with Maxmind to enable the fraud checking service.  Maxmind offers both free and paid accounts.  The transaction fraud checking
  will be done in the background without any user interaction required.  If the fraud score returned for the transaction exceeds your rejection score the customer will 
  be presented with a message indicating that their transaction was rejected by the fraud checking service and they will not be allowed to continue.  If the fraud score is below 
  your rejection score the transaction will continue as normal.
  <BR><BR>This service is FREE for the first 1000 transactions.  There is no cost to open up an account with Maxmind and receive a license key.
  We highly recommend to use this service as a tool to help fight fraud.  After the first 1000 transactions the cost is very minimal at approximately 4 cents per transaction.
  <BR><BR>
  Click <a class=link href=http://www.maxmind.com/app/ccfd_promo?promo=STORESEC1452 target=_blank>here to signup</a> or get more details on <a class=link href=http://www.maxmind.com/app/ccfd_promo?promo=STORESEC1452 target=_blank>Maxminds Fraud Checking</a><BR>
  <BR><BR>
  <% if Processor_id = 1 or Processor_id = 2 or Processor_id = 5 or Processor_id=7 or Processor_id=9 or Processor_id =10 or Processor_id =14 or Processor_id =15 or Processor_id =21 or Processor_id =26 or Processor_id =28 or Processor_id =29 or Processor_id =31 or Processor_id =36 then %>
  Your currently selected gateway fully supports the Maxmind service.<BR><BR>
  <% else %>
  <font color=red><B>IMPORTANT</b>: Your currently selected gateway supports Maxminds services in a limited manner.  We are able to do fraud
  checking without the credit card bin number checks only as your gateway does not give us access to the needed information for those checks.  Thus the fraud checking will
  be done when the user enters the credit card screen versus after the credit card number is entered.
  If you are interested in full fraud checking support a list of supported gateways is included below.<BR>Gateways supporting full service: BluePay, Authorize.net, PlugnPay,
  Payflow Pro, PsiGate, Echo, Linkpoint, eWay, eftNet, Electronic Transfer, CyberSource, Xor, PayPal Pro, ProPay.<BR><BR></font>
  <% end if %>
  
  </td>
  </tr>
  <TR bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>License Key</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= maxmind_license %>" name="maxmind_license" size="60" maxlength=50>
		  <INPUT type="hidden"	name="maxmind_license_C" value="Re|String|0|50|||License Key">
		 <% small_help "maxmind_license" %></td>
  </td>
  </tr>
  <TR bgcolor='#FFFFFF'>
  <td width="25%" class="inputname"><B>Reject Fraud Scores >=</B></td>
  <td width="75%" class="inputvalue">
		  <input type="text" value="<%= maxmind_reject %>" name="maxmind_reject" size="60" maxlength=50>
		  <INPUT type="hidden"	name="maxmind_reject_C" value="Re|Number|0|11|||Reject Fraud Scores">
		 <% small_help "maxmind_reject" %></td>
  </td>
  </tr>


  <TR bgcolor='#FFFFFF'>
  <td width="100%" class="inputname"><B>Enable Maxmind</B>
		</td><td class="inputvalue"><input class="image" type="checkbox" <%= maxmind_enable_checked %> name="maxmind_Enable">
	  <% small_help "Enable_Maxmind" %></td>
  </tr>

	
<% createFoot thisRedirect,1 %>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("maxmind_license","req","Please enter your Maxmind License");
 frmvalidator.addValidation("maxmind_reject","req","Please enter the maximum fraud score to allow.");
 
</script>

