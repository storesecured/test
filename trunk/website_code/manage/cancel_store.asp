<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
if request.form <> "" then
  sMailMessage = "Store_Id="&Store_Id & " has requested to have their store cancelled."&vbcrlf&_
  "Storename="&Site_Name & vbcrlf&_
  "Service="&Service_Type & vbcrlf&_
  "Reason="&request.form("Reason")
  
  sql_update ="update store_settings set store_cancel='"&now()&"' where store_id="&store_id
  conn_store.Execute sql_update
  
  sql_insert = "New_Support_Request "&Store_Id&",'Cancel Store','"&request.form("Reason")&"','Customer','"&Store_Email&"'"
  conn_store.Execute sql_insert
  Send_Mail sNoReply_email,Store_Email,"Cancellation Request","Your store "&Site_Name&" is scheduled for deletion per your request.  If you made this request in error and do not wish to have your store removed please use the link below immediately to cancel the request.  If you do nothing your store will no longer be available after 24-48 hours."&chr(13)&chr(10)&chr(13)&chr(10)&"http://manage.easystorecreator.com/uncancel_store.asp"
  
  Store_Cancel = now()
  response.redirect "cancel_store_action.asp"
end if

sFormAction = "cancel_store.asp"
sFormName = "payment"
sTitle = "Cancel Site"
sFullTitle = "My Account > Cancel Site"
thisRedirect = "cancel_store.asp"
sMenu = "account"
sQuestion_Path = "import_export/my_account/cancel_store.htm"
createHead thisRedirect

%>
<script langauge="JavaScript">
function goCancelStore(theStoreId){
  if (confirm("I have read and agree to the statement above regarding this cancellation?\n\nSelect OK to Continue with cancellation request.")){
	document.forms[0].submit(); }}
</script>
<% if Store_Cancel <> "" then %>
   <TR bgcolor='#FFFFFF'><TD colspan=2>
   Your store is already scheduled for deletion.</td></tr>

<% else %>

	 <TR bgcolor='#FFFFFF'><TD colspan=2 class=instructions>
         By entering your cancellation request below you are agreeing to the following statement:
         <BR><BR>
         All cancellation requests are processed within 48 hours of their entry.  Once a cancellation request
         is received the store will be completely removed from our servers within 24-48 hours.  
         Once the store is removed it CANNOT be restored and none of the data within will be available.
         All future billings will be terminated when you initiate the cancellation.  No credit
         will be given for partially used periods.  If you are within the
         original 45 day testing period you may request a full refund.  If you think you may need any of the data
         contained within your store please ensure that information is captured BEFORE requesting cancellation.
          <BR><BR>Prices are locked for all existing customers.  Please note if you leave and re-return to our service
          you will signup at the then current rates.  Any price lock you previously held will be given up upon cancellation of service.
         </td></tr>
	 <TR bgcolor='#FFFFFF'>
		 <td width="24%" height="23" class="inputname"><B>Reason for Cancellation</B></td>
		 <td width="76%" height="23" class="inputvalue">
				 <textarea name="reason" rows=5 cols=50>Please enter your reason for cancelling here, we want your feedback as to what we can do better.</textarea></td>
		 </tr>
	 <TR bgcolor='#FFFFFF'><TD colspan=2 align=center><input class=buttons type="button" value="Cancel" onClick="JavaScript:goCancelStore();"></td></tr>
<% end if %>
<% createFoot thisRedirect,0 %>

