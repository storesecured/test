<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sFlashHelp="outlook/outlook.htm"
sMediaHelp="outlook/outlook.wmv"
sZipHelp="outlook/outlook.zip"

sTitle = "Email Setup Instructions"
sFullTitle = "General > Email > Setup Instructions"
thisRedirect = "message_list.asp"
sTopic="Message_List"
sMenu="general"
addPicker=1
sQuestion_Path = "general/email.htm"
createHead thisRedirect

if Service_Type < 1  then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		PEARL Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% elseif Trial_Version then %>
        <tr bgcolor='#FFFFFF'>
	<td colspan=2>
		Email is not available for trial users.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
<% else
sql_select = "SELECT Store_User_Id,Store_Password FROM Store_Settings WHERE Store_id="&Store_id
rs_store.open sql_select,conn_store,1,1
Username=rs_store("Store_User_Id")
Password=rs_Store("Store_Password")
rs_store.close
%>

   <% if inStr(Store_Domain,root) > 0 or Store_Domain = "" then %>
		<tr bgcolor='#FFFFFF'><TD>Your free email account will be activated after you choose a <a href=domain.asp class=link>domain name</a></td></tr>
	  <% else %>

		
			<TR bgcolor=FFFFFF><TD>
		</form><form action=http://mail2.storesecured.com/login.aspx target=_blank method=post>
                <input type=hidden name=shortcutlink value=autologin id=shortcutlink>
                <input type=hidden name=email id=email value='managedomain@<%= replace(lcase(Store_Domain),"www.","") %>'>
                <input type=hidden name=password id=password value='<%= Password %>'>
                <input type=submit name="Login" value="Login to Manage Email">
                </form></td><td>
                <form action=http://mail2.storesecured.com/ target=_blank method=post>
                <input type=hidden name=shortcutlink value=autologin id=shortcutlink>
                <input type=submit name="Login" value="Login to Read Email">
                </form>
                </td></tr>
                <tr bgcolor=FFFFFF><TD colspan=2>Please note that email accounts for new domain names are added
                every 4 hours.  If it has been less then 4 hours since you entered your domain name for a transfer you will
                be unable to login to manage your mail.  Please wait at least 4 hours before trying to do so.</td></tr>
			
		<tr bgcolor='#FFFFFF'><TD colspan=2>
  <% Email_Domain = Replace(Store_Domain,"www.","")
			Email_Domain = Replace(Email_Domain,"WWW.","")
			%>
		 <B>Email Configuration</b><BR>
		 Incoming Mail Server = <%= mail_server %><BR>
		 *Outgoing Mail Server = <%= mail_server %><BR>
		 Account Name/Login = youremail@<%= Email_Domain%><BR>
		 Password = whatever you indicated on setup<BR>
		 Email Address = youremail@<%= Email_Domain %> <I>(ie sales@<%= Email_Domain %>, info@<%= Email_Domain %>, yourname@<%= Email_Domain %>)</i><BR>
		 Outgoing mailserver <b>DOES</b> require authentication make sure this box is checked<BR>
		 Logon using Secure Password Authentication should <B>NOT</b> be checked<BR>
     <BR>

		 <BR>
		 <table border=1 cellspacing=0 cellpadding=5 bgcolor=yellow><tr bgcolor='#FFFFFF'><td>
          <B>*Note: Some ISP's block the outgoing mail server port for any 
          use except their own mailservers.  If you are having trouble sending
          out mail please contact your ISP and ask for an alternate outgoing 
          mail server or modify your outgoing mail port for our servers to use the secondary port which is 587.</td>
				</td></tr></table>
				<BR>
		 <B>Spam Policy</B><BR>Any stores who are sending unsoliticted, misrepresentative email, or abusing our email system, will have their
		 account immediately removed with absolutely no refunds.  Email is very important to our customers stores and we will not tolerate
		 anyone who deliberately misuses our service.  We reserve the right to monitor any stores email
		 usage if a complaint has been filed regarding possible spam.

		 <BR><BR><B>How do I check my email?</B><BR>
		 To access your email we recommend using Outlook Express which is a free pop3 email program.<BR>
		<a class="link" href="http://www.microsoft.com/windows/ie/default.asp" target="_blank">Outlook Express</a> - Microsofts free email program, comes with internet explorer<BR>
		 You can use any other SMTP based email program or the built in webmail system as well. <BR><BR>
       <B>Can I have my email forwarded?</B><BR>
		 <% URL = left(Site_Name,len(Site_Name)-1) %>
		 Yes, to forward your mail or configure other mail settings please login to your webmail account.
	  <BR><BR><B>Junk Mail</b><BR>
          There is built in spam filtering which will try to automatically detect and sort out your junk mail.  
          We recommend periodically checking the junk mail folder and/or modifying your spam settings accordingly to ensure you are not receiving false positives.
          <BR><BR><B>Offsite Mail</b><BR>
          If you do not want to use StoreSecureds mailservice and prefer instead to host your mail somewhere else.  Please post a 
          support request with the new mailserver information with the Subject, Offsite Mailserver.  Normal response time for this type of change is 1 business day.</td></tr>

          <% end if %>
          
<% end if %>
<% createFoot thisRedirect, 0 %>


