<!--#include file = "include/headeroutside.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page sends login Id and password to reseller's mail account.
'	Page Name:		   	get_password.asp
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_login.asp
'	Output Page:	    get_password.asp
'	Date & Time:		10 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	##############
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>

<P> 

<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></P>
<%
'asp code here
dim sqlmail,sqlgetemail,rscheckmail,strmail,getmail,getid,getpass

if trim(Request.Form("store_email"))<>""  then
		strmail=trim(Request.Form("Store_Email"))
		strmail=fixquotes(strmail)
		
		'Response.Write "mail1="&strmail
		'code for retriving user name password from database accordig to email address 
	sqlmail="select fld_user_name,fld_password,fld_mail from tbl_reseller_master where fld_mail='"&strmail&"'" 
	
	set rscheckmail=conn.execute(sqlmail)
		if  not rscheckmail.eof then
				getmail=trim(checkencode(rscheckmail("fld_mail")))
				getid=trim(checkencode(rscheckmail("fld_user_name")))
				getpass=trim(checkencode(rscheckmail("fld_password")))
				'Response.Write"mail2"&getmail
				'Response.Write "id"&getid &"<br>"
				'Response.Write "pwd"&getpass &"<br>"
				
				
				'code added here to send the mail to the reseller
				
				subject = "Password Information"
				txtBody =  "<P><FONT face=Arial size=2>Dear&nbsp;"&getid&",</FONT></P>"&_
						   "<P><FONT face=Arial size=2>Your UserID  : "&getid&"</font>"&_
						   "<p><FONT face=Arial size=2>Your Password : "&getpass&"</FONT>"
						   
	'other code from reseller_billing_info for sending mail.
	
	 Set Mail = Server.CreateObject("Persits.MailSender")
		 Mail.From = sNoReply_email
		 Mail.AddAddress strmail
		 Mail.Subject =  subject
		 Mail.Body = txtbody
		 Mail.isHTML = True
		 Mail.Queue=True
		 Mail.Send
		Set Mail = Nothing 
	
	
	
				
				
				intFlag = "1"
			else
				interrorflag=1
				
		end if
		rscheckmail.close
		set rscheckmail=nothing
end if


%>
<DIV id=overDiv 
style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<P>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  
  <TR>
    <TD class=title>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        
        <TR>
          <TD class=title><B>Reseller</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%"></TD></TR></TABLE></TD></TR>
  <TR>
    <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        
        <TR vAlign=top>
          <TD rowSpan=2 width=180></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Forgot Password 
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%"><FORM action=get_password.asp id=FORM1 method=post 
              name=FORM1><INPUT name=Form_Name type=hidden> <INPUT name=redirect 
              type=hidden value=get_password.asp> 
              <TBODY>
              <TR>
                <TD height=10>&nbsp;</TD></TR>
              <TR>
                <!--<TD colSpan=2>Your login information has been sent to john@john.com</td></tr><TR><TD><a class=link href="reseller_login.asp">Click here to login</a></td></tr>-->
 
				<%if interrorflag="1" then %>
                <TR>
                <TD height=10></TD>
                <td><font color="red"><b>Your Email ID does not exist in our database.</font></b></td>
               <TR><TD><a class=link href="reseller_login.asp">Click here to go back</a></td></tr>
              <TR>
                </TR>
              <TR>
              <% end if %>
              
               <%if intflag="1" then %>
              <TR>
                <TD colSpan=2><b><font color="red">Your login information has been sent to <%= strmail%></font></b></td></tr><TR><TD><a class=link href="reseller_login.asp">Click here to login</a></td></tr>
			<% end if%>
 
</SCRIPT>

              <TR>
                <TD class=inputname></TD></TR>
              <TR>
                <TD 
  height=20></TD></TR></TABLE></TD></TR></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM></P>

</BODY>
</HTML>
