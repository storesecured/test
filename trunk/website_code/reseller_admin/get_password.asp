<!--#include file = "include/EScheaderoutside.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>EasyStoreCreator</TITLE>
</HEAD>
<BODY>

<P> 
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="get_password_files/script.js" 
type=text/javascript></SCRIPT>


<META content="MSHTML 5.00.3813.800" name=GENERATOR></P>
<%

dim sqlgetmail,strgetmail,txtgetmail,rsgetmail
if trim(Request.form("Esc_Email"))<> "" then
		txtgetmail=trim(Request.Form("Esc_Email"))
		txtgetmail=fixquotes(txtgetmail)
		sqlgetmail="select fld_user_name,fld_password,fld_email_address from tbl_esc_master where fld_email_address='"&txtgetmail&"'"
		set rsgetmail=conn.execute(sqlgetmail)
		
	if not rsgetmail.eof then
		strusername=trim(checkencode(rsgetmail("fld_user_name")))
		strpassword=trim(checkencode(rsgetmail("fld_password")))
		strgetmail=trim(checkencode(rsgetmail("fld_email_address")))

		'code added here to send the mail to the reseller

		subject="Password Information"
		Txtbody="<p><FONT face=arial size=2>Dear&nbsp;"&strusername&",</FONT></p>"&_
				"<p><FONT face=arial size=2>Your user name is : "&strusername&"</FONT>"&_
				"<p><FONT face=arial size=2>Your password is  : "&strpassword&"</FONT>"
				
	'other code from reseller_billing_info for sending mail.
	
	 Set Mail = Server.CreateObject("Persits.MailSender")
		 Mail.From = "admin@easystorecreator.com "
		 Mail.AddAddress txtgetmail
		 Mail.Subject =  subject
		 Mail.Body = Txtbody
		 Mail.isHTML = True
		 Mail.Queue=True
		 Mail.Send
		Set Mail = Nothing 
	
				
		
						intflag="1"
		
		
	else
		interrorflag=1
	end if
		rsgetmail.close
		set rsgetmail=nothing
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
          <TD class=title><B>EasyStoreCreator</B></TD>
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
               <%if interrorflag="1" then %>
                <TR>
                <TD height=10></TD>
                <td><font color="red"><b>Your Email ID does not exist in our database.</font></b></td>
                </TR><TR><TD><a class=link href="ESc_login.asp">Click here to go back</a></td></tr>
              <TR>
              <% end if %>
             <%if intflag="1" then %>
              <TR>
                <TD colSpan=2>Your login information has been sent to <%= txtgetmail%></td></tr><TR><TD><a class=link href="ESc_login.asp">Click here to login</a></td></tr>
			<% end if%>
 
</SCRIPT>

              <TR>
                <TD class=inputname></TD></TR>
              <TR>
                <TD 
  height=20></TD></TR></TABLE></TD></TR></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM></P>

</BODY>
</HTML>
