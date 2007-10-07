<!--#include file = "include/headeroutside.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page sends login Id and password to reseller's mail account.
'	Page Name:		   	forgot_password.asp
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_login.asp
'	Output Page:	    forget_password.asp
'	Date & Time:		6 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	##############
%>


<HTML><HEAD><TITLE>Reseller</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="forgot_password_files/style.css" rel=stylesheet type=text/css>
<script language="javascript" src="include/commonfunctions.js"></script>

<script language="javascript">
function getmailid()
{
var ErrMsg
ErrMsg=""
	//code for validating email field. 
if (isWhitespace(document.frmforgetpass.Store_Email.value)==true)
			{
			ErrMsg = ErrMsg + "Email address is mandatory.\n"
			}
if (isWhitespace(document.frmforgetpass.Store_Email.value)==false)
		{
		if (IsEmail(document.frmforgetpass.Store_Email.value)==false)
					{
					
					ErrMsg = ErrMsg + "Invalid Email address."
					}
		}
if (ErrMsg !="")
			{
			alert(ErrMsg);
			}
else 
			{
			document.frmforgetpass.action="get_password.asp"
			document.frmforgetpass.submit();
			}
}

</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">

<DIV id=overDiv 
style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  <TBODY>
  <TR>
    <TD class=title>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
          <TD class=title><B>Reseller</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%"></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
        <TR vAlign=top>
          <TD rowSpan=2 width=180></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Forgot Password 
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%"><FORM action=forgot_password.asp method=post 
              name="frmforgetpass"><INPUT name=Form_Name type=hidden> <INPUT name=redirect 
              type=hidden value=forgot_password.asp> 
              <TBODY>
              <TR>
                <TD height=10>&nbsp;</TD></TR>
              <TR>
                <TD colSpan=2>Please enter your email address 
               <!--   <A class=link href="get_password.asp"></A></TD></TD></TR>-->
                
              <TR>
                <TD class=inputname><B>Email</B></TD>
                <TD class=inputvalue><INPUT name="Store_Email" type="text" > </TD>
                </TR>
              <TR>
                <TD align=middle width="17%">&nbsp;</TD>
                <TD align=left vAlign=top width="84%"><BR><INPUT class=Buttons name=I1 type=button value="Get Password" onclick="javascript:getmailid()"></TD>
                <TD align=left vAlign=top width="84%">&nbsp;</TD></TR>
            
</SCRIPT>

              <TR>
                <TD class=inputname></TD></TR>
              <TR>
                <TD 
  height=20></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM></BODY></HTML>




















