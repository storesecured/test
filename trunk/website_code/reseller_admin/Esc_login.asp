<!--#include file = "include/ESCheaderoutside.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page checks login ID and password for the escadmin. 
'	Page Name:		   	Esc_login.htm
'	Version Information: EasystoreCreator
'	Input Page:		    Esc_home.asp
'	Output Page:	    Esc_login.htm
'	Date & Time:		10 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
%>
<HTML><HEAD><TITLE>EasyStoreCreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" type=text/javascript></SCRIPT>
<script language="javascript" src="../include/commonfunctions.js"></script>
<script language="javascript">

function fnesclogin()
{
var ErrMsg
ErrMsg=""

//code here for validating user id 
if (isWhitespace(document.frmesclogin.Esc_User_Id.value)==true)
		{
		ErrMsg= ErrMsg + "Login name is mandatory.\n";
		}

if (isWhitespace(document.frmesclogin.Esc_User_Id.value)==false)
{
	if (isAlphaNumeric(document.frmesclogin.Esc_User_Id.value)==false)
		{
		ErrMsg= ErrMsg + "Login name can not contain special characters.\n"; 
		}
}		
if (isWhitespace(document.frmesclogin.Esc_User_Id.value)==false)
{
	if ((document.frmesclogin.Esc_User_Id.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Login name should be greater than 4 characters.\n"
		}
}
		
if (isWhitespace(document.frmesclogin.Esc_Password.value)==true)
		{
		ErrMsg= ErrMsg + "Password is mandatory. \n"; 
		}

if (isWhitespace(document.frmesclogin.Esc_Password.value)==false)
{
	if (isAlphaNumeric(document.frmesclogin.Esc_Password.value)==false)
		
			{
			ErrMsg= ErrMsg + "Password can not contain special characters. \n"; 
			}
}			
if (isWhitespace(document.frmesclogin.Esc_Password.value)==false)
{
	if ((document.frmesclogin.Esc_Password.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Password should be greater than 4 characters.\n"
		} 
}

			
if (ErrMsg !="")
		{
		alert(ErrMsg);
		}
else
		{
			document.frmesclogin.action="esc_login.asp?action=login";
			document.frmesclogin.submit();
		}


}
</script>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 
rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%



dim sqlgetdata,rsgetdata,txtuserid,txtpassword
intFlag = "0"
		if trim(request.querystring("action"))="login"  then

'taking values entered by user for username and password in variable.

		txtuserid=trim(Request.Form("Esc_User_Id"))
		txtuserid=fixquotes(txtuserid)
		txtpassword=trim(Request.Form("Esc_Password"))
		txtpassword=fixquotes(txtpassword)
		
		'query for retrieving database values

		sqlgetdata="select fld_admin_id,fld_user_name,fld_password from TBL_ESC_Master where fld_user_name='"&txtuserid&"' and "&_
					"fld_password='"&txtpassword&"'"
		set rsgetdata=conn.execute(sqlgetdata)

		if not rsgetdata.eof then
			intescID = trim(rsgetdata("fld_admin_id"))
			session("escid")=intescid
			Response.Redirect "easycreator_admin.asp"
			Response.End
		
		else
			intflag="1"
		end if
		rsgetdata.close
		set rsgetdata=nothing
end if
%>
<DIV id=overDiv 
style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  <TBODY>
  <TR>
    <TD class=title>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
          <TD class=title><B>EasyStoreCreator </B></TD>
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
          <TD class=pagetitle height=400 vAlign=top>Login 
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%"><FORM action="javascript:fnesclogin()" method=post name="frmesclogin"><INPUT name=Form_Name type=hidden> <INPUT name=redirect 
              type=hidden value=login_store.asp> 
              <TBODY>
              <TR>
                <TD height=10>&nbsp;</TD></TR>
                <%if intflag="1" then %>
              <TR>
                <TD height=10></TD>
                <td><font color="red"><b>Invalid Login name/Password</font></b></td>
                </TR>
              <TR>
              <% end if%>
              <TR>
                
              <TR>
                <TD class=inputname><B>Login</B></TD>
                <TD class=inputvalue><INPUT name=Esc_User_Id maxlength=12> 
				</TD>
                <TD align=middle class=inputvalue><B>
                <font color="red"></font>
                </TD></TR>
              <TR>
                <TD class=inputname><B>Password</B></td> 
                <TD class=inputvalue><INPUT name=Esc_Password type=password maxlength=12> 
                                          
                
                <TD align=middle class=inputvalue><B>
                <font color="red"></font>
                </TD></TR>
              <TR>
                <TD align=middle width="17%">&nbsp;</TD>
                <TD align=left vAlign=top width="84%"><BR><INPUT name=ReturnTo type=hidden > 
                <!--<INPUT class=Buttons name=I1 type=button value="Login" 	 onclick="javascript:fnesclogin()" >-->
                <INPUT class=Buttons name=I1 type=submit value="Login" >
                <INPUT class=Buttons name=Reset type=reset value="Reset" ></TD>
                <TD align=left vAlign=top width="84%">&nbsp;</TD></TR>
              <TR>
                <TD colSpan=2 height=10>If you have forgotten your login or 
                  password <A class=link href="forgot_password.asp">click here</A></TD></TR>
                   
              <TR>
            


              <TR>
                <TD class=inputname></TD></TR>
              <TR>
                <TD 
  height=20></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM></BODY></HTML>
