
<!--#include file = "include/headeroutside.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page checks login ID and password for the user for the resellers site
'	Page Name:		   	reseller_login.htm
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_login.htm
'	Date & Time:		6 Aug 2004 		
'	Created By:			Rashmi Badve
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
%>

<HTML><HEAD><TITLE>Reseller</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language="javascript" src="include/commonfunctions.js"></script>
<script language="javascript">

function fnlogin()
{
var ErrMsg

ErrMsg=""
//code here for validating user id 

if (isWhitespace(document.frmlogin.Store_User_Id.value)==true)
		{
		ErrMsg= ErrMsg + "Login name is mandatory.\n";
						  
		}
if (isWhitespace(document.frmlogin.Store_User_Id.value)==false)
{
	if (isAlphaNumeric(document.frmlogin.Store_User_Id.value)==false)
		{
		ErrMsg= ErrMsg + "Login name can not contain special characters.\n"; 
		}
}		
if (isWhitespace(document.frmlogin.Store_User_Id.value)==false)
{
	if ((document.frmlogin.Store_User_Id.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Login name should be greater than 4 characters.\n"
		}
}		
		//code for validating password field
		
if (isWhitespace(document.frmlogin.Store_Password.value)==true)
		{
		ErrMsg= ErrMsg + "Password is mandatory.\n"; 
		}

if (isWhitespace(document.frmlogin.Store_Password.value)==false)
{
	if (isAlphaNumeric(document.frmlogin.Store_Password.value)==false)
		
			{
			ErrMsg= ErrMsg + "Password can not contain special characters.\n"; 
			}
}			
if (isWhitespace(document.frmlogin.Store_Password.value)==false)
{
	if ((document.frmlogin.Store_Password.value.length) < 4 )
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
		document.frmlogin.action = "Reseller_login.asp?action=login";
		document.frmlogin.submit();
		}
}
</script>

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
'asp code here
dim striddb,strpassdb,sqlgetdata,rscheck,txtid,txtpass
intFlag = "0"
if trim(request.querystring("action"))="login"  then

	'taking values entered by user for username and password in variable. 
			txtid=trim(Request.Form("store_user_id"))
			txtpass=trim(Request.Form("store_password"))
			txtid=fixquotes(txtid)
			txtpass=fixquotes(txtpass)
	
	'query for retrieving database values
		sqlgetdata = "select fld_reseller_id,fld_user_name,fld_password from tbl_reseller_master where fld_user_name='"&txtid&"' and "&_
					"fld_password='"&txtpass&"'"
		set rscheck = conn.execute(sqlgetdata)
		if not rscheck.eof then 
			intResellerID = trim(rscheck("fld_reseller_id"))
			session("ResellerID")  = intResellerID
			
			Response.Redirect "Reselleradmin/reseller_home.asp"
			Response.End
						
		else
			intFlag ="1"
		end if
		rscheck.close
		set rscheck=nothing
		
	
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
          <TD class=pagetitle height=400 vAlign=top>Login 
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%"><FORM action="" method=post name="frmlogin"><INPUT name=Form_Name type=hidden> 
              
             
              <TBODY>
                <TD width="20%" height=10>&nbsp;</TD></TR>
              
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
                <TD class=inputvalue><INPUT name=Store_User_Id maxlength=50> 
</TD>
                </TR>
              <TR>
                <TD class=inputname><B>Password</B></td> 
                <TD class=inputvalue><INPUT name=Store_Password type=password maxlength=12> 
                </TD>
                          
                
                </TR>
              <TR>
                <TD align=middle width="17%">&nbsp;</TD>
                <TD align=left vAlign=top width="84%"><BR><INPUT name=ReturnTo type=hidden > 
                <INPUT class=Buttons name=I1 type=button value="Login" 	 onclick="javascript:fnlogin()" >
                
                <INPUT class=Buttons name=Reset type=Reset value="Reset" ></TD>
                <TD align=left vAlign=top width="84%">&nbsp;</TD></TR>
              <TR>
                <TD colSpan=2 height=10>If you have forgotten your login or 
                  password <A class=link href="forgot_password.asp">click here</A></TD></TR>
                   
              <TR>
                <SCRIPT language=JavaScript>
var submitcount = 0;
function submitForm()
{

   return true;

}
</SCRIPT>

              <TR>
                <TD class=inputname></TD></TR>
              <TR>
                <TD 
  height=20></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM></BODY></HTML>
