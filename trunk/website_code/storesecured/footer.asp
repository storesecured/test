<script language="javascript" src="include/commonfunctions.js"></script>
<script language="javascript">

function fnlogin()
{
var ErrMsg

ErrMsg=""
//code here for validating user id 

if (isWhitespace(document.frmLogin.txtUserId.value)==true)
		{
		ErrMsg= ErrMsg + "Username is mandatory.\n";
						  
		}
if (isWhitespace(document.frmLogin.txtUserId.value)==false)
{
	if (isAlphaNumeric(document.frmLogin.txtUserId.value)==false)
		{
		ErrMsg= ErrMsg + "Special characters are not allowed for login name. \n"; 
		}
}		
if (isWhitespace(document.frmLogin.txtUserId.value)==false)
{
	if ((document.frmLogin.txtUserId.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Login name should be greater than 4 characters.\n"
		}
}		
		
if (isWhitespace(document.frmLogin.txtPassword.value)==true)
		{
		ErrMsg= ErrMsg + "Password is mandatory. \n"; 
		}

if (isWhitespace(document.frmLogin.txtPassword.value)==false)
{
	if (isAlphaNumeric(document.frmLogin.txtPassword.value)==false)
		
			{
			ErrMsg= ErrMsg + "Special characters are not allowed for password. \n"; 
			}
}			
if (isWhitespace(document.frmLogin.txtPassword.value)==false)
{
	if ((document.frmLogin.txtPassword.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Password should be greater than 4 characters.\n"
		} 
}

			
if (ErrMsg !="")
		{
		alert(ErrMsg);
		}
if (ErrMsg =="")
		{
		
		document.frmLogin.action = "http://manage.<%=site%>.com/reseller_home.asp";
		//document.frmLogin.action = "../code/Reselleradmin/reseller_home.asp";
		document.frmLogin.submit();
		}
		
		
}
</script>
<%

'code below is of no use 
'*******************************************************************
dim striddb,strpassdb,sqlgetdata,rscheck,txtid,txtpass
intFlag = "0"
if trim(request.querystring("action"))="login"  then

	'taking values entered by user for username and password in variable. 
			txtid=trim(Request.Form("txtUserId"))
			txtpass=trim(Request.Form("txtPassword"))
	
	'------------------------------------------------------------------------------------------------------------------------

set conn1 = server.CreateObject("adodb.connection")
strconn1 = "DRIVER=SQL Server;SERVER=69.20.16.229;UID=melanie;PWD=tom237;DATABASE=wizard"
'strconn1 = "DRIVER=SQL Server;SERVER=198.87.87.59;UID=user;PWD=thinkmore;DATABASE=wizard"		
conn1.open strconn1

	'query for retrieving database values
		sqlgetdata = "select fld_reseller_id,fld_user_name,fld_password from tbl_reseller_master where fld_user_name='"&txtid&"' and "&_
					"fld_password='"&txtpass&"'"
		set rscheck = conn1.execute(sqlgetdata)


		if not rscheck.eof then 

			intResellerID = trim(rscheck("fld_reseller_id"))
			session("ResellerID")  = intResellerID
			Response.redirect "http://manage."&site&".com/reseller_home.asp"
			Response.End
			
		else
			

			Response.write "Reseller_Error.asp?error=3"
			Response.End
			
		end if
		
	
end if	    
'*******************************************************************
%>
</td>
</table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center" valign="top" bgcolor="#000000" class="navigation2" width="1"><img src="images/pxl_transparent.gif" width="1" height="22"></td>
          <td align="center" valign="middle" bgcolor="#000000" class="navigation2"><a class="navigation2" href=index.html>Home</a> | <a class="navigation2" href=contactus.asp>Contact Us</a> | <a class="navigation2" href=livechat.asp>Live Chat</a></td>
        </tr>
      </table>
    </td>
    <td width="1" background="images/pxl_black.gif"><img src="images/pxl_black.gif" width="1" height="1"></td>
  </tr>
</table>

</BODY>

