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
		
		document.frmLogin.action = "http://managereseller.<%=site%>.com/reseller_home.asp";
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
			Response.redirect "http://managereseller."&site&".com/reseller_home.asp"
			Response.End
			
		else
			

			Response.write "Reseller_Error.asp?error=3"
			Response.End
			
		end if
		
	
end if	    
'*******************************************************************
%>
</td></tr></table>
</td>
          <td width="173" align="center" valign="top" bgcolor="#FFFFFF">
            <table width="173" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="center" valign="top"><a href="https://www.<%=site%>.com/resellersignup.asp"><img src="images/btn_signup.gif" border="0" width="173" height="51"></a></td>
              </tr>
              <tr>
                <td align="center" valign="top">
                  <table width="160" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td class="Arial10Black">Signup takes only a few minutes and your 
                      rebranded website builder can be up and running.</td>
                    </tr>
                  </table>
                  <br>
                </td>
              </tr>
              <tr>
                <td align="center" valign="top"><a href="http://managereseller.<%=site%>.com"><img src="images/btn_login_account.gif" border="0" width="173" height="49"></a></td>
              </tr>
              <tr>
                <td align="center" valign="top">
                  <form name="frmLogin" method="post" action="javascript:fnlogin()">
                    <table width="165" border="0" cellspacing="0" cellpadding="0">
                        <%if intflag="1" then %>
              <TR>
                <td width="115"></TD>
                <td align="center" valign="middle" class="Arial10Black"><b>Invalid Userid/Password</font></b></td>
                </TR>
              <TR>
              <% end if%>
                      <tr>
                        <td width="115"><input type="text" name="txtUserId" class="inputboxes" value="User Name" size="20"></td>
                        <td align="center" valign="middle" class="Arial10Black"><a href="http://managereseller.<%=site%>.com/forgot_password.asp" class="Arial10Black"><strong>FORGOT
                          LOGIN?</strong></a></td>
                      </tr>
                      <tr>
                        <td width="115"><input type="password" name="txtPassword" class="inputboxes" value="Password" size="20"></td>
                        <td><input name="imageField" type="image" src="images/btn_login.gif" border="0" width="50" height="21"></td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr>
                <td align="center" valign="top"><!-- BEGIN LivePerson Button Code --><table border='0' cellspacing='2' cellpadding='2'><tr><td align="center"></td><td align='center'><a href='https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&byhref=1' target='chat7400929'  onClick="javascript:window.open('https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&referrer='+escape(document.location),'chat7400929','width=472,height=320');return false;" ><img src='https://server.iad.liveperson.net/hc/7400929/?cmd=repstate&site=7400929&&ver=1&category=en;woman;5' name='hcIcon' width=165 height=60 border=0></a></td></tr></table><!-- END LivePerson Button code -->
      </td>
              </tr>
              <tr>
                <td align="center" valign="top"></td>
              </tr>
              <tr>
                <td align="center" valign="top"></td>
              </tr>
              <tr>
                <td align="center" valign="top"></td>
              </tr>
              <tr>
                <td align="center" valign="top"></td>
              </tr>
              <tr>
                <td align="center" valign="top"></td>
              </tr>
              <tr>
                <td align="center" valign="top"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td align="center" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
          <td align="center" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center" valign="top" bgcolor="#000000" class="navigation2" width="1"><img src="images/pxl_transparent.gif" width="1" height="22"></td>
          <td align="center" valign="middle" bgcolor="#000000" class="navigation2"><a href="overview.asp" class="navigation2">Overview</a>
            | <a href="whatsincluded.asp" class="navigation2">Whats Included</a> | <a href="earnings.asp" class="navigation2">How
            much can I earn</a> | <a href="sample-sites.asp" class="navigation2">Sample Sites</a>
            | <a href="faq.asp" class="navigation2">FAQ</a> | <a href="http://managereseller.<%=site%>.com" class="navigation2">Login</a>
            | <a href="https://www.<%=site%>.com/resellersignup.asp" class="navigation2">Signup</a>
            | <a href="contactus.asp" class="navigation2">Contact Us</a>
            | <a href="billinginfo.asp" class="navigation2">Billing Info</a></td>
        </tr>
      </table>
    </td>
    <td width="1" background="images/pxl_black.gif"><img src="images/pxl_black.gif" width="1" height="1"></td>
  </tr>
</table>

</BODY>
