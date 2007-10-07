
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<META content="text/html; charset=unicode" http-equiv=Content-Type>
<%
title = "Free Website Builder eCommerce Merchant Account Free Online Store Builder"
description = "Free website builder allows easy ecommerce merchant account integration. Free online store builder expedites increased sales. Trial free website builder today."
keyword1="free website builder"
keyword2="ecommerce merchant account"
keyword3="free online store builder"
keyword4=""
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="resellersignup"
includejs=1

'REPLACE 'SINGLE QUOTES
function fixquotes(str)
fixquotes = replace(str&"","'","''")
end function
%>

<!--#include file="header.asp"-->
<META content="MSHTML 5.00.3819.300" name=GENERATOR></HEAD>
<BODY>&nbsp;<!--#INCLUDE FILE="include/clsWhois.asp"-->

<script language="javascript">
function fnCreate()
{
	
	var ErrMsg
	ErrMsg=""
	
	
	if(isWhitespace(document.frmsign.Login.value)==true)
	{
		ErrMsg = ErrMsg + "Login is mandatory.\n";
	}
	if (isWhitespace(document.frmsign.Login.value)==false)
	{
		if ((document.frmsign.Login.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Login name should be greater than 4 characters.\n"
		}
	}
	
	if(isWhitespace(document.frmsign.Password.value)==true)
	{
		ErrMsg = ErrMsg + "Password is mandatory.\n";
	}
	
	if (isWhitespace(document.frmsign.Password.value)==false)
	{
		if ((document.frmsign.Password.value.length) < 4 )
		{
		ErrMsg = ErrMsg + "Password should be greater than 4 characters.\n"
		} 
	}
	
	if(isWhitespace(document.frmsign.password_confirm.value)==true)
	{
		ErrMsg = ErrMsg + "Password Confirmation is mandatory.\n";
	}
	
	if(document.frmsign.Password.value != document.frmsign.password_confirm.value)
	{
		ErrMsg = ErrMsg + "Password Confirmation must be same as password.\n";
	}
	
	if(isWhitespace(document.frmsign.Email.value)==true)
	{
		ErrMsg = ErrMsg + "Email is mandatory.\n";
	}
	
	if (isWhitespace(document.frmsign.Email.value)== false)
		{
			if(IsEmail(document.frmsign.Email.value)==false)
			{
				ErrMsg = ErrMsg + "Invalid Email id. \n" ;		
			}	
		}
	
	if(isWhitespace(document.frmsign.Site_Host.value)==true)
	{
		ErrMsg = ErrMsg + "Site Host is mandatory.\n";
	}
	if(isWhitespace(document.frmsign.Site_Host.value)==false)
	{
		if(isAlphaNumeric(document.frmsign.Site_Host.value)==false)
		{
				ErrMsg = ErrMsg + "Site Name should not contain special characters.\n" ;		
		}
	}
	
	if (ErrMsg!="")
		{
		alert(ErrMsg);
	}
	else
	{document.frmsign.action="resellersignup.asp?signupaction=save";
	document.frmsign.submit();
	}
}
</script>
<%
if trim(Request.QueryString("signupaction"))="save" then
	
	strlogin = trim(Request.Form("Login"))
	sitehost = trim(Request.Form("Site_Host"))
	
	StrDomain = trim(Request.Form("Site_Host"))
	if sitehost = "secure" or sitehost = "manage" or sitehost = "forum" or sitehost = "www" then
		response.redirect "error.asp?Message_id=104&Message_Add=Site already exists."
	end if
	
	if inStr(sitehost,",") > 0 then
		response.redirect "error.asp?Message_id=101&Message_Add=Site Name cannot contain a comma"
	end if
	
	StrDomain = LCase(Replace(StrDomain, " ", ""))
	Subject = "Transfer"
	'if Subject = "1" then
	'	Subject = "Transfer"
	'
	'else
	'	Subject = "Register"
	'
	'end if
	StrDomain = StrDomain&Request.Form("domain_ext")

	
	dim iDomainAvailable
	If Not StrDomain = "" and 1=0 Then
			DomainName = Lcase(Trim(StrDomain))
	        DomainName = Replace(DomainName,"http://","",vbTextCompare)
	        DomainName = Replace(DomainName,"www.","",vbTextCompare)
	        Set whoisdll = Server.CreateObject("WhoisDLL.Whois") 
	        whoisdll.WhoisServer = "whois.internic.net"
	        Result = ""
	        Result = whoisdll.whois(Trim(DomainName))
	        response.write "Result="&Result
	        	response.end
	        StartPos = InStr(Result,"Whois Server:")
	        If StartPos > 0 Then
	        	StartPos = StartPos + Len("Whois Server:")
	        	EndPos = InStr(StartPos,Result,vblf)
	        	WhoisServer = Trim(Mid(Result,StartPos,EndPos-StartPos))
	        	Set whoisdll = Server.CreateObject("WhoisDLL.Whois") 
	        	whoisdll.WhoisServer = WhoisServer
	        	Result = ""
	        	Result = whoisdll.whois(Trim(DomainName))

	        	WhoisResult = Result
	        	iDomainAvailable = "0"
	        	'domain is not avalilable for registeration
	        Else
				'that means the domain is available and it can be registered
				'Response.Write "out"
				If InStr(1,Result,"No match for",vbTextCompare) > 0 Then
					 iDomainAvailable = 1
				end if
	        End If	

	        if iDomainAvailable<>"" then
				if trim(Subject) = "Register" and iDomainAvailable = "0" then
					Response.Redirect "reseller_error.asp?error=1"
				elseif trim(Subject) = "Transfer" and iDomainAvailable = "1" then
					Response.Redirect "reseller_error.asp?error=2"
				end if
			end if	
	End If
	
	'Code here to check if login name & site name  already there or not.
	strGetlogin	= "select fld_user_name,fld_website from tbl_reseller_master"

	set rsGetlogin=conn.execute(strGetlogin)
	if not rsGetlogin.eof then
		do while not rsGetlogin.eof 
				loginname = trim(rsGetlogin("fld_user_name"))
				website = trim(rsGetlogin("fld_website"))
							
			if trim(lcase(strlogin))=trim(lcase(loginname)) then
				intflag = "1"
				exit do
			end if

			if trim(lcase(sitehost))=trim(lcase(website)) then
				intflag1 = "1"
				exit do
			end if	
			
			if not rsGetlogin.eof then
				rsGetlogin.movenext
			end if	
		loop
	end if 		
	
	if rsGetlogin.eof then
				'Retrive the values from text boxes
				currentdate = date()
				strlogin = trim(fixquotes(Request.Form("Login")))
				strpassword = trim(fixquotes(Request.Form ("Password")))
				strconfirm = trim(fixquotes(Request.Form("password_confirm")))
				stremail = trim(Request.Form("Email"))
				strhost = trim(Request.Form("Site_Host"))
			
				'Code here to insert the information of reseller into the database
				
				strinsert = "Put_reseller_info '','','','','','','','','','"&strlogin&"','"&strpassword&"','"&stremail&"','"&StrDomain&"','"&currentdate&"','"&Subject&"'"
				set rscheck = conn.execute(strinsert)
				if not rscheck.eof then 
					'set rsNew = conn.execute("select max(fld_reseller_id) as maxid from tbltemp")
					set rsNew =  conn.execute("SELECT IDENT_CURRENT('tbltemp')")
						if not rsNew.eof then
							'Response.Write "df"&trim(rsNew(0))
							resellerid = trim(rsNew(0))
						end if
					Session("tempresellerid") = resellerid

						
						
				%>
				<script language="javascript">
					//document.location.href = "reseller_payment_mode.asp"
					document.location.href = "reseller_payment.asp"
				</script>
				<%end if
	end if			
end if%>
	
<table border="0" width=525 align=left bordercolor="#000066">

			<tr>
					<td valign="top" align=left >
							<b>Reseller Program</b><br>

							</td>
				
				</tr>

			<form method="post" action="resellersignup.asp?signupaction=save" name="frmsign">
			<tr><td colspan=2 align=left>New Reseller Signup</td></tr>
			
			<% if intflag="1" then%>
			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Login</font></td>
					<td width="98%">
					<font face="Arial" color="red" size="2">Login Already Exists.</font>
					</td>
				</tr>
			<% end if%>
			
			
			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Login</font></td>
					<td width="98%">
					<input name="Login" maxlength=20 size=20 >
					 <input type="hidden" name="Login_C" value="Re|String|0|20|||Login"></td>
				</tr>

			<tr>
			
					<td width="1%" nowrap><font face="Arial" size="2">Password</font></td>
					<td width="98%">
					<input type="password" name="Password"  maxlength=20 size=20 >
					<input type="hidden" name="Password_C" value="Re|String|0|20|||Password"></td>
				</tr>

			<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Password Confirmation</font></td>
					<td width="98%">
					<input type="password" name="password_confirm" maxlength=20  >
					<input type="hidden" name="Password_confirm_C" value="Re|String|0|20|||Password Confirmation"></td>
				</tr>
			 <tr>

					<td width="1%" nowrap><font face="Arial" size="2">Email</font></td>
					<td width="98%">
					<input name="Email" maxlength=50 >
					<input type="hidden" name="Email_C" value="Re|Email|0|50|||Email">
			</td>
				</tr>
				
				<% if intflag1="1" or intflag1="3"  then%>
				<tr>

					<td width="1%" nowrap><font face="Arial" size="2">Site Host</font></td>
					<td width="98%">
					<font face="Arial" color="red" size="2">Site name already exists.</font></td>
				<% end if %>
				</tr>
			
			
			<tr>
					<td width="1%" nowrap><font face="Arial" size="2">Site Host</font></td>
					<td width="98%">
					http://www.<input name="Site_Host" maxlength=50 >
					<input type="hidden" name="Site_Host_C" value="Re|String|0|50|||Site Host">
					<select name="domain_ext">
					<option value=".com">.com</option>
					<option value=".net">.net</option>
					<option value=".org">.org</option>
					<option value=".biz">.biz</option>
					<option value=".us">.us</option>
					<option value=".info">.info</option>
					</select>&nbsp;[e.g. yahoo]</td>
			</tr>		
			<tr>
					<td colspan=2 >Type in the domain that you would like to transfer or register. If you do not enter an extension, ie .com, .net, .org the system will assume .com.<br> Site name should not contain spaces, commas, or any special characters. </td>
			
			</tr>
				
			<tr>

					<td colspan=2 align=left><br>
						<input type="button" border="0" value="Create" onclick="javascript:fnCreate()"> </td>
				</tr>
						
			</form><!--<SCRIPT language="JavaScript">

var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("First_name","req","Please enter a first name.");
 frmvalidator.addValidation("Last_name","req","Please enter a last name.");
 frmvalidator.addValidation("Address","req","Please enter a address.");
 frmvalidator.addValidation("City","req","Please enter a city.");
 frmvalidator.addValidation("State","req","Please enter a state.");
 frmvalidator.addValidation("Zip","req","Please enter a zip.");
 frmvalidator.addValidation("Country","req","Please enter a country.");
 frmvalidator.addValidation("Phone","req","Please enter a phone.");
 frmvalidator.addValidation("Login","req","Please enter a login.");
 frmvalidator.addValidation("Password","req","Please enter a a password.");
 frmvalidator.addValidation("password_confirm","req","Please re-enter your password.");
 frmvalidator.addValidation("Email","req","Please enter a valid email.");
 frmvalidator.addValidation("Email","email","Please enter a valid email.");
 //frmvalidator.addValidation("Website","req","Please enter your website.");
 frmvalidator.addValidation("Site_Host","req","Please enter your site host.");
 //frmvalidator.addValidation("Site_Name","req","Please enter your site name.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");

function DoCustomValidation()
{
  var frm = document.forms[0];
    if (document.forms[0].Password.value == document.forms[0].password_confirm.value)
    {
      return true;	
    }
    else
    {
      alert('The 2 passwords must match.');
      return false;
    }
  }
  

</script>-->

		</table><!--#include file="footer.asp"--></BOdy>
