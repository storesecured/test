<!--#include file = "header.asp"-->
<html>
<head>
<script language=javascript>
function To()
{
	
var  Errmsg = "";
	if(document.all["tdmail"].innerText == "" )
	{
		Errmsg =  "To field should be there.";

	}
	if(Errmsg != "")
	{
		alert(Errmsg);
		return false;
	}
	else
	{
		document.frmContactUs.action = "contactus.asp?action=" + document.all["tdmail"].innerText;
		document.frmContactUs.submit()
	}
	
	
}



function SelectOffice()
{
//for displaying email ids accoring to office chosen.

if(document.all["ddlOffice"].value == 2 )
document.all["tdmail"].innerText = document.all["hdmail"].value  ;
else
document.all["tdmail"].innerText = sSales_email;

document.all["hdmail1"].value = document.all["tdmail"].innerText;

}
</script>
</head>
<body>
<%
title = "StoreSecured Contact Form"
description = "Merchant account provider offers merchant account processor functions. Integrates with many merchant account service provider companies. Apply for your small business merchant account today. We'll completely setup your online store for you."
keyword1="merchant account provider"
keyword2="merchant account processor"
keyword3="merchant account service provider"
keyword4="small business merchant account"
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="contactus"


'if request.form("Update") <> "" then
if(Request.QueryString("action") <> "" ) then
    fromname = Replace(Request.Form("fromname"), "'", "''")
    fromemail = Replace(Request.Form("fromemail"), "'", "''")
    message = Replace(Request.Form("message"), "'", "''")
    tomail = Request.QueryString("action")'Request.Form("hdmail1")
    
    if instr(fromemail,"@") > 0 and instr(fromemail," ") = 0 then
       Set Mail = Server.CreateObject("Persits.MailSender")
       Mail.From = fromemail
       Mail.AddAddress = tomail'"sales@"&host
       Mail.Subject = Name&" contact form from " & fromname
       Mail.Body = message
       Mail.Queue = True
       Mail.Send
       response.write "You will be contacted shortly by storesecured."
    else
       response.write "You did not enter a valid email address.  Please use your browsers back button and try again."
    end if
else
%>
<!--******************************************
-->
<%	If Trim(Request("Updateresel")) <> "" Then%>
Before contacting us check and see if your question is answered in our <a href="kb.asp">
	FAQ</a>
<%	End If
	End If

'****Code Added on 17th Jan 2005***********************
		Dim intReselId, intZip, intphone, intAltPhone  
		Dim strCompName, strAdd, strCity, strState,strFax, strCountry, stremail,strcalltoll 	
		Dim intCount

		Set rsSaveName = Server.CreateObject("Adodb.Recordset")
		
		With rsSaveName
			.Source           = "sp_retrievereseladd " & intresellerid 
			.ActiveConnection = conn
			.CursorType       = adOpenForwardOnly
			.LockType         = adLockReadOnly							
			.Open		
		End With
		If Not rsSaveName.Eof Then
			strCompName = checkencode(Trim(rsSaveName("fld_company_name")))
			strAdd = checkencode(Trim(rsSaveName("fld_address")))
			strCity = checkencode(Trim(rsSaveName("fld_city")))
			strState = checkencode(Trim(rsSaveName("fld_state")))
			intphone = Trim(rsSaveName("fld_phone"))
			intAltPhone = Trim(rsSaveName("fld_alt_phone")) 
			strFax = Trim(rsSaveName("fld_fax"))
			intZip = Trim(rsSaveName("fld_zip_code"))
			intCountry = Trim(rsSaveName("fld_country"))
			stremail = Trim(rsSaveName("fld_email"))
			strcalltoll = Trim(rsSaveName("fld_calltoll"))
		End If
		strCountryQuery = "select Country from sys_countries where Country_id ="&intCountry
		set rscountry = conn.execute(strCountryQuery)
		If Not rscountry.Eof Then
			strCountry = checkencode(Trim(rscountry("Country")))
		End If
		If Trim(Request("Updateresel")) <> "" Then
			fromreselname = Replace(Request("fromreselname"), "'", "''")
			fromreselemail = Replace(Request("fromreselemail"), "'", "''")
			messageresel = Replace(Request("messageresel"), "'", "''")
			
			if instr(fromreselemail,"@") > 0 and instr(fromreselemail," ") = 0 then
				 Set Mail = Server.CreateObject("Persits.MailSender")
				 Mail.From = fromreselemail
				 Mail.AddAddress stremail
				 Mail.Subject = Name&" contact form from " & fromreselname
				 Mail.Body = messageresel
				 Mail.Queue = True
				 Mail.Send				 
				 response.write "<BR><BR>You will be contacted shortly by "&strCompName&"."
			else
				 response.write "<br>You did not enter a valid email address.  Please use your browsers back button and try again."
			end if
		'End If
		Else
%>
<form method="post" action="Contactus.asp" name="frmContactUs" ID="Form2">
	<table>
		<tr>
			<td colspan="2"> <b>Local Office:</b> <br>           
				<b><%=strCompName%></b><br>
				<%If strCompName <> "" Then%>
				<%=strCompName%>
				<BR>
				<%End if%>
				<%If strAdd <> "" Then%>
				<%=strAdd%>
				<br>
				<%End if%>
				<%If strCity <> "" Then%>
				<%=strCity%>,
				<%End if%>
				<%If strState <> "" Then%>
				<%=strState%>,
				<%End if%>
				<%If intZip <> "" Then%>
				<%=intZip%>.
				<br>
				<%End if%>
				<%If strCountry <> "" Then%>
				<%=strCountry%>
				<br>
				<br>
				<%End if%>
				<BR>
				<BR>
				<%If stremail <> "" Then%>
				<B>Email:</B>
				<%=stremail%>
				<BR>
				<%End if%>
				<%If strcalltoll <> "" Then%>
				<B>Call toll free:</B>
				<%=strcalltoll%>
				<br>

				<%End if%>


				
				<%If intphone <> "" Then%>
				<B>Phone number:</B>
				<%=intphone%>
				<br>
				<%end if%>
				

				<%If intAltPhone <> "" Then%>
				<B>Alternate number:</B>
				<%=intAltPhone%>
				<br>
				<%End if%>
				<%If strFax <> "" Then%>
				<B>Fax:</B>
				<%=strFax %>
				<br>
				<%End if%>
			</td>
		</tr>
		<!--	<tr>
		<td>Name</td>
		<td><input type=text name=fromreselname size=20></td>
	</tr>
	<tr>
		<td>Email</td>
		<td><input type=text name=fromreselemail size=20></td>
	</tr>
	<tr>
		<td>To</td>
		<td><%=stremail%></td>
	</tr>	
	<tr>
		<td colspan=2><textarea rows="5" name="messageresel" cols="50"></textarea></td>
	</tr>
	<tr>
		<td colspan=2><input type="submit" name=Updateresel value="Send"></td>
	</tr>	-->
		<%	
'***************************
%>
	</table>
</form>
<%'End If%>

<%End If
	If strCompName = "" Then
%>

<%End If%>
<br>
<br>
<!--*****************************-->
<b>Main Office:</b><br>
StoreSecured<br>
10272 Aviary Drive<br>
San Diego, CA 92131<br>
United States<BR>
<br>
Current customers please login to the admin interface to submit a support 
request.
<BR>
<BR>
<table border=1 cellspacing=0 cellpadding=5>
<tr><td><b>Email</b></td><td><a href=mailto:<%=sSales_email %> class=link><%=sSales_email %></a></td><td>Anytime</td></tr>
<tr><td><b>Toll Free #</b></td><td>866-324-2764 (Continental US)</td><td rowspan=2>8:30 - 14:30 PST Mon-Fri<BR>Phones Closed Sat-Sun</td></tr>
<tr><td><b>Alternate #</b></td><td>619-651-8855 (International)</td></tr>
<tr><td><b>Fax #</b></td><td>888-750-6936</td><td>Anytime</td></tr>
<tr><td><b>Live Chat</b></td><td><a href=livechat.asp class=link>Click here</a></td><td>4:00 - 19:30 PST Mon-Fri<BR>8:30 - 17:00 PST Sat-Sun</td></tr>
</table>
<BR>
<form method="post" action="Contactus.asp" name="ContactUs" ID="Form1">
	<b>Send mail to:
		<SELECT id="ddlOffice" name="Select1" onchange="javascript:SelectOffice();" >
			<OPTION  value="1">Main office</OPTION>
			<OPTION selected value="2">Local office</OPTION>
		</SELECT>
		<br>
	</b>&nbsp;
	<table ID="Table1">
		<tr>
			<td width="50">Name</td>
			<td><input type="text" name="fromname" size="20" ID="Text1"></td>
		</tr>
		<tr>
			<td width="50">Email</td>
			<td width="200"><input type="text" name="fromemail" size="20" ID="Text2"></td>
		</tr>
		<tr>
			<td width="50">To</td>
			
			<td width="400" id=tdmail>
			<%=stremail%>
			</td>
					
		</tr>
		<tr>
			<td colspan="2"><textarea rows="5" name="message" cols="50" ID="Textarea1"></textarea></td>
		</tr>
		<tr>
			<td colspan="2"><input type="button" name="Update" value="Send" ID="Submit1" onclick="javascript:To();"></td>
		</tr>
	</table>
	<input type=hidden id="hdmail1" NAME="hdmail1" value =<%=stremail%> >
	   <input type=hidden id=hdmail NAME="hdmail" value = <%=stremail%> >
</form>
<!--*****************************-->
<!--#include file="footer.asp"-->
</body></html>
