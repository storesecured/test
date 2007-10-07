<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%



sFormAction = "cybersource_key_action.asp"
sName = "cybersource_key_form"
sTitle = "Upload Cybersource Security Key"
sSubmitName = "Upload_Key"
thisRedirect = "cybersource_key.asp"
sEncType = "ENCTYPE='multipart/form-data'"
sMenu="general"
createHead thisRedirect

if SizeUsage > 100 then%>
	<TR bgcolor='#FFFFFF'><td><font color=red>You are not allowed to upload anything because you are currently over your size limit.
	<BR><BR>
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</font></td></tr>
<%
else
%>

			<TR bgcolor='#FFFFFF'>
				<TD colspan="3">

					Upload your security key file for Cybersource here. If you do not have one please follow the procedure below<br><br>
					
					Getting a New Security Key from Cybersource<br>
					1. To create a security key, log into the Enterprise Business Center: https://ebc.cybersource.com/<br>
					2 Click the Account Mgmt tab.<br>
					3 Click Transaction Security Keys in the navigation pane. <br>
					4. Click Generate Key. A warning message appears.<br>
					5 To accept the security warning, click Yes. (Otherwise, click No or More Info.)
					   The download may take several minutes during which the applet may appear as a large gray box. Another warning message appears.<br>
					6 Verify that the certificate is signed by “CyberSource Corporation”, then click Yes. Another warning message appears.<br>
					7 To accept the certificate, click Yes or Always. (Otherwise, click No or More Details) 
						The New Security Key window appears.When the Generate Certificate Request button appears, the download is complete.<br>
					8 Click Generate Certificate Request.Messages appear in the Messages box while a new key is generated. Another warning message appears.<br>
					9 To accept the certificate, click Yes or Always. (Otherwise, click No or More Details.) 
						Messages continue to appear in the box. Your browser opens its Save As dialog box.<br>
					10 Choose a safe location for your key (<username>.p12). <br>
					11 To verify that your key is active, click Transaction Security Keys in the navigation pane. Your new key is listed at the bottom of the table.<br>
						You have finished generating your key.<br><br>

				</TD>
			</TR>
				<TR bgcolor='#FFFFFF'><TD class="inputname" width="14%"><B>Key File</B></td>
				<td class="inputvalue" align="center" width="80%"><input type="file" name="Key1" size="35"></td>
					<td class="inputvalue" align="center" ><B><a class="help" onmouseover="return overlib('Path of the Cybersource key file to be uploaded from your computer to our server');" onmouseout="return nd();">(?)</a></td>
				</TR>
				<TR bgcolor='#FFFFFF'><TD width="100%" colspan="3" align="center" ><input type="submit" class="Buttons" value="Upload" name="Upload_Key"></td>
				</tr>
<%
end if

createFoot thisRedirect, 2%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);

</script>
