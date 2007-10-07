<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

' CHECK IF IN EDIT MODE, IF YES RETRIEVE CURRENT VALUES
' FROM THE DATABASE

sTitle ="Additional Domain Name"
sFormName="Additional_Domain"
sFormAction = "store_settings.asp"
thisRedirect = "domain.asp"
sMenu = "general"
createHead thisRedirect

%>


					<tr bgcolor='#FFFFFF'><td><b>Additional Domain Name</b></td>
						<td class="inputvalue">
							http://<input type="text" name="Store_Addl_Domain" value="<%= Store_Domain2 %>" size="35" maxlength=63 onKeyPress="return goodchars(event,'0123456789abcdefghijklmnopqrstuvwxyz-.')">
							<input type="hidden" name="Store_Addl_Domain_C" value="Op|String|0|63|.| ,%,^,&,(,),#,@,:,;,',>,<,/,=,+,{,},[,],\,~,`,$,*,_|Additional Domain Name">		</tr>
							
					

<% createFoot thisRedirect, 1%>
