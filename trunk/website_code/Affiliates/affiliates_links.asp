<!--#include virtual="common/connection.asp"-->
<%

if Session("AFFILIATE_SESSION")="TRUE" then
        %><!--#include file="affiliates_session.asp"--><%
else
        'AFFILIATE IS NOT LOGED IN,
        'REDIRECT AFFILIATE TO LOGIN WINDOW
        Session("Is_Login") = False
        Response.Redirect "affiliates_login.asp"
end if

%>

<SCRIPT language="javascript">

function ShowHTMLBox(box)
{
	
	if (document.getElementById("ShowHTML"+box).style.display=="block")
	{
		document.getElementById("ShowHTML"+box).style.display="none"
		document.getElementById("ShowHide"+box).value="Show Banner HTML"
	}
	else
	{
		document.getElementById("ShowHTML"+box).style.display="block"	
		document.getElementById("ShowHide"+box).value="Hide Banner HTML"
	}
}
</SCRIPT>



<%
       sql_select = "select Store_Settings.Site_Name,Store_Settings.Store_Domain,Store_Settings.Use_Domain_Name, Store_Affiliates.Code from store_settings,store_affiliates where Store_Settings.store_id="&store_id&" and Store_Affiliates.Affiliate_ID="&Session("AFFILIATE_AFFILIATE_ID")

        rs_Store.open sql_select,conn_store,1,1
        Store_Domain=rs_Store("Store_Domain")
        Site_Name=rs_Store("Site_Name")
        Code = rs_Store("Code")

       if rs_store("Use_Domain_Name") then
          Site_name="http://"&Store_domain&"/"
       else
          Site_name="http://"&Site_name&"/"
       end if
        rs_Store.close

        %>
         <!--#include file="affiliates_header.asp"-->
<!--		 <tr><td colspan=4><font face="Arial" size="2">Your referral link is <%= Site_Name %>store.asp/CAME_FROM/<%= Code %></font></td></tr>-->
		<%
		

		sql_banners="select Banner_Image, Banner_Text, View_Order from Store_Affiliate_Banners where Store_Id="&Store_Id &" order by View_Order"
		rs_Store.open sql_banners,conn_store,1,1

		%>
<tr>
	<td colspan=4>
		<table border="0" cellspacing="0" cellpadding="0" align="center"  width="70%">
		<% 
		i=1
		while not rs_Store.EOF 	%>		

			<tr>
				<td width="100%"> &nbsp;</td>
			</tr>

			<% if rs_Store("Banner_Image") <> "" then %>
			
				<tr>
					<td align="center">
						<a href="<%= Site_Name %>store.asp/CAME_FROM/<%= Code %>"><IMG src="<%= Site_Name %>images/<%=rs_Store("Banner_Image")%>" border="0" alt="<%=rs_Store("Banner_Text")%>"></a>
					</td>
				</tr>

				<tr>
				<td colspan=4 align="center"><br><input name="showHide<%=i%>" type="button" value="Show Banner HTML" style="BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; PADDING-LEFT: 5px; FONT-WEIGHT: normal; FONT-SIZE: 10pt; BORDER-LEFT: #ffffff 1px solid; COLOR: #000000; BORDER-BOTTOM: #ffffff 1px solid; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BACKGROUND-COLOR: #dddddd" onclick="javascript:ShowHTMLBox('<%=i%>')"></td>
				</tr>
				<tr>
					<td colspan=4  align="center">
					<div name="ShowHTML<%=i%>" id="ShowHTML<%=i%>" style="display:none" >
						<textarea style="BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; PADDING-LEFT: 5px; FONT-WEIGHT: normal; FONT-SIZE: 10pt; PADDING-BOTTOM: 2px; BORDER-LEFT: #ffffff 1px solid; COLOR: #000000; PADDING-TOP: 2px; BORDER-BOTTOM: #ffffff 1px solid; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BACKGROUND-COLOR: #dddddd" cols="100" name="BannerText" rows="3"  readonly><a href="<%= Site_Name %>store.asp/CAME_FROM/<%= Code %>"><IMG src="<%= Site_Name %>images/<%=rs_Store("Banner_Image")%>" border="0" alt="<%=rs_Store("Banner_Text")%>"></a></textarea>
					</div>
					</td>
				</tr>
	
			<% else %>

				<tr>
					<td align="center">
						<a href="<%= Site_Name %>store.asp/CAME_FROM/<%= Code %>"><font face="verdana" size="1"><%=rs_Store("Banner_Text")%></font></a>
					</td>
				</tr>	


				<tr>
					<td colspan=4 align="center"><br><input name="showHide<%=i%>" type="button" value="Show Banner HTML" style="BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; PADDING-LEFT: 5px; FONT-WEIGHT: normal; FONT-SIZE: 10pt; BORDER-LEFT: #ffffff 1px solid; COLOR: #000000; BORDER-BOTTOM: #ffffff 1px solid; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BACKGROUND-COLOR: #dddddd" onclick="javascript:ShowHTMLBox('<%=i%>')"></td>
				</tr>
				<tr>
					<td colspan=4  align="center">
						<div name="ShowHTML<%=i%>" id="ShowHTML<%=i%>" style="display:none" >
						<textarea style="BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; PADDING-LEFT: 5px; FONT-WEIGHT: normal; FONT-SIZE: 10pt; PADDING-BOTTOM: 2px; BORDER-LEFT: #ffffff 1px solid; COLOR: #000000; PADDING-TOP: 2px; BORDER-BOTTOM: #ffffff 1px solid; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif; BACKGROUND-COLOR: #dddddd" cols="100" name="BannerText" rows="3"  readonly><a href="<%= Site_Name %>store.asp/CAME_FROM/<%= Code %>"><font face="verdana" size="1"><%=rs_Store("Banner_Text")%></font></a></textarea>
						</div>
					</td>
				</tr>

			<% end if %>

			<tr>
				<td width="100%"> &nbsp;</td>
			</tr>
	
			<tr>
				<td width="70%" height="1" style="BACKGROUND-COLOR: #dddddd"> 
				</td>
			</tr>
	
		<% 	
			rs_Store.movenext
			i=i+1
			wend 
		%>
		
		<tr><td colspan=4><font face="Arial" size="2">Your referral link is <%= Site_Name %>store.asp/CAME_FROM/<%= Code %></font></td></tr>


		</table>
	</td>
</tr>
