<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<html><head><style>td {COLOR: 000000;TEXT-DECORATION: none;font-size : 10pt;font-family :Verdana;}
td.to {COLOR: 000000;TEXT-DECORATION: none;font-size : 10pt;font-family :Verdana;}</style></head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<%


Print_Invoices = request.querystring("Orders")
Oids = Print_Invoices
Oid = replace(Print_Invoices,",","*")

if Print_Invoices = "" then
   fn_error "You must select at least one order"
else

on error resume next
		 %>
		<tr>
				<td width="100%" colspan="3" height="74">
					<table width="600" border="0" cellpadding="2" cellspacing="0">
						<tr valign=top>
						<td colspan=2>
								<hr width="600" align="left">
								Pick List <B>Order # :</B><% =Oids %>
								
								<% text_color="" %>
								<% text_size="2" %>
								<% text_face="Arial" %>
								<% If Request.QueryString("Cart_Type") = "Invoice" then %>	  
									<% Call Create_PickList (Request.QueryString("Cart_Type"),Oid) %>
								<% Else	%>
									<% 
									
									Call Create_PickList ("Big",Oid) 
									%>
								<% End If %>
				
							</td>
						</tr>
					</table>
				</td>
			</tr>

		<%
			response.write "<tr><td width='100%' colspan='3' height='74'><div id='pgBreak' style='page-break-after:always;'><img src=images/spacer.gif width=1 height=2></div></td></tr>"

end if			' querystring condn


%>
