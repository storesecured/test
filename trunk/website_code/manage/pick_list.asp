<!--#include file="Global_Settings.asp"-->
<!--#include virtual="common/crypt.asp"-->
<!--#include file="pagedesign.asp"-->
<%

Print_Invoices = request.querystring("Orders")
Oids = Print_Invoices
Oid = replace(Print_Invoices,",","*")
sMenu="orders"
sTitle="Orders > Pick List"

if Print_Invoices = "" then
   fn_error "You must select at least one order"
else
createHead thisRedirect
on error resume next
%>
	<tr bgcolor='FFFFFF'>
		<td width="100%" colspan="3">
			
				<table width="600" height="10" border=0>
				<tr bgcolor='FFFFFF'>
					<td colspan=2>
					<a href="pick_list_print.asp?Orders=<%=Print_Invoices%>" class=link target=_blank><b>Printable Copy</b></a> &nbsp; &nbsp;
                </tr>
                
               <tr bgcolor='FFFFFF'>
				<td colspan=2>
						<hr width="600" align="left">
						Pick List   Order # : <% =Oids %>
								
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
	response.write "<tr bgcolor='FFFFFF'><td width='100%' colspan='3' height='74'><div id='pgBreak' style='page-break-after:always;'><img src=images/spacer.gif width=1 height=2></div></td></tr>"
	
end if			' querystring condn

%>
<% createFoot thisRedirect, 0%>
