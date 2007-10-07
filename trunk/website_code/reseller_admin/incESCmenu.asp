<%
Response.Buffer = true
Response.Expires = -1441

'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page contains the menu for the easystorecreator admin	
'	Page Name:		   	incESCmenu.asp
'	Version Information:EasystoreCreator
'	Input Page:		    included in all page of ESC admin
'	Output Page:	    
'	Date & Time:		12 th August
'	Created By:			Devki Anote
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 

%>
<TABLE border=0 cellPadding=0 cellSpacing=0>
       <TBODY>
        <TR>
          <TD class=menu<%=msel1%> noWrap><A class=menu<%=msel2%> href="easycreator_admin.asp" >Home</A></TD>
            
          <TD class=menu<%=msel3%> noWrap><A class=menu<%=msel4%> href="reseller_List.asp">Reseller List</A></TD>
          
          <TD class=menu<%=msel5%> noWrap><A class=menu<%=msel6%> href="Resellers_planpricing.asp">Reseller Plan Pricing</A></TD>
          <TD class=menu<%=msel7%> noWrap><A class=menu<%=msel8%> href="Resellers_Reports.asp">Reports</A></TD>
          <TD class=menu<%=msel11%> noWrap><A class=menu<%=msel12%> href="Reseller_fee_status.asp">License Fee Status</A></TD>
          <TD class=menu<%=msel13%> noWrap><A class=menu<%=msel14%> href="refund_Request.asp">Accept Requests</A></TD>
          <TD class=menu<%=msel15%> noWrap><A class=menu<%=msel16%> href="Update_Refund.asp">Refund Status</A></TD>
          <TD class=menu<%=msel9%> noWrap><A class=menu<%=msel10%> href="logout.asp">Logout</A></TD>
          <TD align=right class=menu noWrap width="100%"><FONT color=white>Admin</FONT></TD>
        </TR>
        </TBODY>
</TABLE>