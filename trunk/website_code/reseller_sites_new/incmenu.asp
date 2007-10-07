
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page contains the top menu
'	Page Name:		   	incmenu.asp
'	Version Information:EasystoreCreator
'	Input Page:		    included in all the reseller pages
'	Output Page:	    
'	Date & Time:		9 th August
'	Created By:			Devki Anote
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 


'code here to retreive the id of the logged in reseller 

sql = "select fld_reseller_id from tbl_reseller_master where fld_reseller_id="&session("ResellerID")
set rs = conn.execute(sql)

if not rs.eof then 
	resellerid = rs("fld_reseller_id")
end if
set rs1=conn.execute ("getsitename "&resellerid&" ")
if not rs1.eof then
sitename=trim(rs1("fld_website"))
'Response.Write "sitename"&sitename
end if

%>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
			<TD class=menu<%=mSel1%> noWrap><A class=menu<%=mSel2%> href="reseller_home.asp">Customize Sales Website</A></TD>
			<TD class=menu<%=mSel3%> noWrap><A class=menu<%=mSel4%> href="http://reseller.storesecured.com?reseller=<%=session("resellerid")%>" target=_blank>Preview</A></TD>
			<!--<TD class=menu<%=mSel3%> noWrap><A class=menu<%=mSel4%> href="http://enigma/esc/reseller_sales/default.asp?reseller=<%=session("resellerid")%>" target=_blank>Preview</A></TD>-->
			<TD class=menu<%=mSel5%> noWrap><A class=menu<%=mSel6%> href="Resellers_Customer_list.asp">Customer List</A></TD>
			<TD class=menu<%=mSel7%> noWrap><A class=menu<%=mSel8%> href="Resellers_Sales.asp">My Sales History</A></TD>
			<TD class=menu<%=mSel9%> noWrap><A class=menu<%=mSel0%> href="../logout.asp">Logout</A></TD>
			<TD align=right class=menu noWrap width="100%"><FONT color=white>Reseller Id #: <B><%=resellerid%></B></FONT></TD>
    </TR>
       </TBODY>
       </TABLE>
  