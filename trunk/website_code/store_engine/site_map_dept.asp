<!--#include file="include/header.asp"-->
<%
Server.ScriptTimeout = 10

set deptfields1=server.createobject("scripting.dictionary")
sql_select="SELECT top 1000 full_name FROM store_dept WITH (NOLOCK) WHERE Store_id="&Store_id&" AND Visible=1 order by Full_Name, View_Order"
fn_print_debug sql_select
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
sText = "<table width='100%'>"

on error goto 0
Sub_Department_Id_temp = ""
FOR deptrowcounter1= 0 TO deptfields1("rowcount")
         Full_name = deptdata1(deptfields1("full_name"),deptrowcounter1)
	    sLink = fn_dept_url(Full_name,"")
	    sString="<B><a href='"&sLink&"' class=link>"&Full_name&"</a></b>"
	    sText = sText & ("<tr><td colspan=2>"&sString&"</td></tr>")

Next
sText=sText&("</table>")

set deptfields1 = Nothing

response.Write sText

%>
<!--#include file="include/footer.asp"-->
