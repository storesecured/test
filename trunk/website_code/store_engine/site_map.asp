<!--#include file="include/header.asp"-->
<%
Server.ScriptTimeout = 10

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if

set deptfields1=server.createobject("scripting.dictionary")
sql_select="SELECT top 1000 Item_Name, Item_Page_Name, full_name FROM sv_items_dept_combine WITH (NOLOCK) WHERE Store_id="&Store_id&" AND Show=1 AND visible=1 AND (0="&shide_out_stock&" OR quantity_control=0 OR (quantity_in_stock>quantity_control_number)) order by Full_Name, View_Order, Item_Name"
fn_print_debug sql_select
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
sText = "<table>"

on error goto 0
Sub_Department_Id_temp = ""
FOR deptrowcounter1= 0 TO deptfields1("rowcount")
	 Item_Name = deptdata1(deptfields1("item_name"),deptrowcounter1)
	 Item_Page_Name = deptdata1(deptfields1("item_page_name"),deptrowcounter1)
	 Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
	 if Full_Name <> Full_Name_temp then
	    sDept=""
	    sLink = fn_dept_url("","")
	    sString="<B><a href='"&sLink&"' class=link>Top</a></b>"
	    for each sDeptName in split(Full_Name," > ")
		     if sDept = "" then
 			    sDept = sDeptName
		     else
 			    sDept = sDept & " > " & sDeptName
		     end if
		     sLink = fn_dept_url(sDept,"")
		     sString = sString & ("<b> > <a href='"&sLink&"' class=link>"&sDeptName&"</a></b>")
	    next
		sText = sText & "<tr><td colspan=2>"&sString&"</td></tr>"
	    Full_Name_temp=Full_Name
	 end if
	 if Item_Name<>"" then
	    sLink = fn_item_url(Full_Name,Item_Page_Name)
	    sText = sText & "<tr><td width='10%'>&nbsp;</td><td><a href='"&sLink&"' class=link>"&Item_Name&"</a></td></tr>"
	 end if
Next

sText = sText & "<tr><td colspan=2 class=link><b>Pages</b></td></tr>"
set deptfields1=server.createobject("scripting.dictionary")

sql_select="SELECT Page_Name, File_Name, Is_Link from store_pages WITH (NOLOCK) where Store_id="&Store_id&" and (Navig_Link_Menu<>0 or Navig_Button_Menu<>0) order by view_order,page_name"
fn_print_debug sql_select
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)

FOR deptrowcounter1= 0 TO deptfields1("rowcount")
	 Page_Name = deptdata1(deptfields1("page_name"),deptrowcounter1)
	 File_Name = deptdata1(deptfields1("file_name"),deptrowcounter1)
	 Is_Link = deptdata1(deptfields1("is_link"),deptrowcounter1)
	 sLink = fn_page_url(File_Name,Is_Link)
	 sText = sText & "<tr><td width='10%'>&nbsp;</td><td><a href='"&sLink&"' class=link>"&Page_Name&"</a></td></tr>"
Next

sText = sText & "<tr><td width='10%'>&nbsp;</td><td><BR><BR><font class=small>Powered by EasyStoreCreator <a href=http://www.easystorecreator.com class=link target=_blank><font class=small>Shopping Cart Software</font></a></font></td></tr>"
sText = sText & "</table>"

set deptfields1 = Nothing

response.Write sText

%>
<!--#include file="include/footer.asp"-->
