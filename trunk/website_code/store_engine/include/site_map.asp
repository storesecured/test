<!--#include file="header.asp"-->
<%
Server.ScriptTimeout = 10

if Hide_outofStock_Items then
	sHide_Out_stock=1
else
   sHide_Out_stock=0
end if

set deptfields1=server.createobject("scripting.dictionary")
sql_select="SELECT top 2500 Item_Name, Item_Page_Name, full_name FROM Store_Items WITH (NOLOCK) left outer join store_dept WITH (NOLOCK) on store_items.store_id=store_dept.store_id and store_items.sub_department_id = store_dept.department_id WHERE Store_Dept.Store_id="&Store_id&" and (show<>0 or show is Null) AND (0="&shide_out_stock&" OR quantity_control=0 OR (quantity_in_stock>quantity_control_number)) order by Full_Name, store_items.View_Order, Item_Name"
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
sText = "<table>"

on error goto 0
Sub_Department_Id_temp = ""
FOR deptrowcounter1= 0 TO deptfields1("rowcount")
	 Item_Name = deptdata1(deptfields1("item_name"),deptrowcounter1)
	 Item_Page_Name = deptdata1(deptfields1("item_page_name"),deptrowcounter1)
	 Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
	 if Full_Name <> Full_Name_temp then
	    sDeptName=replace(Full_Name," > ","/")
	    sDeptString=""
	    sString="<B><a href='"&Switch_Name&"/items/list.htm' class=link>Top</a></b>"
	    for each sThisDeptName in split(Full_Name," > ")
		     if sDeptString = "" then
 			    sDeptString = sThisDeptName
		     else
 			    sDeptString = sDeptString & "/" & sThisDeptName
		     end if
		     sString = sString & ("<b> > <a href='"&Switch_Name&"items/"&sDeptString&"/list.htm' class=link>"&sThisDeptName&"</a></b>")
	    next
		sLink = Switch_Name&"items/"&sDeptName&"list.htm"
		sText = sText & "<tr><td colspan=2>"&sString&"</td></tr>"
	    Full_Name_temp=Full_Name
	 end if
	 if Item_Name<>"" then
	    sLink = Site_Name&"items/"&sDeptName&"/"&Item_Page_Name&"_detail.htm"
	    sText = sText & "<tr><td width='10%'>&nbsp;</td><td><a href='"&sLink&"' class=link>"&Item_Name&"</a></td></tr>"
	 end if
Next

sText = sText & "<tr><td colspan=2 class=link><b>Pages</b></td></tr>"
set deptfields1=server.createobject("scripting.dictionary")

sql_select="SELECT Page_Id, Page_Name, File_Name, Is_Link from store_pages where Store_id="&Store_id&" and (Navig_Link_Menu<>0 or Navig_Button_Menu<>0) order by view_order,page_name"
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)

FOR deptrowcounter1= 0 TO deptfields1("rowcount")
	 Page_Id = deptdata1(deptfields1("page_id"),deptrowcounter1)
	 Page_Name = deptdata1(deptfields1("page_name"),deptrowcounter1)
	 File_Name = deptdata1(deptfields1("file_name"),deptrowcounter1)
	 Is_Link = deptdata1(deptfields1("is_link"),deptrowcounter1)
	 if is_link<>0 then
	   sLink = File_Name
	 else
        if File_Name = "custom.asp" then
  	        sLink = Switch_Name&File_Name&"/page_id/"&Page_Id&"/Page_Name/"&replaceSpecial(Page_Name)
        elseif File_Name = "store.asp" then
            sLink = Switch_Name
        else
  	        sLink = Switch_Name&File_Name
        end if
	 end if
	 sText = sText & "<tr><td width='10%'>&nbsp;</td><td><a href='"&sLink&"' class=link>"&Page_Name&"</a></td></tr>"
Next

sText = sText & "<tr><td width='10%'>&nbsp;</td><td><BR><BR><font class=small>Powered by EasyStoreCreator <a href=http://www.easystorecreator.com class=link target=_blank><font class=small>Shopping Cart Software</font></a></font></td></tr>"
sText = sText & "</table>"

set deptfields1 = Nothing

response.Write sText

function replaceSpecial (sName)
   sNewName = sName
   sNewName = replace(sNewName," ","_")
   sNewName = replace(sNewName,"<","")
   sNewName = replace(sNewName,">","")
   sNewName = replace(sNewName,"/","")
   sNewName = replace(sNewName,"?","")
   sNewName = replace(sNewName,"&","")
   sNewName = replace(sNewName,"%","")
   sNewName = replace(sNewName,"=","")
   sNewName = replace(sNewName,".","")
   sNewName = replace(sNewName,",","")
   sNewName = replace(sNewName,"#34;","")
   sNewName = replace(sNewName,"#39;","")
   replaceSpecial = sNewName
end function

%>
<!--#include file="footer.asp"-->
