<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include virtual="common/common_functions.asp"-->

<%

sFormAction = "Store_Settings.asp"
sName = "Store_Activation"
sFormName = "activation"
sTitle = "List All Links"
sFullTitle = "Design > <a href=page_links_manager.asp class=white>Links</a> > List All"
sSubmitName = "Store_Activation_Update"
thisRedirect = "links.asp"
sTopic="Activate"
sMenu = "design"
sQuestion_Path = "design/page_links.htm"
createHead thisRedirect

Switch_Name=Site_Name

set deptfields1=server.createobject("scripting.dictionary")
sql_select="SELECT item_name,Item_page_Name, full_name FROM sv_items_dept_combine WHERE Store_id="&Store_id&" order by Full_Name, Item_Name"
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)

on error goto 0
Sub_Department_Id_temp = ""
FOR deptrowcounter1= 0 TO deptfields1("rowcount")
	 Item_page_Name = deptdata1(deptfields1("item_page_name"),deptrowcounter1)
	 Item_Name = deptdata1(deptfields1("item_name"),deptrowcounter1)
	 full_name = deptdata1(deptfields1("full_name"),deptrowcounter1)
	 if Full_Name <> Full_Name_temp then
		sLink = fn_dept_url(full_name,"")
		sText = sText & "<tr bgcolor=ffffff><td colspan=2>"&full_name&":<a href='"&sLink&"'>"&sLink&"</a></td></tr>"
	    Full_Name_temp=Full_Name
	 end if
	 if Item_page_Name<>"" then
	    sLink = fn_item_url(full_name,item_page_name)
	    sText = sText & ("<tr bgcolor=ffffff><td width='10%'>&nbsp;</td><td>"&Item_Name&":<a href='"&sLink&"' class=link>"&sLink&"</a></td></tr>")
	 end if
Next

sText=sText & "<tr bgcolor=ffffff><td colspan=2 class=link><b>Pages</b></td></tr>"
set deptfields1=server.createobject("scripting.dictionary")

sql_select="SELECT Page_Id, Page_Name, File_Name,Is_Link from store_pages where Store_id="&Store_id&" and allow_link=1 order by page_name"
Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)

FOR deptrowcounter1= 0 TO deptfields1("rowcount")
	 Page_Id = deptdata1(deptfields1("page_id"),deptrowcounter1)
	 Page_Name = deptdata1(deptfields1("page_name"),deptrowcounter1)
	 File_Name = deptdata1(deptfields1("file_name"),deptrowcounter1)
	 Is_Link = deptdata1(deptfields1("is_link"),deptrowcounter1)
	 sLink = fn_page_url(File_name,is_Link)
	 sText = sText & ("<tr bgcolor=ffffff><td width='10%'>&nbsp;</td><td><B>"&Page_Name&":</b><a href="&sLink&" class=link>"&sLink&"</a></td></tr>")
Next

response.Write sText


set deptfields1 = Nothing

function replaceSpecial (sName)
   sNewName = sName
   sNewName = replace(sNewName," ","_")
   sNewName = replace(sNewName,"<","")
   sNewName = replace(sNewName,">","")
   sNewName = replace(sNewName,"/","")
   sNewName = replace(sNewName,"?","")
   sNewName = replace(sNewName,"&","")
   replaceSpecial = sNewName
end function

createFoot thisRedirect, 0%>

