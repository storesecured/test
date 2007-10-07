<!--#include file="include/header_noview.asp"-->

<% 
iPage_Id=fn_get_querystring("Page_Id")

if iPage_Id<>"" then
	iPage_Id=replace(iPage_id,"/","")
	if isNumeric(iPage_Id) then
		sql_select = "select file_name from store_pages WITH (NOLOCK) where store_id="&store_Id&" and page_id="&iPage_Id
		rs_store.open sql_select
		if not rs_store.eof then
			sFilename=rs_store("file_name")
			rs_store.close
			fn_redirect_perm Switch_Name&sFilename
		end if
		rs_store.close
	end if
end if

fn_redirect_perm Switch_Name&"404.asp"

%>
