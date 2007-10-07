<!--#include file="Global_Settings.asp"-->

<%
Item_Id = request.querystring("Item_Id")
sql_select = "exec wsp_item_copy "&Store_Id&","&Item_Id&";"
conn_store.Execute sql_select

sql_select = "select max(item_id) as item_id from store_items where store_id="&Store_Id&" and item_name='"&checkstringforQ(request.querystring("Item_Name"))&" Copy'"
rs_Store.open sql_select,conn_store,1,1
if not rs_store.eof then
    Item_Copy=rs_Store("Item_Id")
else
    Item_Copy=Item_Id
end if
rs_Store.close


response.redirect "item_edit.asp?Sub_Department_Id="&Request.querystring("Sub_department_id")&"&Item_Id="&Item_Copy


%>
