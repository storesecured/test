<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/sub.asp"-->


<%

If not CheckReferer then
	Response.Redirect "admin_error.asp?message_id=2"
end if

If Form_Error_Handler(Request.Form) <> "" then
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><%

else
    if request.querystring("Delete") <> "" then
        if isNumeric(request.querystring("Delete")) then
            sql_delete = "exec wsp_page_delete "&store_id&","&page_id&";"
            conn_store.Execute sql_delete
            server.execute "reset_design.asp"
            Response.Redirect "page_manager.asp"
        end if
    else
        'RETRIEVE PARAMETHERS
        Page_name = checkStringForQ(Request.Form("Page_name"))
        if request.form("op") <> "edit" then
            sType="insert"
            'CREATE A NEW PAGE ID
            sql="select max(page_id)+1 as new_page_id from store_pages where store_id="&store_id&""
            rs_store.open sql,conn_store,1,1
	        page_id=rs_store("new_page_id")
            rs_store.close

            if isnull(page_id) or page_id < 50 then
	            page_id = 50
		    end if
	    else
            sType="update"
            Page_Id = request.form("Page_Id")
        end if

        Page_description = checkstringforQ(Request.Form("Page_description"))
        Location=request.form("Location")

        Navig_Link_Menu = Request.Form("Navig_Link_Menu")
        Navig_Button_Menu = Request.Form("Navig_Button_Menu")

        if Request.Form("protect_page") <> "" then
            protect_page = 1
        else
            protect_page = 0
        end if
        if Request.Form("use_template") <> "" then
            use_template = 1
        else
            use_template = 0
        end if
        View_Order = Request.Form("View_Order")
        Meta_Title = checkStringForQ(request.form("Meta_Title"))
        Meta_Keywords = checkStringForQ(request.form("Meta_Keywords"))
        Meta_Description = checkStringForQ(request.form("Meta_Description"))
        Page_Head = replace(request.form("Page_Head"),"'","''")
        File_name = request.form("file_name")
        Customer_group = Request.Form("Customer_group")
        if customer_group="" then
           customer_group=0
        end if
	 
        Page_Content_Top=nullifyQ(Request.Form("Page_Content_Top"))
        Page_Content_Top=replace(Page_Content_Top,Site_Name2&"images","/images")

        Page_Content_Bottom=nullifyQ(Request.Form("Page_Content_Bottom"))
        Page_Content_Bottom=replace(Page_Content_Bottom,Site_Name2&"images","/images")

        Page_Content_Top=replace(replace(Page_Content_Top,"<OBJ_TEXTAREA_START","<TEXTAREA"),"<OBJ_TEXTAREA_END","</TEXTAREA")
        Page_Content_Top=replace(replace(Page_Content_Top,"<OBJ_COMMENT_START>","<!--"),"<OBJ_COMMENT_END>","-->")
        Page_Content_Bottom=replace(replace(Page_Content_Bottom,"<OBJ_TEXTAREA_START","<TEXTAREA"),"<OBJ_TEXTAREA_END","</TEXTAREA")
        Page_Content_Bottom=replace(replace(Page_Content_Bottom,"<OBJ_COMMENT_START>","<!--"),"<OBJ_COMMENT_END>","-->")

        file_name=file_name&".asp"
        
        Page_Content_Top=replace(Page_Content_Top,"href=""#","href="""&file_Name&"#")
        Page_Content_Bottom=replace(Page_Content_Bottom,"href=""#","href="""&file_Name&"#")

	    sql_update = "exec wsp_page_update "&store_id&",'"&stype&"',"&page_id&","&customer_group&_
	        ","&protect_Page&","&use_template&",'"&meta_title&"','"&meta_keywords&"','"&Meta_Description&"','"&_
	        Page_name&"','"&Page_description&"','"&Page_Head&"','"&file_name&_
	        "',"&Navig_Link_Menu&","&Navig_Button_Menu&","&View_Order&_
	        ",'"&Page_Content_Top&"','"&Page_Content_Bottom&"';"
	    on error resume next
	    'response.write sql_update
	    'response.end
	    conn_store.execute sql_update
	    if err.number<>0 then
	        response.Redirect "error.asp?Message_id=101&Message_Add="&server.urlencode(err.description)
	    end if
	    on error goto 0
    end if
    
    server.execute "reset_design.asp"

    Response.redirect "new_page.asp?op=edit&id="&page_id

end if
%>

