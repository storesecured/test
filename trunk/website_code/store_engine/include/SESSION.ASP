
<%

if isNull(Enable_Ip_tracking) or Enable_Ip_tracking="" then
   Enable_Ip_tracking=1
end if

UserAgent = replace(left(Request.ServerVariables("HTTP_USER_AGENT"),250),"'","''")
if fn_is_bot (UserAgent) = 1 then
   'dont lookup info for bots
   Shopper_Id=0
else
    if isNumeric(fn_get_querystring("Shopper_Id")) then
    		Shopper_Id = fn_get_querystring("Shopper_Id")
	end if
	fn_print_debug "shopper_id="&shopper_id
    if Shopper_Id="" then
       Shopper_Id = Session("Shopper_Id")
    end if
    fn_print_debug "session shopper_id="&Session("Shopper_Id")
    
    if fn_get_querystring("NewShopper")<>"" then
       Session("Shopper_Id")=""
       Shopper_Id=""
       NewShopper=1
    end if
    setlocale("en-us")

    Server_Name = Request.ServerVariables("SERVER_NAME")
    if Session("Shopper_Id")="" or Session("Shopper_Id")<>Shopper_Id and NewShopper<>1 then
      Referer = replace(left(Request.ServerVariables("HTTP_REFERER"),240),"'","''")
      ClientIP = Request.ServerVariables("REMOTE_ADDR")
      Server_Name = Request.ServerVariables("SERVER_NAME")
      Accept = Request.ServerVariables("HTTP_ACCEPT")
    end if
    
    came_from=fn_get_querystring("CAME_FROM")
    if came_from<>"" then
        response.cookies("EASYSTORECREATOR1")("CAME_FROM") = CAME_FROM_Cookie
	    response.cookies("EASYSTORECREATOR1").expires = DateAdd("d",Affiliate_Cookie,Now())
    else
        came_from = fn_sanitize(request.cookies("EASYSTORECREATOR1")("CAME_FROM"))
    end if
    
    if Shopper_Id=Null or Shopper_Id="" then
        Shopper_Id=0
    end if

    'only use shopper id when switching hosts so its a good idea to look up the information again from the db
    SQL_select_Shopper_Id = "exec wsp_shopper_lookup "&Store_Id&","&Shopper_Id&",'"&ClientIP&"','"&UserAgent&"','"&Accept&"','"&CAME_FROM&"','"&Server_Name&"','"&Referer&"',"&cint(Enable_Ip_tracking)&";"
    session("sql")=SQL_select_Shopper_Id
    fn_print_debug SQL_select_Shopper_Id
    
    rs_store.open SQL_select_Shopper_Id,conn_store,1,1
    if not rs_store.bof and not rs_store.eof then
        Shopper_Id = cstr(rs_Store("Shopper_Id"))
        came_from = rs_Store("came_from")
        oid = rs_Store("oid")
        if oid="" or oid=Null or not(isNumeric(this_oid)) then
        	oid=0
        end if
        Cid = rs_Store("Cid")
        Redirect_To_Cart = rs_Store("Redirect_To_Cart")
        rs_store.close
        
        if cid<>"0" then
            SQL_select = "exec wsp_customer_lookup "&Store_Id&","&cid&",0;"
            session("sql")=SQL_select
		  fn_print_debug SQL_select
            rs_store.open SQL_select,conn_store,1,1
            on error resume next
            user_id = rs_store("user_id")
            if err.number<>0 then
            	'customer id doesnt exist, so remove it
            	sql_update = "exec wsp_customer_logout "&store_id&","&Shopper_id&";"
            	conn_store.execute sql_update
            	fn_redirect Switch_Name
            end if
            on error goto 0
            password = rs_store("password")
            Protected_page_access = rs_Store("Protected_page_access")
            first_name = rs_store("first_name")
            last_name = rs_store("last_name")
            company = rs_store("company")
            address1 = rs_store("address1")
            address2 = rs_store("address2")
            city = rs_store("city")
            state = rs_store("state")
            zip = rs_store("zip")
            country = rs_store("country")
            email = rs_store("email")
            phone = rs_store("phone")
            fax = rs_store("fax")
            spam = rs_store("spam")
            is_residential = rs_store("is_residential")
            ccid = rs_store("ccid")
            orders_dept = rs_store("orders_dept")
            tax_exempt = rs_store("tax_exempt")
            budget_left = rs_store("budget_left")
            reward_left = rs_store("reward_left")
            rs_store.close
            Groups=fn_Get_Cid_Groups(cid)
        else
            first_name = ""
            last_name = ""
            budget_left = 0
            reward_left = 0
            Groups=""
        end if
    else
        rs_store.close
    end if
    fn_print_debug "Groups="&Groups
end if
Session("Shopper_Id")=Shopper_Id

function fn_is_bot (sUserAgent)
	sBotList="pingdom,openisearch,heeii,larbin,spider,wordpress,clamav,cfetch/1.0,java/1.,sortprice,zeus,shelob,pagebull,walhello,voyager,siteuptime,mediapartners-google,webwombat,feedfetcher,spider,ia_archiver,websitepulse,checker,bot,crawl,seek,spider,archive,scooter,seek,quality,creep,slurp,internetseer,link,retrieve,nameprotect,firefly,ask jeeves,fetch,gather,survey,extract,filter,libwww-perl"
	if fn_Is_Instr_Collection(sBotList,sUserAgent,",") then
		fn_is_bot=1
	else
		fn_is_bot=0
	end if
end function

%>
