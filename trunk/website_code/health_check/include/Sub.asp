
<%
Function fn_website_create (Store_Id)
    on error goto 0
    response.write "Store_id="&Store_Id
    Set objComp = GetObject("IIS://127.0.0.1/W3SVC")
    if err.number<>0 then
          response.write err.description
          response.write err.number
          response.end
    end if

    on error resume next
    Set objNew = GetObject("IIS://127.0.0.1/W3SVC/"&Store_Id)
    if err.number<>0 then
        'site already exists update bindings instead
        fn_server_bindings (Store_Id)
    else
        'reset error number and create site
        err.number=0
    
        Set objNew = ObjComp.Create("IIsWebServer", Store_Id )
        objNew.ServerComment = Store_Id
        objNew.SetInfo
        
        sCompleted = fn_virtual_directory_create (Store_Id)
        sCompleted = fn_server_bindings (Store_Id)
        
        Set App = GetObject("IIS://127.0.0.1/w3svc/" & Store_Id & "/ROOT")
        App.AppCreate2 POOLED
        Set App = Nothing
        
        if fServerVersion then
            'cant start on a local version cause only 1 site can run at once
            objNew.Start
        end if
    end if
       
    set objNew = Nothing
    Set objComp = Nothing
    
    fn_website_create = 1
end function

function fn_server_bindings (Store_Id)

    'response.Write "in fn_server_bindings"
    iArray=0
    on error goto 0
    sServerList=""
  
    IpAddy=null
    
    sql_select = "SELECT store_url from store_domains where store_id="&Store_Id&" and redirect_only=0"
    set myfieldsUrl=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,sql_select,mydataUrl,myfieldsUrl,noRecordsUrl)

    if noRecordsUrl = 0 then
	    FOR rowcounterUrl= 0 TO myfieldsUrl("rowcount")
	        Store_url = mydataUrl(myfieldsUrl("store_url"),rowcounterUrl)
                if Store_url<>"" then
                    Store_url = fn_url_rewrite(Store_url)
                    Store_url = replace(Store_url,"www.","")
                    if not(Is_In_Collection(sServerList,IpAddy & ":80:"&Store_url,",")) then
                        sServerList=sServerList&","&IpAddy & ":80:" & Store_url
                        sServerList=sServerList&","&IpAddy & ":80:" & "www." & Store_url
                    end if
                end if
        Next
    end if

    iBindingSize = ((myfieldsUrl("rowcount")+1)*2)-1

    Dim aryBindings()
    Redim aryBindings(iBindingSize)

    iIndex=-1
    for each thissite in split(sServerList,",")
            if thissite<>"" then
                iIndex=iIndex+1
                aryBindings(iIndex)= thissite
                
            end if
    next

    if iIndex<>iBindingSize then
       Redim Preserve aryBindings(iIndex)
       response.write "<BR>setting redim preserve = "&iIndex
    end if
    'response.write "<BR>iIndex ="&iIndex


    Set objBind = GetObject("IIS://127.0.0.1/W3SVC/"&Store_id)
    objBind.SetInfo
    objBind.ServerBindings = aryBindings

    'do the same for secure sites but replace port
    iIndex=-1
    sServerList=replace(sServerList,":80:",":443:")
    for each thissite in split(sServerList,",")
            if thissite<>"" then
                iIndex=iIndex+1
                aryBindings(iIndex)= thissite
                'response.Write "<BR>aryBindings("&iIndex&")= "&thissite
            end if
    next

    objBind.SecureBindings = aryBindings
    'response.write "<BR>iIndex ="&iIndex
    objBind.SetInfo
    
    set objBind = Nothing
    'set aryBindings = Nothing
    'response.end
    fn_server_bindings=1
End function

function fn_virtual_directory_create (Store_Id)
on error resume next
    Set objVirt = GetObject("IIS://127.0.0.1/W3SVC/"&Store_Id)
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root")
    
    objRoot.Path = fn_get_sites_folder(Store_Id,"Root")
    objRoot.DefaultDoc="store.asp, index.html, index.htm, default.html, default.htm"
    objRoot.AccessRead=True
    objRoot.AccessScript=True
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/common")
    objRoot.Path = fn_get_code_folder("Common")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/include")
    objRoot.Path = fn_get_code_folder("Include")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/store_engine")
    objRoot.Path = fn_get_code_folder("Store_Engine")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/images")
    objRoot.Path = fn_get_network_folder(Store_Id,"Images")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/images/images_themes")
    objRoot.Path = fn_get_code_folder("Images_Themes")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/images/images_"&Store_Id)
    objRoot.Path = fn_get_network_folder(Store_Id,"Images")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/logs")
    objRoot.Path = fn_get_code_folder("Logs")
    objRoot.AccessExecute=True
    objRoot.DefaultDoc="stats.pl"
    objRoot.SetInfo
    
    set objVirt = Nothing
    set objRoot = Nothing
    
end function
%>
