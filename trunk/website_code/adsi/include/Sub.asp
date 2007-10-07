
<%
Function fn_website_create (Store_Id,sDomainList)
    Set objComp = GetObject("IIS://" & ServerName & "/W3SVC")
    on error resume next
    
    Set objNew = GetObject("IIS://" & ServerName & "/W3SVC/"&Store_Id)
    if err.number<>0 then
        'site already exists delete it first
        response.Write "in delete site"
        objComp.Delete "IIS://" & ServerName & "/W3SVC/"&Store_Id
    end if
    err.number=0
    
    Set objNew = ObjComp.Create("IIsWebServer", Store_Id )
    objNew.ServerComment = Store_Id
    objNew.SetInfo
    
    sCompleted = fn_virtual_directory_create (Store_Id)
    sCompleted = fn_server_bindings (Store_Id, sDomainList)
    
    Set App = GetObject("IIS://" & ServerName & "/w3svc/" & Store_Id & "/ROOT")
    App.AppCreate2 POOLED
    Set App = Nothing
    
    if fServerVersion then
        'cant start on a local version cause only 1 site can run at once
        objNew.Start
    end if
    
    set objNew = Nothing
    Set objComp = Nothing
    
    fn_website_create = 1
end function

function fn_server_bindings (Store_Id, sDomainList)
    response.Write "in fn_server_bindings"
    iArray=0
    on error resume next
    sServerList=""
  
    IpAddy=null
    response.Write "<BR>"&sDomainList
    for each sDomain in split(sDomainList,",")
        if sDomain<>"" then
            sDomain = fn_url_rewrite(sDomain)
            sDomain = replace(sDomain,"www.","")
            if not(Is_In_Collection(sServerList,IpAddy & ":80:"&sDomain,",")) then
                sServerList=sServerList&","&IpAddy & ":80:" & sDomain
                sServerList=sServerList&","&IpAddy & ":80:" & "www." & sDomain
            end if
        end if
    next
    response.Write sServerList
    
    Dim aryBindings()
    iIndex=0
    for each thissite in split(sServerList,",")
            if thissite<>"" then
                ReDim Preserve aryBindings(iIndex)
                aryBindings(iIndex)= thissite
                response.Write "aryBindings("&iIndex&")= "&thissite
                iIndex=iIndex+1
            end if
    next
    
    Set objBind = GetObject("IIS://"&ServerName&"/W3SVC/"&Store_id)
    objBind.SetInfo
    objBind.ServerBindings = aryBindings

    'do the same for secure sites but replace port
    Dim aryBindingsSecure()
    iIndex=0
    sServerList=replace(sServerList,":80:",":443:")
    for each thissite in split(sServerList,",")
            if thissite<>"" then
                ReDim Preserve aryBindingsSecure(iIndex)
                aryBindingsSecure(iIndex)= thissite
                iIndex=iIndex+1
            end if
    next

    objBind.SecureBindings = aryBindingsSecure
    objBind.SetInfo
    
    set objBind = Nothing

    fn_server_bindings=1
End function

function fn_virtual_directory_create (Store_Id)
on error resume next
    Set objVirt = GetObject("IIS://" & ServerName & "/W3SVC/"&Store_Id)
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
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/images/images_themes")
    objRoot.Path = fn_get_code_folder("Images_Themes")
    objRoot.SetInfo
    
    Set objRoot = objVirt.Create ("IIsWebVirtualDir", "Root/images/images_"&SiteID)
    objRoot.Path = fn_get_sites_folder(Store_Id,"Images")
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
