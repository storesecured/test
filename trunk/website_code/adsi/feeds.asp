<!--#include virtual="common/connection.asp"-->
<!--#include file="include/sub.asp"-->
<!--#include virtual="common/common_functions.asp"-->

<%

server.scripttimeout=20000

Set fso = CreateObject("Scripting.FileSystemObject")

set storefields=server.createobject("scripting.dictionary")
sql_select="SELECT store_settings.store_id,site_name,use_domain_name,store_domain,store_name,item_display, item_rows,Hide_outofStock_Items,large_upload,froggle_enable,froggle_filename,froggle_username,froggle_password,froggle_server FROM store_settings WITH (NOLOCK) inner join store_external WITH (NOLOCK) on store_settings.store_id=store_external.store_id WHERE store_settings.store_id<>101 and trial_version=0 and store_cancel is Null and overdue_payment=0 and store_active=1 and (Feeds_last_created Is Null OR force_resend<>0 OR (large_upload=0 and datediff(d,Feeds_last_created,getdate())>=7) or (large_upload<>0 and datediff(m,Feeds_last_created,getdate())>=1)) order by last_login desc"
Call DataGetrows(conn_store,sql_select,storedata,storefields,noRecords)
FOR storerowcounter = 0 TO storefields("rowcount")
    if response.isclientconnected=0 then
		response.end
    end if
    sStore_Id = storedata(storefields("store_id"),storerowcounter)
    Store_name = storedata(storefields("store_name"),storerowcounter)
    Hide_outofStock_Items = storedata(storefields("hide_outofstock_items"),storerowcounter)
    item_display = storedata(storefields("item_display"),storerowcounter)
    item_rows = storedata(storefields("item_rows"),storerowcounter)
    Froggle_Enable = storedata(storefields("froggle_enable"),storerowcounter)
    Filename = storedata(storefields("froggle_filename"),storerowcounter)
    Username = storedata(storefields("froggle_username"),storerowcounter)
    Password = storedata(storefields("froggle_password"),storerowcounter)
    sServer = storedata(storefields("froggle_server"),storerowcounter)
    Site_Name = storedata(storefields("site_name"),storerowcounter)
    Store_Domain = storedata(storefields("store_domain"),storerowcounter)
    Use_Domain_Name = storedata(storefields("use_domain_name"),storerowcounter)
    Large_Upload = storedata(storefields("large_upload"),storerowcounter)
    iItemsPerPage=(item_display+1)*item_rows
    
    if Use_Domain_Name and Store_Domain<>"" then
        Site_Name=Store_Domain
    end if
    Site_Name="http://"&Site_Name&"/"
    Switch_Name=Site_Name


    if Hide_outofStock_Items then
	    sHide_Out_stock=1
    else
       sHide_Out_stock=0
    end if

    sRssHead = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf&_
        "<rss version=""2.0""><channel>"&_
        "<title>"&Store_Name&"</title>"&_
        "<description>"&Store_Name&" listing of products</description>"&_
        "<link>"&Site_Name&"</link>"&_
        "<lastBuildDate>"&return_RFC822_Date(now(),"PST")&"</lastBuildDate>"&_
        "<language>en</language>"&vbcrlf
    sRss=""
    
    Root_Folder = fn_get_sites_folder(sStore_Id,"Root")
    sSitemapFile = Root_Folder&"auto_sitemap.xml"
    sPagesSitemapFile = Root_Folder&"pages.xml"
    sDepartmentsSitemapFile = Root_Folder&"departments.xml"
    sItemsSitemapFile = Root_Folder&"items.xml"
    sPagesSitemapFile_Mod=""
    sDepartmentsSitemapFile_Mod=""
    sItemsSitemapFile_Mod=""

    if fso.FileExists( sPagesSitemapFile ) then
        Set fs = fso.GetFile(sPagesSitemapFile) 
        sPagesSitemapFile_Mod = fs.DateLastModified
    end if
    if fso.FileExists( sDepartmentsSitemapFile ) then
        Set fs = fso.GetFile(sDepartmentsSitemapFile) 
        sDepartmentsSitemapFile_Mod = fs.DateLastModified
    end if
    if fso.FileExists( sItemsSitemapFile ) then
        Set fs = fso.GetFile(sItemsSitemapFile) 
        sItemsSitemapFile_Mod = fs.DateLastModified
    end if
    
    sMainMap = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"&vbcrlf&_
    		"<sitemapindex xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9 http://www.sitemaps.org/schemas/sitemap/0.9/siteindex.xsd"" xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">"&vbcrlf

    sPagesSitemapFile_Mod=""
    sPagesSitemapFile_Mod = make_page_sitemap (sStore_Id,sPageSiteMapFile_Mod)
    sPagesSitemapFile_Mod = w3c_date(sPagesSitemapFile_Mod)
    sMainMap=sMainMap&"<sitemap><loc>"&Site_Name&"pages.xml</loc><lastmod>"&sPagesSitemapFile_Mod&"</lastmod></sitemap>"&vbcrlf


     sItemsSitemapFile_Mod=""
     sItemsSitemapFile_Mod = make_item_sitemap (sStore_Id,sItemsSitemapFile_Mod,iItemsPerPage,Large_Upload)
     if sItemsSitemapFile_Mod<>"" then
    		sItemsSitemapFile_Mod = w3c_date(sItemsSitemapFile_Mod)
    		sDepartmentsSitemapFile_Mod = sItemsSitemapFile_Mod
    		sMainMap=sMainMap&"<sitemap><loc>"&Site_Name&"departments.xml</loc><lastmod>"&sDepartmentsSitemapFile_Mod&"</lastmod></sitemap>"&vbcrlf&_
        "<sitemap><loc>"&Site_Name&"items.xml</loc><lastmod>"&sItemsSitemapFile_Mod&"</lastmod></sitemap>"&vbcrlf
     end if

     sMainMap=sMainMap&"</sitemapindex>"


   
    Set fs = fso.OpenTextFile(sSitemapFile, 2,true)
	fs.Write sMainMap
	fs.Close
     sql_update = "update store_External set feeds_last_created=getdate(),force_resend=0 where store_id="&sStore_Id
     conn_store.execute sql_update


next

function make_page_sitemap (sStore_Id,sPageSiteMapFile_Mod)

    if sPageSiteMapFile_Mod<>"" then
        page_change_count=0
        sql_count = "select count(page_id) from store_pages WITH (NOLOCK) where store_id="&sStore_Id&" and sys_modified>'"&sPageSiteMapFile_Mod&"'"
        rs_store.open sql_count, conn_store, 1, 1
	    if not rs_store.eof then
		    page_change_count = rs_store(0)
	    end if
	    rs_store.close

        if page_change_count=0 then
            make_page_sitemap = sPageSiteMapFile_Mod
            exit function
        end if
    end if
    
    set deptfields1=server.createobject("scripting.dictionary")
    sql_select="SELECT top 5000 Page_Id, Page_Name, File_Name, Is_Link, Sys_Modified, navig_button_menu, navig_link_menu FROM store_pages WHERE Store_id="&sStore_id&" AND allow_link=1 and (navig_link_menu<>0 or navig_button_menu<>0) order by view_order"
    Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)
    
    sSitePageMap = "<?xml version=""1.0"" encoding=""UTF-8""?>"&vbcrlf&_
	"<urlset xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9 http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"" xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">"&vbcrlf



    FOR deptrowcounter1= 0 TO deptfields1("rowcount")
        Page_Id = deptdata1(deptfields1("page_id"),deptrowcounter1)
        Page_Name = deptdata1(deptfields1("page_name"),deptrowcounter1)
        File_Name = deptdata1(deptfields1("file_name"),deptrowcounter1)
        Navig_button_menu = deptdata1(deptfields1("navig_button_menu"),deptrowcounter1)
        Navig_link_menu = deptdata1(deptfields1("navig_link_menu"),deptrowcounter1)
        Page_Mod = deptdata1(deptfields1("sys_modified"),deptrowcounter1)
        Page_Mod_W3C = w3c_date(Page_Mod)
        Is_Link = deptdata1(deptfields1("is_link"),deptrowcounter1)
        if Navig_button_menu<>0 or navig_link_menu<>0 then
            sPriority=.8
        else
            sPriority=.2
        end if
        on error goto 0
        sLink = fn_page_url(File_Name,Is_Link)


        if File_Name = "store.asp" then
            sPriority=1
        elseif File_Name = "browse_dept_items.asp" then
            sPriority=.9
        elseif instr(File_Name,"_action")>0 then
            sPriority=.1
        end if
        
        if Is_Link and instr(sLink,Site_Name)=0 then
        	'dont add its not in the right domain
        else
        	sSitePageMap = sSitePageMap & ("<url><loc>"&sLink&"</loc><lastmod>"&Page_Mod_W3C&"</lastmod><priority>"&sPriority&"</priority></url>"&vbcrlf)
	   end if
    Next

    sSitePageMap = sSitePageMap & "</urlset>"
    Set fs = fso.OpenTextFile(sPagesSitemapFile, 2,true)
	fs.Write sSitePageMap
	fs.Close
    
    make_page_sitemap = now()

end function


function make_item_sitemap (sStore_Id,sItemsSitemapFile_Mod,iItemsPerPage,Large_Upload)
    sExprDate=DateAdd("d",30,now())
    sYear=DatePart("yyyy",sExprDate)
    sMonth=DatePart("m",sExprDate)
    sDay=DatePart("d",sExprDate)
    if sMonth<10 then
    		sMonth="0"&sMonth
	end if
	if sDay<10 then
		sDay="0"&sDay
	end if
	sDateExpr=sYear&"-"&sMonth&"-"&sDay
	on error resume next
    if sItemsSitemapFile_Mod<>"" then
        item_change_count=0
        sql_count = "select count(item_id) from store_items WITH (NOLOCK) where store_id="&sStore_Id&" and sys_modified>'"&sItemsSitemapFile_Mod&"'"
        rs_store.open sql_count, conn_store, 1, 1
	    if not rs_store.eof then
		    item_change_count = rs_store(0)
	    end if
	    rs_store.close

        if item_change_count=0 then
            make_item_sitemap = sItemsSitemapFile_Mod
            response.write "exiting"
            exit function
        end if
    end if
    sFroogleHead="expiration_date"&chr(9)&"link"&chr(9)&"title"&chr(9)&"description"&chr(9)&"price"&chr(9)&"image_link"&chr(9)&"brand"&chr(9)&"condition"&chr(9)&"product_type"&chr(9)&"mpn"&vbcrlf
    
    set deptfields1=server.createobject("scripting.dictionary")
    if Large_Upload then
    	  sql_top=""
    	else
    		sql_top="top 3000"
    	end if
    sql_select="SELECT "&sql_top&" * FROM sv_items_dept_feed WITH (NOLOCK) WHERE Store_id="&sStore_id&" AND Show=1 AND (0="&shide_out_stock&" OR quantity_control=0 OR (quantity_in_stock>quantity_control_number)) order by Full_Name, Item_Name"
    Call DataGetrows(conn_store,sql_select,deptdata1,deptfields1,noRecords1)

    sSiteDeptMapHead = "<?xml version=""1.0"" encoding=""UTF-8""?>"&vbcrlf&_
        "<urlset xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9 http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"" xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">"&vbcrlf
    sSiteDeptMap=""

    sSiteItemMapHead = "<?xml version=""1.0"" encoding=""UTF-8""?>"&vbcrlf&_
        "<urlset xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9 http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"" xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">"&vbcrlf
    sSiteItemMap=""

    FOR deptrowcounter1= 0 TO deptfields1("rowcount")
        count_dept_items=count_dept_items+1
        Item_Name = deptdata1(deptfields1("item_name"),deptrowcounter1)
        Item_Name = StripHTML(Item_Name)
        Item_Page_Name = deptdata1(deptfields1("item_page_name"),deptrowcounter1)
        Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
        Item_Id = deptdata1(deptfields1("item_id"),deptrowcounter1)
        Item_Mod=deptdata1(deptfields1("item_mod"),deptrowcounter1)
        Item_Mod_W3C = w3c_date(Item_Mod)
        Item_Mod_RFC=return_RFC822_Date(Item_Mod,"PST")
        Meta_Title = deptdata1(deptfields1("meta_title"),deptrowcounter1)
        Meta_Description = deptdata1(deptfields1("meta_description"),deptrowcounter1)
        Description_L = deptdata1(deptfields1("description_l"),deptrowcounter1)
        Description_S = deptdata1(deptfields1("description_s"),deptrowcounter1)
        Full_Name = deptdata1(deptfields1("full_name"),deptrowcounter1)
        Brand = deptdata1(deptfields1("brand"),deptrowcounter1)
        Condition = deptdata1(deptfields1("condition"),deptrowcounter1)
        Product_type = deptdata1(deptfields1("product_type"),deptrowcounter1)
        
        strippedFull_Name = StripHTML(Full_Name)
	   if Meta_Description<>"" then
            sDescription=Meta_Description
        elseif Description_L<>"" then
            sDescription=Description_L
        else
            sDescription=Description_S
        end if
        sDescription=fn_replace(StripHTML(sDescription),chr(9),"")
        sDescription=fn_replace(sDescription,chr(10),"")
        sDescription=fn_replace(sDescription,chr(13),"")
        if Meta_Title<>"" then
            sTitle=Meta_Title
        else
            sTitle=Item_Name
        end if
        sTitle=StripHTML(sTitle)
        if Full_Name <> Full_Name_temp then
            Dept_Mod=deptdata1(deptfields1("dept_mod"),deptrowcounter1)
            Dept_Mod_W3C = w3c_date(Dept_Mod)
            department_id=deptdata1(deptfields1("department_id"),deptrowcounter1)
            sDeptLink = fn_dept_url(Full_name,"")
            sSiteDeptMap = sSiteDeptMap & ("<url><loc>"&sDeptLink&"</loc><lastmod>"&Dept_Mod_W3C&"</lastmod><priority>.7</priority></url>"&vbcrlf)
            Full_Name_temp=Full_Name
        end if
        sItemLink = fn_item_url(Full_Name,Item_Page_Name)
        
        sDescripRss=replace(sDescription,"""","")
        sTitleRss=replace(sTitle,"""","")

        sRss = sRss & ("<item><title>"&sTitleRss&"</title>"&_
            "<description>"&sDescripRss&"</description>"&_
            "<link>"&sItemLink&"</link>"&_
            "<guid>"&Site_Name&"browse_item_details.asp/Item_Id/"&Item_Id&"/categ_id/"&department_id&"</guid>"&_
            "<category domain="""&sDeptLink&""">"&strippedFull_Name&"</category>"&_
            "<pubDate>"&Item_Mod_RFC&"</pubDate></item>"&vbcrlf)
        
        sSiteItemMap = sSiteItemMap & ("<url><loc>"&sItemLink&"</loc><lastmod>"&Item_Mod_W3C&"</lastmod><priority>.5</priority></url>"&vbcrlf)
        iPrice = deptdata1(deptfields1("retail_price"),deptrowcounter1)

        if iPrice>0 then
            iSpecial = deptdata1(deptfields1("retail_price_special_discount"),deptrowcounter1)
			sStart = deptdata1(deptfields1("special_start_date"),deptrowcounter1)
			sEnd = deptdata1(deptfields1("special_end_date"),deptrowcounter1)
			sImageName = deptdata1(deptfields1("imagel_path"),deptrowcounter1)
			item_sku = deptdata1(deptfields1("item_sku"),deptrowcounter1)
			custom_link = deptdata1(deptfields1("custom_link"),deptrowcounter1)
			if iSpecial <> 0 and now() > sStart and now() < sEnd then
				iPrice = iPrice * (1-(iSpecial/100))
			end if
			iPrice = formatnumber(iPrice,2,0,0,0)
			if sImageName = "" or sImageName = "0" then
				sImageName = deptdata1(deptfields1("images_path"),deptrowcounter1)
			end if
			if sImageName <> "" then
				if Instr(sImageName,"http://") > 0 then
					sImagePath = sImageName
				else
					sImagePath = Site_Name&"images/"&sImageName
				end if
			else
		        sImagePath = ""
			end if
			sImagePath=replace(sImagePath,"//","/")
			sImagePath=replace(sImagePath,"http:/","http://")
			if instr(custom_link,"http://")>0 then
                sItemLink=custom_link
            end if
            sFroogle = sFroogle&(sDateExpr&chr(9)&sItemLink&chr(9)&Item_Name&chr(9)&sDescription&chr(9)&iPrice&chr(9)&sImagePath&chr(9)&Brand&chr(9)&Condition&chr(9)&product_type&chr(9)&item_sku&vbcrlf)

        end if
        on error resume next
        if (deptrowcounter1 = deptfields1("rowcount")) or ((deptrowcounter1 < deptfields1("rowcount")) and (Full_Name <> deptdata1(deptfields1("full_name"),deptrowcounter1+1))) then
            if count_dept_items > iItemsPerPage then
                FOR iPage= 0 TO cint(count_dept_items/iItemsPerPage)-1
                    if iPage<>0 then
                        sDeptLink = fn_dept_url(Full_name,iPage)
                        sSiteDeptMap = sSiteDeptMap & ("<url><loc>"&sDeptLink&"</loc><lastmod>"&Dept_Mod_W3C&"</lastmod><priority>.6</priority></url>"&vbcrlf)
                    end if
                next
            end if
            count_dept_items=0
        end if
    Next
    
    if sSiteDeptMap<>"" then
    		sSiteDeptMap = sSiteDeptMapHead & sSiteDeptMap & "</urlset>"
	else
		sSiteDeptMap=" "
	end if
	if sSiteItemMap<>"" then
    		sSiteItemMap = sSiteItemMapHead & sSiteItemMap & "</urlset>"
    	else
		sSiteItemMap=" "
	end if
    sRss = sRss & vbcrlf&"</channel></rss>"

    Set fs = fso.OpenTextFile(sItemsSitemapFile, 2,true)
	fs.Write sSiteItemMap
	fs.Close
	


	Set fs = fso.OpenTextFile(sDepartmentsSitemapFile, 2,true)
	fs.Write sSiteDeptMap
	fs.Close
	
	if sFroogle<>"" then
		on error goto 0
		Set fs = fso.OpenTextFile(Root_Folder&"froogle.txt", 2,true)
	    fs.Write sFroogleHead&sFroogle
	    fs.Close
    end if
    
	Set fs = fso.OpenTextFile(Root_Folder&"products.xml", 2,true)
    sRss=replace(sRss,"“","")
    sRss=replace(sRss,"”","")
    sRss=replace(sRss,"'","")
    sRss=replace(sRss,"©","")
    fs.Write sRssHead&sRss
    fs.Close
    
    if sSiteItemMap=" " then
    	make_item_sitemap = ""
    else
    	make_item_sitemap = now()
    end if
end function

Function return_RFC822_Date(myDate, offset)
    Dim myDay, myDays, myMonth, myYear
    Dim myHours, myMonths, mySeconds

    myDate = CDate(myDate)
    myDay = WeekdayName(Weekday(myDate),true)
    myDays = Day(myDate)
    myMonth = MonthName(Month(myDate), true)
    myYear = Year(myDate)
    myHours = zeroPad(Hour(myDate), 2)
    myMinutes = zeroPad(Minute(myDate), 2)
    mySeconds = zeroPad(Second(myDate), 2)

    return_RFC822_Date = myDay&", "& _
        myDays&" "& _
        myMonth&" "& _
        myYear&" "& _
        myHours&":"& _
        myMinutes&":"& _
        mySeconds&" "& _
        offset
End Function

Function zeroPad(m, t)
   zeroPad = String(t-Len(m),"0")&m
End Function

Function w3c_date (sDate)
    w3c_date = DatePart("yyyy",sDate)&"-"&zeroPad(DatePart("m",sDate),2)&"-"&zeroPad(DatePart("d",sDate),2)
end function

Function stripHTML(strHTML)
'Strips the HTML tags from strHTML
  if strHTML <> "" then
	 Dim objRegExp, strOutput
	 Set objRegExp = New Regexp

	 objRegExp.IgnoreCase = True
	 objRegExp.Global = True
	 objRegExp.Pattern = "<(.|\n)+?>"

	 'Replace all HTML tag matches with the empty string
	 strOutput = objRegExp.Replace(strHTML, "")

	 'Replace all < and > with &lt; and &gt;
	 strOutput = Replace(strOutput, "<", "&lt;")
	 strOutput = Replace(strOutput, ">", "&gt;")

	 strOutput = Replace(strOutput, vbcrlf, " ")
	 strOutput = Replace(strOutput, "&nbsp;", " ")
	 strOutput = Replace(strOutput, "&amp;", " and ")
	 strOutput = Replace(strOutput, "&#34;", " ")
      strOutput = Replace(strOutput, "&", " and ")
      strOutput = Replace(strOutput, "é", "e")
      strOutput = Replace(strOutput, " ", " ")
      strOutput = Replace(strOutput, "’", "'")
      strOutput = Replace(strOutput, "´", "'")
      strOutput = Replace(strOutput, "®", " ")
      strOutput = Replace(strOutput, "™", " ")
      strOutput = Replace(strOutput, "¼", " 1/4 ")
	 strOutput = Replace(strOutput, "½", " 1/2 ")
      strOutput = Replace(strOutput, "¾", " 3/4 ")
      strOutput = Replace(strOutput, "`", "`")
      strOutput = Replace(strOutput, "°", " ")
      strOutput = Replace(strOutput, "º", " ")
      strOutput = Replace(strOutput, "—", "-")
      strOutput = Replace(strOutput, "–", "-")
      strOutput = Replace(strOutput, "Ö", "O")
      strOutput = Replace(strOutput, "•", "-")
      strOutput = Replace(strOutput, "…", "...")
      strOutput = Replace(strOutput, "‘", "`")

	 stripHTML = strOutput	  'Return the value of strOutput
	 Set objRegExp = Nothing
  else
	 stripHTML = strHTML
  end if
End Function

%>