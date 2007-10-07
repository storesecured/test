<%

   
'retreiing the name of the page that is being excecuted
tracking_page_name = trim(tracking_page_name)&".asp"  'this will be the hardcoded name in each page

sqlProc = "select fld_Reseller_ID,fld_Website,fld_Display_Name from tbl_Reseller_Master where fld_Website ='"&host&"'"
set rsName  = conn_store.Execute(sqlProc)
if not rsName.eof then
	intresellerid = rsName("fld_Reseller_ID")
	sDefaultUrl  = replace("www."&trim(rsName("fld_Website")),"www.","www"&sLocalAddName&".")
	Name = trim(rsName("fld_Display_Name"))
	Site_Name = trim(rsName("fld_Website"))
	site_host = Name
else
    fn_redirect_perm "http://www."&sDefaultUrl
end if



'code here to retreive the page id corresponing to that particula page name

sqlProc = "wsp_reseller_page_id '"&trim(tracking_page_name)&"'"
set rsProc = conn_store.Execute(sqlProc)
if not rsProc.eof then
    pageid = rsProc("fld_page_id")
else
    pageid=0
end if

sqlProc = "Get_Reseller_Page_Key " &intresellerid&" ,"&pageid
set rsProc = conn_store.Execute(sqlProc)
if not rsProc.eof then
		keyword1 = trim(rsProc("fld_Reseller_Page_Keyword1"))
		keyword2 = trim(rsProc("fld_Reseller_Page_Keyword2"))
		keyword3 = trim(rsProc("fld_Reseller_Page_Keyword3"))
		keyword4 = trim(rsProc("fld_Reseller_Page_Keyword4"))
		keyword5 = trim(rsProc("fld_Reseller_Page_Keyword5"))
end if 
'------------------------------------------------------------------------------------------------------------------------

'code here to retreive the description corresponding to that particular page id

'------------------------------------------------------------------------------------------------------------------------
sqlProcDesc = "Get_Reseller_Page_Desc  " &intresellerid&" ,"&pageid
set rsProcDesc  = conn_store.Execute(sqlProcDesc)
if not rsProcDesc.eof then 
	description = trim(rsProcDesc("fld_desc_text"))
end if 
'------------------------------------------------------------------------------------------------------------------------


'code here to retreive the LOGO image set by the reseller
'------------------------------------------------------------------------------------------------------------------------
getimagename="select fld_title_image from tbl_reseller_logo where fld_reseller_id="&intresellerid
set rsgetlogo=conn_store.execute(getimagename)
if not rsgetlogo.eof then
	logoimagename=trim(rsgetlogo("fld_title_image"))
end if 
'------------------------------------------------------------------------------------------------------------------------

if logoimagename="" then
   logoimagename=logo
end if


'************************code ends here******************************************************
'Code added by Sudha Ghogare on 13 th August 2004 to show reseller rates
if intResellerId <> "" then 
for i=1 to 5
			strget="Get_reseller_plan "&intResellerId&" ,"&i&""
			set rsget=conn_store.execute(strget)
			if not rsget.eof then
				redim preserve planrate(i)
				planrate(i)=formatnumber(trim(rsget("fld_rate")),2)
			end if
	
next	
end if
'*********************************Code ends here******************************

'code added here to calculate the yearly pricing 
'*******************code starts here****************************************
term = 12

 %>