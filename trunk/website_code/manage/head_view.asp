<%
on error resume next
sRedirectArray = split(myStructure("FileName"),"?")
sRedirect = sRedirectArray(0)
sFormAction = myStructure("FormAction")
sTitle = myStructure("Title")
sFullTitle=myStructure("FullTitle")
thisRedirect = sRedirect
sMenu = myStructure("Menu")

if myStructure("HeaderList") = "" or isNull(myStructure("HeaderList")) then
   myStructure("HeaderList") = myStructure("ColumnList")
end if
if myStructure("SortList") = "" or isNull(myStructure("SortList")) then
   myStructure("SortList") = myStructure("ColumnList")
end if

if myStructure("EditRecord") = "" or isNull(myStructure("EditRecord")) then
   myStructure("EditRecord") = myStructure("NewRecord")
end if
if instr(myStructure("EditRecord"),"?") = 0 then
   myStructure("EditRecord") = myStructure("EditRecord") & "?"
else
   myStructure("EditRecord") = myStructure("EditRecord") & "&"
end if

if instr(myStructure("FileName"),"?") = 0 then
   myStructure("FileName") = myStructure("FileName") & "?"
else
   myStructure("FileName") = myStructure("FileName") & "&"
end if

RowsPerPage = 25
if request.querystring("Records") <> "" then
   RowsPerPage = request.querystring("Records")
elseif request.form("Records") <> "" then
   RowsPerPage = request.form("Records")
elseif Request.Cookies("EASYSTORECREATOR-"&myStructure("FileName"))("RowsPerPage")<>"" then
   RowsPerPage = Request.Cookies("EASYSTORECREATOR-"&myStructure("FileName"))("RowsPerPage")
end if
Start_Row = 0
End_Row = RowsPerPage - 1
if request.querystring("Start_Row") <> "" then
   Start_Row = cdbl(request.querystring("Start_Row"))
   if Start_Row < 0 then
      Start_Row = 0
   end if
   End_Row = cdbl(Start_Row) + cdbl(RowsPerPage)-1
end if

if request.querystring("Sort_By") <> "" then
	Sort_By = request.querystring("Sort_By")
elseif Request.Cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Sort_By")<>"" then
   Sort_By = Request.Cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Sort_By")
else
   Sort_By = myStructure("DefaultSort")
end if

if request.querystring("Sort_Direction") <> "" then
   Sort_Direction = request.querystring("Sort_Direction")
elseif Request.Cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Sort_Direction")<>"" then
   Sort_Direction = Request.Cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Sort_Direction")
end if

response.cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Sort_By") = Sort_By
response.cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Sort_Direction") = Sort_Direction
response.cookies("EASYSTORECREATOR-"&myStructure("FileName"))("RowsPerPage") = RowsPerPage
response.cookies("EASYSTORECREATOR-"&myStructure("FileName"))("Store_Id") = Store_id
response.cookies("EASYSTORECREATOR-"&myStructure("FileName")).expires = DateAdd("d",180,Now())

iFound = 0
	for each sColumn in split(myStructure("SortList"),",")
		if Sort_By = sColumn then
		   sort_by_array = split(Sort_By," ")
		   Sort_By_Compare = Sort_By
     Sort_By = sort_by_array(0)
			iFound = 1
			exit for
		end if
	next
	if iFound = 0 then
		'invalid column name
		Sort_By = myStructure("DefaultSort")
	end if

if instr(sFormAction,"?") = 0 then
   sFormAction = sFormAction & "?"
else
   sFormAction = sFormAction & "&"
end if
sFormAction = sFormAction&"Sort_By="&server.urlencode(Sort_By)&"&Sort_Direction="&Sort_Direction&"&Start_Row=0&Form="&Server.urlencode(myStructure("Form"))

 if  Request.Form("Delete_Id") ="SEL" then

	if myStructure("DeleteTable") = "" then
	myStructure("DeleteTable") = myStructure("TableName")
	end if

  if request.form("DELETE_IDS") <> "" then

		' delete file if table is store_pages
		' ------------------------------------
		tableName = myStructure("DeleteTable")
		
		sql_update = "delete from "&myStructure("DeleteTable")&" where "&myStructure("PrimaryKey")&" in ("&request.form("DELETE_IDS")&") and store_Id="&Store_Id
		conn_store.Execute sql_update

	end if
end if

if Request.Querystring("Delete_Id") <> "" then
  if myStructure("DeleteTable") = "" then
     myStructure("DeleteTable") = myStructure("TableName")
  end if
	
		tableName = myStructure("DeleteTable")
		
		' Get the file name of the deleted page
		' --------------------------------------
		if  trim(tableName) = "Store_Pages" then
		
		sql_filename  = "select file_name from "&myStructure("DeleteTable")&" where "&myStructure("PrimaryKey")&" = "&request.querystring("Delete_id")&" and store_Id="&Store_Id
		rs_Store.open sql_filename,conn_store,1,1
		if not rs_store.eof then
		file_name = Rs_store(0)
		end if
		rs_Store.Close

		Root_Folder = fn_get_sites_folder(Store_Id,"Root")
		Set FileObject = CreateObject("Scripting.FileSystemObject")
		File_Full_Name = Root_Folder&file_name

		' delete if file found
			if FileObject.FileExists(File_Full_Name) Then
			FileObject.DeleteFile(File_Full_Name)
			Set FileObject = Nothing
			end if 

		end if
		' --------------------------------------
		

		' Delete Query
		sql_update = "delete from "&myStructure("DeleteTable")&" where "&myStructure("PrimaryKey")&" = "&request.querystring("Delete_id")&" and store_Id="&Store_Id
		conn_store.Execute sql_update

end if

if sNeedSubmit<>1 or request.querystring<>"" or request.form<>"" then
	sql_select = "select "&myStructure("ColumnList")&" from "&myStructure("TableName")&" WITH (NOLOCK) where Store_id = "&Store_Id
	if myStructure("TableWhere") <> "" then
			sql_select = sql_select & " and " & myStructure("TableWhere")
	end if
	sql_select = sql_select & " order by "&Sort_By
	if Sort_Direction = 1 then
	   sql_select = sql_select & " desc"
	end if
     set myfields=server.createobject("scripting.dictionary")
	Call DataGetrowsPaged(conn_store,sql_select,mydata,myfields,noRecords,Start_Row,RowsPerPage)
     
	if myfields("rowcount")=-2 then
	   'we need to get the real rowcount
	   sql_select = "select count("&myStructure("PrimaryKey")&") from "&myStructure("TableName")&" where Store_id = "&Store_Id
	  if myStructure("TableWhere") <> "" then
		 sql_select = sql_select & " and " & myStructure("TableWhere")
	  end if
	  sGroup_by=instr(lcase(sql_select),"group by")
	  if sGroup_by>0 then
	     'group by doesnt work in counts
	     sql_select=left(sql_select,sGroup_by-1)
	  end if
	  set rs_store=conn_store.execute(sql_select)

	  myfields("rowcount")=rs_store(0)-1
	  rs_store.close

	end if

	if cdbl(End_Row) > cdbl(myfields("rowcount")) then
	   End_Row = myfields("rowcount")
	end if

	if cdbl(Start_Row) > cdbl(End_Row) then
	   Start_Row = 0
	end if
end if

createHead thisRedirect
%>
