<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
server.scripttimeout = 400
on error goto 0
sCountItems = 0

sFormAction = "export_items2.asp"
sName = ""
sFormName = "Emails"
sTitle = "Export Departments"
sSubmitName = "Update"
thisRedirect = "export_departments.asp"
sMenu = "importexport"
createHead thisRedirect

sql_select = "SELECT Department_ID,Department_Name,Department_Description,Belong_to,Last_Level,Department_Image_Path,Department_Html,Visible,View_Order,Meta_Keywords,Meta_Description,Protect_Page,Show_Name,Meta_Title,Department_Html_Bottom from store_dept where Store_id="&Store_id

	rs_Store.open sql_select,conn_store,1,1
	if rs_Store.eof = false then
	TxtExporthead = "Department ID" & chr(9)&_
	              "Department Name" & chr(9)&_
	              "Department Description" & chr(9)&_
	              "Parent" & chr(9) &_
	              "Department Image Path" & chr(9)&_
	              "Department Html" & chr(9)&_
	              "Visible" & chr(9)&_
	              "View Order" & chr(9)&_
	              "Meta Keywords" & chr(9)&_
	              "Meta Description" & chr(9)&_
	              "Protect Page" & chr(9)&_
	              "Show Name" & chr(9)&_
                 "Meta Title" & chr(9)&_
                 "Department HTML Bottom"

	TxtExport=TxtExporthead & vbcrlf
	do while not rs_Store.Eof
			Department_ID=rs_Store("Department_ID")
			Department_Name=rs_Store("Department_Name")
			Department_Description=rs_Store("Department_Description")
			Belong_to=rs_Store("Belong_to")
			if Belong_to > 0 then
				sqlbelongto= "select Department_Name from store_dept where store_id=" & store_id & " and  Department_ID='" & Belong_to & "'"
				set rs_belongto= server.createObject("ADODB.Recordset")
				if rs_belongto.state=1 then rs_belongto.close
				rs_belongto.open sqlbelongto,conn_store,1,1
				if not rs_belongto.eof and not rs_belongto.bof then
				    lclbelontodept=rs_belongto("Department_Name")
				else
				    lclbelontodept="Does not Exist"
				end if
			else
				lclbelontodept="Main Department"
			end if
			'Last_Level=rs_Store("Last_Level")
			'if Last_Level=True then 
			'	Last_Level=1
			'else
			'	Last_Level=0
			'end if
			Department_Image_Path=rs_Store("Department_Image_Path")
			'Full_Name=rs_Store("Full_Name")
			Department_Html=rs_Store("Department_Html")
			Department_Html_Bottom=rs_Store("Department_Html_Bottom")
			Visible=rs_Store("Visible")
			if Visible=true then 
				Visible="Y"
			else
				Visible="N"
			end if
			View_Order=rs_Store("View_Order")
			Meta_Title=rs_Store("Meta_Title")
			Meta_Keywords=rs_Store("Meta_Keywords")
			Meta_Description=rs_Store("Meta_Description")
			Protect_Page=rs_Store("Protect_Page")
			if Protect_Page=true then
				Protect_Page="Y"
			else
				Protect_Page="N"
			end if
			Show_Name=rs_Store("Show_Name")
			if Show_Name=true then 
				Show_Name="Y"
			else
				Show_Name="N"
			end if

			TxtExportDept =Department_ID & chr(9) &_
			              removeLines(Department_Name) &  chr(9) &_
			              removeLines(Department_Description) &  chr(9) &_
			              lclbelontodept &  chr(9) &_
			              Department_Image_Path & chr(9) &_
			              removeLines(Department_Html) &  chr(9) &_
			              Visible &  chr(9) &_
			              View_Order &  chr(9) &_
			              removeLines(Meta_Keywords) &  chr(9) &_
			              removeLines(Meta_Description) &  chr(9) &_
			              Protect_Page &  chr(9) &_
			              Show_Name &  chr(9) &_
			              removeLines(Meta_Title) &  chr(9) &_
			              removeLines(Department_Html_Bottom)

			TxtExportDept=TxtExportDept & vbcrlf
			TxtExport=TxtExport&TxtExportDept
		rs_Store.moveNext
	loop
	end if
	set myFields=nothing

Set FileObject = CreateObject("Scripting.FileSystemObject")

File_Name = "departments_"&dateserial(year(now()),month(now()),day(now()))
File_Name = Replace(File_Name,"/","-")
eFile_Name = File_Name
File_Name = File_Name & ".txt"
Export_Folder = fn_get_sites_folder(Store_Id,"Export")
File_Full_Name = Export_Folder&File_Name
If FileObject.FileExists(File_Full_Name) Then
   FileObject.DeleteFile(File_Full_Name)
end if
Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
MyFile.Write TxtExport
MyFile.Close
set FileObject = Nothing
Response.Redirect "Download_Exported_Files.asp?export=departments&new="&efile_name


function removeLines(sText)
	sText=nulltoNothing(sText)
    removeLines=replace(replace(replace(replace(sText,vbcrlf,""),vbcr,""),vblf,""),chr(9),"")
end function

function nulltoNothing(sText)
	if not isnull(sText) then
		nulltoNothing = sText
	else
		nulltoNothing = ""
	end if
end function
%>
