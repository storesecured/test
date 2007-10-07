<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Department"
sHeader="department"
sMenu="inventory"

Dim arrColumns(13)
arrColumns(0) = "Department_Id:Op:Integer:"
arrColumns(1) = "Department_Name:Re:Text:50"
arrColumns(2) = "Department_Description:Op:Text:2500"
arrColumns(3) = "Belongs_To:Op:Text:200"
arrColumns(4) = "Department_Image_Path:Op:Text:100"
arrColumns(5) = "Department_Html:Op:Html:3000"
arrColumns(6) = "Visible:Op:Boolean::Y"
arrColumns(7) = "View_Order:Op:Integer::0"
arrColumns(8) = "Meta_Keywords:Op:Text:1500"
arrColumns(9) = "Meta_Description:Op:Text:1500"
arrColumns(10) = "Protect_Page:Op:Boolean::N"
arrColumns(11) = "Show_Name:Op:Boolean::N"
arrColumns(12) = "Meta_Title:Op:Text:100"
arrColumns(13) = "Department_Html_Bottom:Op:Html:3000"
sProcName = "wsp_department_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
