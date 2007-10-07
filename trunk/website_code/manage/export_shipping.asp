<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%
server.scripttimeout = 400
on error goto 0
sCountItems = 0

sFormAction = "export_items2.asp"
sName = ""
sFormName = "Emails"
sTitle = "Export Shipping"
sSubmitName = "Update"
thisRedirect = "export_shipping.asp"
sMenu = "importexport"
createHead thisRedirect
%>

<%
on error goto 0
sql_select_shipping = "select * from Store_Shippers_class_all where store_id="&store_id 
sql_select_shipping = sql_select_shipping & " order by Shippers_class"
'response.write "<BR>sql_select_items=" & sql_select_shipping
set rs_Shipping= server.createObject("ADODB.Recordset")
rs_Shipping.open sql_select_shipping, conn_store, 1, 1
if not rs_Shipping.eof then

	'=========================='==========================
	TxtExport = TxtExport& "Shippers_class"& chr(9) 
	TxtExport=TxtExport &"Name"& chr(9)
	TxtExport=TxtExport &"Matrix low"& chr(9)
	TxtExport=TxtExport &"Matrix high"& chr(9)
	TxtExport=TxtExport &"Base fee"& chr(9)
	TxtExport=TxtExport &"Weight fee"& chr(9)
	TxtExport=TxtExport &"Countries"& chr(9)
	TxtExport=TxtExport &"Ship Location"& chr(9)
	TxtExport=TxtExport &"Zip Start"& chr(9)
	TxtExport=TxtExport &"Zip End"& chr(9)
	TxtExport=TxtExport &"Always Insert"& chr(9)
	TxtExport=TxtExport & chr(13) & chr(10)
		'==================='==========================

	do while not rs_Shipping.eof 
		Shippers_class= rs_Shipping("Shippers_class")
		Shipping_method_name= rs_Shipping("Shipping_method_name")
		Matrix_low= rs_Shipping("Matrix_low")
		Matrix_high= rs_Shipping("Matrix_high")
		Base_fee= rs_Shipping("Base_fee")
		Weight_fee= rs_Shipping("Weight_fee")
		Countries= rs_Shipping("Countries")
		allcountry=""
		if not isNull(Countries) then
		sCountries= split(Countries,",")
		for each work in sCountries
		  	if not isnull(work) and not isempty(work) and work <> "" and isNumeric(work) Then
				select_data="select country from Sys_Countries where Country_id = ('"&work&"')" 
				set rs_store = server.createObject("ADODB.Recordset")
				if rs_store.state = 1 then rs_store.close  
				rs_store.open select_data, conn_store, 1, 1
				if not rs_store.eof then
					countryCode=""
					countryCode= rs_store("Country")
					allcountry= allcountry &":" & countryCode 
				else
					Error = Error & "Error in Country Field <BR>"&Countries&"<BR>"
				end if
			end if
		next
		end if
		if allcountry <> "" then
			allcountry= right(trim(allcountry),(len(trim(allcountry))-1))
		end if
		Ship_Location_Id= rs_Shipping("Ship_Location_Id")
		sql_select="select location_name from store_ship_location where Ship_Location_Id=" & Ship_Location_Id & " and store_id=" & store_id
		if rs_Store.state=1 then rs_Store.close
		rs_Store.open sql_select,conn_store,1,1
		if rs_Store.eof then
			Response.write "Error "
		else
			location_name=rs_store("location_name")
		end if
	 	rs_Store.close
		
		Zip_Start= rs_Shipping("Zip_Start")
		Zip_End= rs_Shipping("Zip_End")
		
		TxtExport =TxtExport & Shippers_class&chr(9)&Shipping_method_name& chr(9)&Matrix_low&chr(9)	&Matrix_high&chr(9)	&Base_fee& chr(9)&Weight_fee&chr(9)& allcountry&chr(9)&location_name& chr(9)& Zip_Start & chr(9)& Zip_End& chr(9)&"Y"&chr(9)
		TxtExport=TxtExport & chr(13) & chr(10)
		rs_Shipping.moveNext
	loop
end if
	set myFields=nothing
'response.end
Set FileObject = CreateObject("Scripting.FileSystemObject")

File_Name = "shipping_"&dateserial(year(now()),month(now()),day(now()))
File_Name = Replace(File_Name,"/","-")
eFile_Name = File_Name
File_Name = File_Name & ".txt"

Export_Folder = fn_get_sites_folder(Store_Id,"Export")
File_Full_Name = Export_Folder&File_Name

Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
MyFile.Write TxtExport
MyFile.Close
set FileObject = Nothing

fn_redirect "Download_Exported_Files.asp?export=shipping&new="&efile_name
%>
