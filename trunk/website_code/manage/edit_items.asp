<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/department_list.asp"-->
<!--#include file="include/location_list.asp"-->
<!--#include virtual="common/common_functions.asp"-->


<%

server.scripttimeout = 2000
on error resume next

sFlashHelp="items/items.htm"
sMediaHelp="items/items.wmv"
sZipHelp="items/items.zip"


if request.form="" and request.querystring("Form") <> "" then
   Form_Array = split(request.querystring("Form"),"^")
   col1 = Form_Array(0)
	col1_value = Form_Array(1)
	col1_oper = Form_Array(2)
	col2 = Form_Array(3)
	col2_value = Form_Array(4)
	Sub_Department_Id = Form_Array(5)
end if

if request.form<>"" then
	if request.form("col1")<>"" then
		col1=request.form("col1")
		col1_value=checkstringforQ(request.form("col1_value"))
		if request.form("col2")<>"" then
			col1_oper=request.form("col1_oper")
			col2=request.form("col2")
			col2_value=checkstringforQ(request.form("col2_value"))
		end if
	end if
end if

if col1<>"" then
	if col1<>"" then
		col1_new=replace(col1,"**",col1_value)
		if instr(col1_new,"<>0")>0 then
			'bit field
			col1_value=lcase(col1_value)
			if col1_value="y" or col1_value="yes" or col1_value="1" or col1_value="t" or col1_value="true" then
			else
				col1_new=replace(col1_new,"<>0","=0")
			end if
		elseif instr(col1_new,"'")=0 then
			'no quotes means it should be a number, check
			if not isNumeric(col1_value) then
				fn_error "The search has failed.  You have entered a non-numeric value ("&col1_value&") into a field that requires a number."
			end if
		elseif instr(col1_new,"created")>0 or instr(col1_new,"date")>0 then
			'keywords means its a date
			if not isDate(col1_value) then
				fn_error "The search has failed.  You have entered a non-date value ("&col1_value&") into a field that requires a date."
			else
             col1_value=formatdatetime(col1_value)
             col1_new=replace(col1,"**",col1_value)
         end if
		end if
		sql_where=sql_where & " ("&col1_new&") "
		if col2<>"" then
			col2_new=replace(col2,"**",col2_value)
			if instr(col2_new,"<>0")>0 then
				'bit field
				col2_value=lcase(col2_value)
				if col2_value="y" or col2_value="yes" or col2_value="1" or col2_value="t" or col2_value="true" then
				else
					col2_new=replace(col2_new,"<>0","=0")
				end if
			elseif instr(col2_new,"'")=0 then
				'no quotes means it should be a number, check
				if not isNumeric(col2_value) then
					fn_error "The search has failed.  You have entered a non-numeric value ("&col2_value&") into a field that requires a number."
				end if
			elseif instr(col2_new,"created")>0 or instr(col2_new,"date")>0 then
				'keywords means its a date
				if not isDate(col2_value) then
					fn_error "The search has failed.  You have entered a non-date value ("&col2_value&") into a field that requires a date."
				else
                 col2_value=formatdatetime(col2_value)
                 col2_new=replace(col2,"**",col2_value)
    			end if
    		end if
			sql_where=sql_where & col1_oper & " ("&col2_new&") "
		end if
		sql_where = "("&sql_where&") AND "
	end if
end if


if Sub_Department_Id="" then
	Sub_Department_Id = Request("Sub_Department_Id")
end if

if Sub_Department_Id<>"-1" and Sub_Department_Id <> "" then
	sql_where = sql_where & " item_id in (select item_id from store_items_dept where store_id="&store_Id&" and department_id="&sub_department_id&")"
else
	sql_where = sql_where & " 1=1"
end if


sSubmitName="submit"
sQuestion_Path = "inventory/edit_item.htm"
sNeedSubmit=1
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Items"
myStructure("TableWhere") = sql_where
myStructure("ColumnList") = "item_id,item_sku,item_name,retail_price"
if service_type>=5 then
   myStructure("ColumnList")=myStructure("ColumnList")&",quantity_in_stock"
end if
myStructure("HeaderList") = myStructure("ColumnList")&",details"

myStructure("DefaultSort") = "item_sku"
myStructure("PrimaryKey") = "item_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("Footer") = "<Input class=buttons type='Button' value='Quick Edit Selected' onclick=""JavaScript:goQuickEdit();""><Input class=buttons type='Button' value='Mass Edit Selected' onclick=""JavaScript:goMassEdit();""><Input class=buttons type='Button' value='Mass Edit Selected Attributes' onclick=""JavaScript:goMassAttrEdit();""><BR><input type=""button"" OnClick=JavaScript:self.location=""item_edit.asp"" class=""Buttons"" value=""Advanced Add Item"" name=""Create_new"">"
myStructure("BackTo") = ""
myStructure("Menu") = "inventory"
myStructure("FileName") = "edit_items.asp?Id="&request.querystring("Id")
myStructure("FormAction") = "edit_items.asp?Id="&request.querystring("Id")
myStructure("Title") = "Item Search"
myStructure("FullTitle") = "Inventory > Items > Search"
myStructure("CommonName") = "Item"
myStructure("NewRecord") = "item_basic_edit.asp"
myStructure("EditRecord") = "item_basic_edit.asp"
myStructure("viewItem") = "item_basic_edit.asp"
myStructure("Heading:details") = "Advanced Edit"
myStructure("Heading:item_id") = "Id"
myStructure("Heading:item_sku") = "Sku"
myStructure("Heading:item_name") = "Name"
myStructure("Heading:show") = "Visible"
myStructure("Heading:retail_price") = "Price"
myStructure("Heading:quantity_in_stock") = "Stock"
myStructure("Format:details") = "TEXT"
myStructure("Format:item_id") = "STRING"
myStructure("Format:item_sku") = "STRING"
myStructure("Format:item_name") = "STRING"
myStructure("Format:show") = "LOOKUP"
myStructure("Format:retail_price") = "CURR"
myStructure("Format:quantity_in_stock") = "INT"
myStructure("Link:details") = "item_edit.asp?Id=PK"
myStructure("Lookup:show") = "False:No^True:Yes"
myStructure("Form") = col1&"^"&col1_value&"^"&col1_oper&"^"&col2&"^"&col2_value&"^"&Sub_Department_Id

%>
<script langauge="JavaScript">

	function goMassAttrEdit(){
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked)
				tsel = tsel+1;
		if (tsel<=1){
			alert("To edit a single item please use the edit link.");
			return;}
		document.forms[0].action="item_mass_edit_attr.asp";
		document.forms[0].submit();}
		
	function goQuickEdit(){
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked)
				tsel = tsel+1;
		if (tsel<=1){
			alert("To edit a single item please use the edit link.");
			return;}
		document.forms[0].action="item_quick_edit.asp";
		document.forms[0].submit();}

	function goMassEdit(){
		tsel=0;
		for (i=0; i<document.forms[0].DELETE_IDS.length;i++)
			if (document.forms[0].DELETE_IDS[i].checked)
				tsel = tsel+1;
		if (tsel<=1){
			alert("To edit a single item please use the edit link.");
			return;}
		document.forms[0].action="item_mass_edit.asp";
		document.forms[0].submit();}	
	
		

</script>




<!--#include file="head_view.asp"-->

<%



if Request.Form("Export_To_Text_File") <> "" or Request.Form("Export_To_QuickBooks_File") <> "" then
    sql_select_items = "select dbo.wf_dept_name(i.store_id,item_id) as department_name, item_page_name,"&_
        "item_handling,brand,condition,product_type,attributes_exist, accessories_exist, "&_
        "location_name,item_sku,item_name,retail_price,wholesale_price,images_path, "&_
        "imagel_path,taxable,description_s,description_l,item_weight,quantity_in_stock,quantity_control, "&_
        "quantity_control_number,retail_price_special_discount,special_start_date,special_end_date, "&_
        "shipping_fee,file_location,item_remarks,fractional,show,item_id, "&_
        "view_order,meta_keywords,meta_description, meta_title,M_d_1, M_d_2, M_d_3, M_d_4, M_d_5, "&_
        "custom_link, u_d_1, u_d_2, u_d_3, u_d_4, u_d_5, u_d_1_name, "&_ 
        "u_d_2_name, u_d_3_name, u_d_4_name, u_d_5_name "&_
        "from store_items i WITH (NOLOCK) "&_ 
        "left outer join store_ship_location l WITH (NOLOCK) on "&_
        "i.ship_location_id=l.ship_location_id "&_
        "where i.store_id="&store_id 
	if myStructure("TableWhere") <> "" then
		sql_select_items = sql_select_items & " and " & myStructure("TableWhere")
	end if
	

	File_Name = "items_on_"&dateserial(year(now()),month(now()),day(now()))
	File_Name = Replace(File_Name,"/","-")
end if

if Request.Form("Export_To_Text_File") <> "" then
	
	'CODE FOR SAVING THE RESULTS IN A TEXT FILE, FOR
	'USING IN AN EXTERNAL SOFTWARE
 
	Export_Folder = fn_get_sites_folder(Store_Id,"Export")
	File_Full_Name = Export_Folder&File_Name&".txt"

	TxtExport="Item Sku{{***}}Departments{{***}}Item Name" & "{{***}}"&_
		"Retail Price{{***}}Cost{{***}}Small Image" & "{{***}}"&_
		"Large Image{{***}}Taxable{{***}}Small Description" & "{{***}}"&_
		"Large Description{{***}}Item Weight{{***}}Quantity in Stock{{***}}"&_
		"Quantity Control{{***}}Quantity Control Number{{***}}Retail Price Special Discount{{***}}"&_
		"Special Start Date{{***}}Special End Date{{***}}Special Description{{***}}"&_
		"Shiping Fee{{***}}Downloadable Filename{{***}}Item Remarks{{***}}Fractional Selling{{***}}"&_
		"Show{{***}}Item Filename{{***}}Attributes{{***}}View Order{{***}}"&_
		"Meta Keywords{{***}}Meta Description{{***}}Item Accessories{{***}}"&_
		"Meta Title{{***}}Ship Location{{***}}Extended Field 1{{***}}"&_
		"Extended Field 2{{***}}Extended Field 3{{***}}Extended Field 4{{***}}"&_
		"Extended Field 5{{***}}Custom Link{{***}}User Defined Field 1{{***}}"&_
		"Use User Defined Field 1{{***}}User Defined Field 2{{***}}"&_
		"Use User Defined Field 2{{***}}User Defined Field 3{{***}}"&_
		"Use User Defined Field 3{{***}}User Defined Field 4{{***}}"&_
		"Use User Defined Field 4{{***}}User Defined Field 5{{***}}"&_
		"Use User Defined Field 5{{***}}Handling Fee{{***}}"&_
		"Brand{{***}}Condition{{***}}Product Type{{***}}"&_
		"{{*****}}"

	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select_items,mydata,myfields,noRecords)
	i=0
	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			If Response.IsClientConnected=false Then
				exit for
			end if
			if rowcounter mod 500 = 0 then
			        'every 5 lines pause for 5 seconds, to prevent from using to much cpu     
                                Set WaitObj = Server.CreateObject ("WaitFor.Comp")
                                WaitObj.WaitForSeconds 5
                                set WaitObj=Nothing
                
                        end if
				Item_Id=mydata(myfields("item_id"),rowcounter)
				sDept_ids=fn_get_dept_ids(Item_Id)
				sDepartmentSql="select department_name from store_dept where store_id="&store_Id&" and department_Id in ("&sDept_ids&")"
				sDepartment_Names = fn_get_list(sDepartmentSql,":")
				TxtExportItem = mydata(myfields("item_sku"),rowcounter)& "{{***}}"&_
				sDepartment_Names& "{{***}}"&_
				mydata(myfields("item_name"),rowcounter)& "{{***}}"&_
				formatnumber(mydata(myfields("retail_price"),rowcounter),2,0,0,0)& "{{***}}"&_
				formatnumber(mydata(myfields("wholesale_price"),rowcounter),2,0,0,0)& "{{***}}"&_
				mydata(myfields("images_path"),rowcounter)& "{{***}}"&_
				mydata(myfields("imagel_path"),rowcounter)& "{{***}}"&_
				intoBit(mydata(myfields("taxable"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("description_s"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("description_l"),rowcounter))& "{{***}}"&_
				formatnumber(mydata(myfields("item_weight"),rowcounter),2,0,0,0)& "{{***}}"&_
				mydata(myfields("quantity_in_stock"),rowcounter)& "{{***}}"&_
				intoBit(mydata(myfields("quantity_control"),rowcounter))& "{{***}}"&_
				mydata(myfields("quantity_control_number"),rowcounter)& "{{***}}"&_
				mydata(myfields("retail_price_special_discount"),rowcounter)& "{{***}}"&_
				mydata(myfields("special_start_date"),rowcounter)& "{{***}}"&_
				mydata(myfields("special_end_date"),rowcounter)& "{{***}}{{***}}"&_
				formatnumber(mydata(myfields("shipping_fee"),rowcounter),2,0,0,0)& "{{***}}"&_
				mydata(myfields("file_location"),rowcounter)& "{{***}}"&_
				mydata(myfields("item_remarks"),rowcounter)& "{{***}}"&_
				intoBit(mydata(myfields("fractional"),rowcounter))& "{{***}}"&_
				intoBit(mydata(myfields("show"),rowcounter))& "{{***}}"&_
				mydata(myfields("item_page_name"),rowcounter)& "{{***}}"
			Item_Id = mydata(myfields("item_id"),rowcounter)
			if mydata(myfields("attributes_exist"),rowcounter)<>0 then
    			sql_select_items = "select * from store_items_attributes WITH (NOLOCK) where Item_Id="&Item_Id&" and Store_Id="&Store_Id
    			set myfields1=server.createobject("scripting.dictionary")
    			Call DataGetrows(conn_store,sql_select_items,mydata1,myfields1,noRecords1)
    			if noRecords1 = 0 then
    				FOR rowcounter1= 0 TO myfields1("rowcount")
    				   TxtExportItem = TxtExportItem&mydata1(myfields1("attribute_class"),rowcounter1)&":"
    				   TxtExportItem = TxtExportItem&mydata1(myfields1("attribute_value"),rowcounter1)&":"
    				   TxtExportItem = TxtExportItem&mydata1(myfields1("display_order"),rowcounter1)&":"
    				   if mydata1(myfields1("default"),rowcounter1) then
    					   sDefault = 1
    				   else
    					   sDefault = 0
    				   end if
    				   TxtExportItem = TxtExportItem&sDefault&":"
    				   Use_Item = mydata1(myfields1("use_item"),rowcounter1)
    
    				   if Use_Item then
    					   TxtExportItem = TxtExportItem&"-1:"
    					   TxtExportItem = TxtExportItem&mydata1(myfields1("iitem_id"),rowcounter1)&":"
    					   TxtExportItem = TxtExportItem&mydata1(myfields1("iitem_quantity"),rowcounter1)
    				   else
    					   TxtExportItem = TxtExportItem&"0:"
    					   TxtExportItem = TxtExportItem&mydata1(myfields1("attribute_price_difference"),rowcounter1)&":"
    					   TxtExportItem = TxtExportItem&mydata1(myfields1("attribute_weight_difference"),rowcounter1)&":"
    					   TxtExportItem = TxtExportItem&mydata1(myfields1("attribute_sku"),rowcounter1)&":"
    					   TxtExportItem = TxtExportItem&mydata1(myfields1("attribute_hidden"),rowcounter1)
    				   end if
    				   if rowcounter1+1 <= myfields1("rowcount") then
    					   TxtExportItem = TxtExportItem&"|"
    				   end if
    				Next
    			End If
         end if
         
			TxtExportItem = TxtExportItem&("{{***}}"&mydata(myfields("view_order"),rowcounter)& "{{***}}"&_
				removeLines(mydata(myfields("meta_keywords"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("meta_description"),rowcounter))& "{{***}}")
			if mydata(myfields("accessories_exist"),rowcounter)<>0 then
			    sAccessSql="select item_sku from store_items i inner join store_items_accessories a on i.store_id=a.store_id and i.item_id=a.accessory_item_id where a.store_id="&store_Id&" and a.item_id="&Item_Id
				sAccessories = fn_get_list(sAccessSql,":")
				TxtExportItem = TxtExportItem&(sAccessories)
			end if
			TxtExportItem = TxtExportItem&("{{***}}"&_
				mydata(myfields("meta_title"),rowcounter)& "{{***}}"&_
				mydata(myfields("location_name"),rowcounter)& "{{***}}"&_
				removeLines(mydata(myfields("m_d_1"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("m_d_2"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("m_d_3"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("m_d_4"),rowcounter))& "{{***}}"&_
				removeLines(mydata(myfields("m_d_5"),rowcounter))& "{{***}}"&_
				mydata(myfields("custom_link"),rowcounter)& "{{***}}"&_
				removeLines(mydata(myfields("u_d_1_name"),rowcounter))& "{{***}}"&_
				intoBit(mydata(myfields("u_d_1"),rowcounter))&"{{***}}"&_
				removeLines(mydata(myfields("u_d_2_name"),rowcounter))& "{{***}}"&_
				intoBit(mydata(myfields("u_d_2"),rowcounter))&"{{***}}"&_
				removeLines(mydata(myfields("u_d_3_name"),rowcounter))& "{{***}}"&_
				intoBit(mydata(myfields("u_d_3"),rowcounter))&"{{***}}"&_
				removeLines(mydata(myfields("u_d_4_name"),rowcounter))& "{{***}}"&_
				intoBit(mydata(myfields("u_d_4"),rowcounter))&"{{***}}"&_
				removeLines(mydata(myfields("u_d_5_name"),rowcounter))& "{{***}}"&_
				intoBit(mydata(myfields("u_d_5"),rowcounter))&"{{***}}"&_
				formatnumber(mydata(myfields("item_handling"),rowcounter),2,0,0,0)& "{{***}}"&_
				mydata(myfields("brand"),rowcounter)& "{{***}}"&_
				mydata(myfields("condition"),rowcounter)& "{{***}}"&_
				mydata(myfields("product_type"),rowcounter)& "{{***}}"&_
				"{{*****}}")
			dept_name = mydata(myfields("department_name"),rowcounter)
			TxtExport=TxtExport & TxtExportItem
			i=i+1
		Next
	end if
	set myFields=nothing

	TxtExport=replace(replace(TxtExport,"&#34;",""""),"&#39;","'")
     TxtExport=replace(TxtExport,chr(9),"")
	TxtExport=replace(TxtExport,"{{***}}",chr(9))
	TxtExport=replace(TxtExport,"{{*****}}",vbcrlf)

   Set FileObject = CreateObject("Scripting.FileSystemObject")
	If FileObject.FileExists(File_Full_Name) Then
      FileObject.DeleteFile(File_Full_Name)
   end if
   Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)

	MyFile.Write TxtExport
	MyFile.Close
	set FileObject=Nothing

	Response.Redirect "Download_Exported_Files.asp"

end if

if Request.Form("Export_To_QuickBooks_File") <> "" then
		
	'code for selected items 21june2007 sr
		
		selected_items = Request.Form("DELETE_IDS")
		
		'complete
	

	'CODE FOR SAVING THE RESULTS IN A QUICKBOOKS FORMAT FILE, FOR
	'USING IN QUICKBOOKS SOFTWARE

	Set FileObject = CreateObject("Scripting.FileSystemObject")
	Set rs_Details = Server.CreateObject("ADODB.Recordset")
	
	Export_Folder = fn_get_sites_folder(Store_Id,"Export")
	File_Name = File_Name&".iif"
	File_Full_Name = Export_Folder&File_Name
	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)

	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select_items,mydata,myfields,noRecords)

	TxtExport = "!INVITEM"& "{{***}}"&"NAME"& "{{***}}"&"REFNUM"& "{{***}}"&"TIMESTAMP"& "{{***}}"&_
		"INVITEMTYPE"& "{{***}}"&"DESC"& "{{***}}"&"PURCHASEDESC"& "{{***}}"&"ACCNT"& "{{***}}"&_
		"ASSETACCNT"& "{{***}}"&"COGSACCNT"& "{{***}}"&"QNTY"& "{{***}}"&"QNTY"& "{{***}}"&"PRICE"& "{{***}}"&_
		"COST"& "{{***}}"&"TAXABLE"& "{{***}}"&"PAYMETH"& "{{***}}"&"TAXVEND"& "{{***}}"&"TAXDIST"& "{{***}}"&_
		"PREFVEND"& "{{***}}"&"REORDERPOINT"& "{{***}}"&"EXTRA"& "{{***}}"&"CUSTFLD1"& "{{***}}"&_
		"CUSTFLD2"& "{{***}}"&"CUSTFLD3"& "{{***}}"&"CUSTFLD4"& "{{***}}"&"CUSTFLD5"& "{{***}}"&_
		"DEP_TYPE"& "{{***}}"&"ISPASSEDTHRU"& "{{***}}"&"HIDDEN"& "{{***}}"&"DELCOUNT"& "{{***}}"&_
		"USEID"& chr(13) & chr(10)

	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			If Response.IsClientConnected=false Then
				exit for
			end if
			if rowcounter mod 500 = 0 then
			        'every 500 lines pause for 5 seconds, to prevent from using to much cpu     
                                Set WaitObj = Server.CreateObject ("WaitFor.Comp")
                                WaitObj.WaitForSeconds 5
                                set WaitObj=Nothing
                
                        end if
			TxtExportItem = "INVITEM"& "{{***}}"&mydata(myfields("item_name"),rowcounter)& "{{***}}"&_
				mydata(myfields("item_id"),rowcounter)& "{{***}}"&mydata(myfields("sys_created"),rowcounter)& "{{***}}"&_
				"{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"&_
				formatnumber(mydata(myfields("quantity_in_stock"),rowcounter),2,0,0,0)& "{{***}}"&_
				formatnumber(mydata(myfields("quantity_in_stock"),rowcounter),2,0,0,0)& "{{***}}"&_
				formatnumber(mydata(myfields("retail_price"),rowcounter),2,0,0,0)& "{{***}}"&_
				formatnumber(mydata(myfields("wholesale_price"),rowcounter),2,0,0,0)& "{{***}}"&_
				intoBit(mydata(myfields("taxable"),rowcounter))&"{{***}}"&_
				"{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"&_
				"{{***}}"& "{{***}}"& "{{***}}"& "{{***}}"&"N"& "{{***}}"& "{{***}}"& "{{***}}"& chr(13) & chr(10)
			TxtExport=TxtExport&TxtExportItem
		Next
	end if
	set myFields=nothing
	TxtExport=replace(replace(TxtExport,"&#34;",""""),"&#39;","'")
	TxtExport=replace(TxtExport,chr(9),"")
	TxtExport=replace(TxtExport,"{{***}}",chr(9))

	MyFile.Write TxtExport
	myfile.close



end if


%>

<SCRIPT LANGUAGE="VBScript">
	
	sub display()


	
		Dim WshShell
		dim FValue
		dim SValue		
		dim temp
		dim depid		
		
		Set WshShell = CreateObject("wscript.shell") 	
		
		temp="<%=sql_select_items%>"



		if temp="" then
			msgbox "Please Click Export To Quickbooks button First",0,"Quickbooks"
			exit sub
		end if
		temp=temp & "|" & "TABLE=ITEMS" & "|" &"<%=selected_items%>"	



		'Run Quick Book Exe	
		WshShell.Run "C:\bin\Quick.exe" & " " & temp
		
		
		
		
		exit sub




		
		fvalue= "<%=col1_value%>"
		svalue="<%=col2_value%>"
		depid="<%=Sub_Department_Id%>"
								
		'Both Text boxes are empty then do nothing
		If fvalue="" and svalue="" then	
		'First text box not empty then pass it as argument with Where clause
			temp= "|" & "TABLE=ITEMS" & "|" & depid
					
		ElseIf FValue<>"" and SValue="" then
		
			Temp="WHERE " & "<%=col1_new%>" & " " & "|" & "TABLE=ITEMS"	& "|" & depid			
			
			
		'Second text box not empty then pass it as argument with Where clause	
		ElseIf Fvalue="" and SValue<>"" then
			Temp="WHERE " & "<%=col2_new%>" & " " & "|" & "TABLE=ITEMS"	& "|" & depid
		Else
		'both text box not empty then pass them as argument with Where clause				
			temp="WHERE " & "<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>" & " " & "|" & "TABLE=ITEMS" & "|" & depid		
		end if		
						
		
		
msgbox temp


exit sub		
		    WshShell.Run "C:\bin\Quick.exe" & " " & temp   

	end sub
	
	
</script>



<input type=hidden name=records value=<%= RowsPerPage %> ID="Hidden1">
<tr>
		<td width="100%" colspan="3" height="74" align=center>
			<table width='100%' border='0' cellspacing='1' cellpadding=2 ID="Table1">

				<TR bgcolor=FFFFFF>
					<td width="35%" class="inputname">Department</td>
					<td width="60%" class="inputvalue">
						<%= create_dept_list ("Sub_Department_Id",Sub_Department_ID,1,"-1|{}|All|{new}|") %>
						</td>
				</tr>
				<TR bgcolor=FFFFFF><td class=inputname width=40%>
					<select name=col1 ID="Select1">
					<% if col1<>"" then %>
						<option value="<%= col1 %>"><%= replace(replace(replace(replace(replace(col1,"**",""),"%%",""),"<>0","="),"1=1",""),"''","") %></option>
						<option value="">Select column to search on</option>
					<% else %>
						<option value="">Select column to search on</option>
					<% end if %>
					<option value="Item_Id = **">Item Id = (Number)</option>
					<option value="Item_Sku like '%**%'">Item Sku like</option>
					<option value="item_name like '%**%'">Item Name like</option>
					<option value="description_s like '%**%'">Small Description like</option>
					<option value="description_l like '%**%'">Large Description like</option>
					<option value="images_path like '%**%'">Small Image like</option>
					<option value="imagel_path like '%**%'">Large Image like</option>
					<option value="sys_created > '**'">Created After (Date)</option>
					<option value="sys_created < '**'">Created Before (Date)</option>
					<option value="retail_price >= **">Retail Price >= (Number)</option>
					<option value="retail_price <= **">Retail Price <= (Number)</option>
					<option value="wholesale_price >= **">Cost >= (Number)</option>
					<option value="wholesale_price <= **">Cost <= (Number)</option>
					<option value="item_handling >= **">Handling Fee >= (Number)</option>
					<option value="item_handling <= **">Handling Fee <= (Number)</option>
					<option value="item_weight >= **">Weight >= (Number)</option>
					<option value="item_weight <= **">Weight <= (Number)</option>
					<option value="waive_shipping<>0">Waive Shipping = (Y/N)</option>
					<option value="show_homepage<>0">Show on Homepage = (Y/N)</option>
					<% if Service_Type>=5 then %>
					<option value="quantity_in_stock >= **">Quantity in Stock >= (Number)</option>
					<option value="quantity_in_stock <= **">Quantity in Stock <= (Number)</option>
					<option value="taxable<>0">Taxable = (Y/N)</option>
					<option value="hide_price<>0">Hide Price = (Y/N)</option>
					<option value="quantity_control<>0">Quantity Control = (Y/N)</option>
					<option value="cust_price<>0">Custom Price = (Y/N)</option>
					<option value="show<>0">Visible = (Y/N)</option>
					<option value="fractional<>0">Fractional = (Y/N)</option>
					<option value="retail_price_special_discount >= **">Discount % >= (Number)</option>
					<option value="retail_price_special_discount <= **">Discount % <= (Number)</option>
					<option value="view_order >= **">View Order >= (Number)</option>
					<option value="view_order <= **">View Order <= (Number)</option>
					<% end if %>
					<% if Service_Type>=7 then %>
					<option value="use_price_by_matrix<>0">Quantity Discount = (Y/N)</option>
					<% end if %>
					<option value="m_d_1 like '%**%' OR m_d_2 like '%**%' or m_d_3 like '%**%' or m_d_4 like '%**%' or m_d_5 like '%**%'">Extended Field like</option>

					
					</td><td class=inputvalue><input type=text name=col1_value value="<%= col1_value %>"  ID="Text1" size=49>
					<select name=col1_oper ID="Select2">
					<option value=AND>AND</option>
					<option value=OR>OR</option>
					</select></td></tr>
					<TR bgcolor=FFFFFF><td class=inputname width=40%>
					<select name=col2 ID="Select3">
					<% if col2<>"" then %>
						<option value="<%= col2 %>"><%= replace(replace(replace(replace(replace(col2,"**",""),"%%",""),"<>0","="),"1=1",""),"''","") %></option>
						<option value="">Select column to search on</option>
					<% else %>
						<option value="">Select column to search on</option>
					<% end if %>
					<option value="Item_Id = **">Item Id = (Number)</option>
					<option value="Item_Sku like '%**%'">Item Sku like</option>
					<option value="item_name like '%**%'">Item Name like</option>
					<option value="description_s like '%**%'">Small Description like</option>
					<option value="description_l like '%**%'">Large Description like</option>
					<option value="images_path like '%**%'">Small Image Path like</option>
					<option value="imagel_path like '%**%'">Large Image Path like</option>
					<option value="sys_created > '**'">Created After (Date)</option>
					<option value="sys_created < '**'">Created Before (Date)</option>
					<option value="retail_price >= **">Retail Price >= (Number)</option>
					<option value="retail_price <= **">Retail Price <= (Number)</option>
					<option value="wholesale_price >= **">Cost >= (Number)</option>
					<option value="wholesale_price <= **">Cost <= (Number)</option>
					<option value="item_handling >= **">Handling Fee >= (Number)</option>
					<option value="item_handling <= **">Handling Fee <= (Number)</option>
					<option value="item_weight >= **">Weight >= (Number)</option>
					<option value="item_weight <= **">Weight <= (Number)</option>
					<option value="waive_shipping<>0">Waive Shipping = (Y/N)</option>
					<option value="show_homepage<>0">Show on Homepage = (Y/N)</option>
					<% if Service_Type>=5 then %>
					<option value="quantity_in_stock >= **">Quantity in Stock >= (Number)</option>
					<option value="quantity_in_stock <= **">Quantity in Stock <= (Number)</option>
					<option value="taxable<>0">Taxable = (Y/N)</option>
					<option value="hide_price<>0">Hide Price = (Y/N)</option>
					<option value="quantity_control<>0">Quantity Control = (Y/N)</option>
					<option value="cust_price<>0">Custom Price = (Y/N)</option>
					<option value="show<>0">Visible = (Y/N)</option>
					<option value="fractional<>0">Fractional = (Y/N)</option>
					<option value="retail_price_special_discount >= **">Discount % >= (Number)</option>
					<option value="retail_price_special_discount <= **">Discount % <= (Number)</option>
					<option value="view_order >= **">View Order >= (Number)</option>
					<option value="view_order <= **">View Order <= (Number)</option>										
					<% end if %>
					<% if Service_Type>=7 then %>
					<option value="use_price_by_matrix<>0">Quantity Discount = (Y/N)</option>
					<% end if %>
                                        <option value="m_d_1 like '%**%' OR m_d_2 like '%**%' or m_d_3 like '%**%' or m_d_4 like '%**%' or m_d_5 like '%**%'">Extended Field like</option>

						</td><td class=inputvalue><input type=text name=col2_value value="<%= col2_value %>" ID="Text2" size=60></td></tr>
				
			<TR bgcolor=FFFFFF><td colspan=3 align=center><input type="submit" value="Search Items" class=buttons ID="Submit1" NAME="Submit1">
			<% if Service_Type>=7 then %>
					<BR><input class="Buttons" name="Export_To_Text_File" type="submit" value="Export To Text" ID="Submit2">
					<input class="Buttons" name="Export_To_QuickBooks_File" type="submit" value="Export To Quickbooks" ID="Submit3">
			<% end if %>

				</td></tr>
				<% if Request.Form("Export_To_QuickBooks_File") <> "" then %>
		<TR bgcolor='#FFFFFF'><TD colspan="7" width="35%" align="center">Please read the <a href=quickbooks.asp class=link>Quickbooks instructions</a> and then Click on <b>Insert in QuickBook</b> button to Import data!!
                <BR>
                <input class="Buttons" name="Insert In to Quickbooks" type="button"  value="Insert in QuickBook" onclick ="Display()" ID="Button1">
                </td></tr>					
                                                 <% end if %>

	</table></td></tr>

<% if (request.querystring<>"" or request.form<>"")and Request.Form("Export_To_QuickBooks_File") = "" then %>
<!--#include file="list_view.asp"-->
<%
else
        if myStructure("AddAllowed") then %>
	       <tr bgcolor="#FFFFFF"><td><input type="button" OnClick=JavaScript:self.location="<%= myStructure("NewRecord") %>" class="Buttons" value="Add Item" name="Create_new">
               <input type="button" OnClick=JavaScript:self.location="item_edit.asp" class="Buttons" value="Advanced Add Item" name="Create_new"></td></tr>

        <% end if
end if

if Request.QueryString("Delete_Id") <> "" then
	Delete_Id = request.querystring("Delete_Id")
elseif Request.Form("Delete_Id") ="SEL" and request.form("DELETE_IDS") <> "" then
	Delete_Id = request.form("DELETE_IDS")
else
	Delete_Id = ""
end if

if Delete_Id<>"" then
	sql_delete = "delete from Store_items_accessories where store_Id="&Store_Id&" and Item_id in ("&delete_id&");"&_
		"delete from store_items_attributes where store_Id="&Store_Id&" and Item_id in ("&delete_id&");"&_
		"delete from Store_items_Price_Matrix where store_Id="&Store_Id&" and Item_id in ("&delete_id&");"&_
		"delete from Store_items_Price_group where store_Id="&Store_Id&" and Item_id in ("&delete_id&");"&_
		"delete from Store_items_configs_dets where store_Id="&Store_Id&" and Item_id in ("&delete_id&");"&_
		"delete from Store_items_dept where store_Id="&Store_Id&" and Item_id in ("&delete_id&");"&_
		"delete from store_items_configs where store_Id="&Store_Id&" and Item_id in ("&delete_id&")"
		
	conn_store.Execute sql_delete
end if
  
function removeLines(sText)
	sText=nulltoNothing(sText)
    removeLines=replace(replace(replace(replace(sText,vbcrlf,""),vbcr,""),vblf,""),"{{***}}","")
end function

function nulltoNothing(sText)
	if not isnull(sText) then
		nulltoNothing = sText
	else
		nulltoNothing = ""
	end if
end function

function intoBit(sText)
	if cint(sText)=0 then
		intoBit="N"
	else
		intoBit="Y"
	end if
end function

function fn_get_list (sSqlStatement,sJoin)
	on error goto 0
	sList=""
	set getfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sSqlStatement,getdata,getfields,noRecords)
	if noRecords = 0 then
	    FOR getrowcounter= 0 TO getfields("rowcount")
		    if sList="" then
		        sList=getdata("0",getrowcounter)
            else
                sList=sList&sJoin&getdata("0",getrowcounter)
            end if
	    next
	end if
	set getfields = Nothing
	fn_get_list=sList
end function

function fn_replace_tab(sText)
	if isNull(sText) then
		sText=""
	else
	     sText=replace(sText,"{{***}}","{{***}}")
	end if
	fn_remove_tab=sText
end function

createFoot thisRedirect, 0%>

