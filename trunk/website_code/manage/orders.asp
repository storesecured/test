<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
on error goto 0

server.scripttimeout=4800

End_Date = FormatDateTime(now(),2)
sQuestion_Path = "reports/orders.htm"

if request.form("SearchOrders") <> "" or request.form("Export_To_Text_File") <> "" or request.form("Export_To_Quickbooks_File") <> ""then

elseif request.form("Ship_Labels") <> "" then
	response.redirect "print_labels.asp?Orders="&server.urlencode(request.form("DELETE_IDS"))&"&Start_Cell="&request.form("Start_Cell")
elseif request.form("Ship") <> "" then
	response.redirect "ship_orders_batch.asp?Orders="&server.urlencode(request.form("DELETE_IDS"))
elseif request.form("Print_Orders") <> "" then
	response.redirect "order_printable_multi.asp?Orders="&server.urlencode(request.form("DELETE_IDS"))
elseif request.form("Print_Packing") <> "" then
	response.redirect "order_printable_multi.asp?Orders="&server.urlencode(request.form("DELETE_IDS"))&"&Cart_Type=Packing"
elseif request.form("Print_Invoices") <> "" then
	response.redirect "order_printable_multi.asp?Orders="&server.urlencode(request.form("DELETE_IDS"))&"&Cart_Type=Invoice"
elseif request.form("Pick_List") <> "" then
	response.redirect "pick_list.asp?PickList=Y&Orders="&server.urlencode(request.form("DELETE_IDS"))
end if

if request.form="" and request.querystring("Form") <> "" then
	Form_Array = split(request.querystring("Form"),"^")
	col1 = Form_Array(0)
	col1_value = Form_Array(1)
	col1_oper = Form_Array(2)
	col2 = Form_Array(3)
	col2_value = Form_Array(4)
	ShippingStatus = Form_Array(5)
	Show = Form_Array(6)
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
	ShippingStatus = request.form("ShippingStatus")
	Show = request.form("Show")
end if

if request.querystring("view_cid") <> "" then
	col1 = "ccid=**"
	col1_value = request.querystring("view_cid")
	col1_oper = "AND"
	col2 = ""
	col2_value = ""
	ShippingStatus = 0
	Show = 0
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
					fn_error "The search has failed.  You have entered a non-numeric value ("&col2_value&") into a field that requires a number.")
				end if
			elseif instr(col2_new,"created")>0 or instr(col2_new,"date")>0 then
				'keywords means its a date
				if not isDate(col2_value) then
					fn_error "The search has failed.  You have entered a non-date value ("&col2_value&") into a field that requires a date.")
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

sNeedSubmit=1
set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_purchases"

if ShippingStatus="" then
	sql_where=sql_where & "ShippedPr='No' and "
elseif ShippingStatus=1 then
	sql_where=sql_where & "ShippedPr='Yes' and "
elseif ShippingStatus=2 then
	sql_where=sql_where & "ShippedPr='No' and "
elseif ShippingStatus="3" then
	sql_where=sql_where & "ShippedPr='Partial' and "
end if
if not isNumeric(Show) or Show="" then
	sql_where=sql_where & "verified<>0 and returned is Null and purchase_completed=1 and "
elseif Show=1 then
	sql_where=sql_where & "verified<>0 and returned is Null and purchase_completed=1 and "
elseif Show=2 then
	sql_where=sql_where & "verified=0 and returned is Null and purchase_completed=1 and "
elseif Show=3 then
	sql_where=sql_where & "returned<>0 and "
elseif Show=4 then
	sql_where=sql_where & "purchase_completed=0 and "
end if
	

End_Date1 = Dateadd("d",1,End_Date)
sSubmitName="SearchOrders"
myStructure("DeleteTable") = "store_purchases"
myStructure("ColumnList") = "oid,purchase_date,ccid,shipfirstname+' '+shiplastname as shipto,grand_total"
myStructure("HeaderList") = "oid,purchase_date,ccid,shipto,grand_total,details"
myStructure("SortList") = "oid,purchase_date,ccid,shipto,grand_total"
myStructure("TableWhere") = sql_where&" 1=1"
myStructure("DefaultSort") = "oid"
myStructure("PrimaryKey") = "oid"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 0
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "orders"
myStructure("FileName") = "orders.asp"
myStructure("FormAction") = "orders.asp"
myStructure("Title") = "Order Search"
myStructure("FullTitle") = "Orders > Search"
myStructure("CommonName") = "Order"
myStructure("NewRecord") = "order_edit.asp"
myStructure("Heading:oid") = "Order Id"
myStructure("Heading:details") = "View"
myStructure("Heading:purchase_date") = "Date"
myStructure("Heading:shipto") = "Ship To"
myStructure("Heading:ccid") = "Cust Id"
myStructure("Heading:payment_method") = "Payment Method"
myStructure("Heading:grand_total") = "Grand Total"
myStructure("Format:shipto") = "STRING"
myStructure("Format:ccid") = "STRING"
myStructure("Format:purchase_date") = "DATE"
myStructure("Format:oid") = "STRING"
myStructure("Format:payment_method") = "STRING"
myStructure("Format:grand_total") = "CURR"
myStructure("Format:details") = "TEXT"
myStructure("Link:oid") = "order_details.asp?Id=THISFIELD"
myStructure("Link:ccid") = "modify_customer.asp?Id=THISFIELD"
myStructure("Link:details") = "order_details.asp?Id=PK"
myStructure("Form") = col1&"^"&col1_value&"^"&col1_oper&"^"&col2&"^"&col2_value&"^"&ShippingStatus&"^"&Show
myStructure("Footer") = "<input class=buttons type='Button' name='DeleteAll' value='Print Selected Labels' onClick=""JavaScript:goPrintSeleted();"">Start at Label <input type=text name=Start_Cell value=1 size=2 onKeyPress=""return goodchars(event,'0123456789')""><BR><input class=buttons type='Button' name='DeleteAll' value='Print Selected Packing' onClick=""JavaScript:goPrintPackingSelected();"">&nbsp;<input class=buttons type='Button' name='DeleteAll' value='Pick List' onClick=""JavaScript:goPickList();"">&nbsp;<input class=buttons type='Button' name='DeleteAll' value='Ship Selected' onClick=""JavaScript:goShipSeleted();""><BR><input class=buttons type='Button' name='DeleteAll' value='Print Selected Invoices' onClick=""JavaScript:goPrintInvoicesSelected();"">&nbsp;<input class=buttons type='Button' name='DeleteAll' value='Print Selected Orders' onClick=""JavaScript:goPrintOrdersSeleted();"">"
 %>
<!--#include file="head_view.asp"-->
<%

if Request.Form("Export_To_Text_File") <> "" or Request.Form("Export_To_QuickBooks_File") <> "" then
	sql_select_orders="select * from store_purchases where store_id="&Store_Id & " and "&myStructure("TableWhere")&" order by "&Sort_By

   if Sort_Direction = 1 then
		sql_select_orders = sql_select_orders & " desc"
	end if
	'sql_select_orders=replace(sql_select_orders,"rpt_order.","rpt_order2.")

	File_Name = "orders_on_"&End_Date
	File_Name = Replace(File_Name,"/","-")
	QB_ACCOUNT="RECEIVABLE"
	QB_ACCOUNT2="INCOMES"
end if

if Request.Form("Export_To_Text_File") <> "" then
   on error resume next
	'CODE FOR SAVING THE RESULTS IN A TEXT FILE, FOR
	'USING IN AN EXTERNAL SOFTWARE
	Set FileObject = CreateObject("Scripting.FileSystemObject")

    Export_Folder = fn_get_sites_folder(Store_Id,"Export")
	File_Full_Name = Export_Folder&File_Name&".txt"
   If FileObject.FileExists(File_Full_Name) Then
      FileObject.DeleteFile(File_Full_Name)
   end if

	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)

	TxtExp="Verified" & chr(9) & "Purchase date" & chr(9)& "Order#" & chr(9)& "Customer#" & chr(9)&_
		"Shipping method" & chr(9)&"Shipped?" & chr(9) & "Payment method" & chr(9)& "Grand total" & chr(9)&_
		"Customer name" & chr(9)& "Customer phone" & chr(9)& "Customer email" & chr(9)& "Item_id"&chr(9)&_
		"Item_SKU"&chr(9)& "Item_Name"&chr(9)& "Additional"&chr(9)& "Quantity"&chr(9)& "Sale_Price"&chr(9)&_
		"Ship Last Name"&chr(9)& "Ship First Name"&chr(9)& "Ship Company"&chr(9)& "Ship Address1"&chr(9)&_
		"Ship Address2"&chr(9)& "Ship City"&chr(9)& "Ship State"&chr(9)& "Ship Zip"&chr(9)& "Ship Country"&chr(9)&_
		"Ship Phone"&chr(9)& "Ship Fax"&chr(9)& "Ship Email"&chr(9)& "Bill Last Name"&chr(9)& "Bill First Name"&chr(9)&_
		"Bill Company"&chr(9)& "Bill Address1"&chr(9)& "Bill Address2"&chr(9)& "Bill City"&chr(9)& "Bill State"&chr(9)&_
		"Bill Zip"&chr(9)& "Bill Country"&chr(9)& "Bill Phone"&chr(9)& "Bill Fax"&chr(9)& "Bill Email"&chr(9)&_
		"Customer Notes"& chr(13) & chr(10)

	set myfields=server.createobject("scripting.dictionary")
	Call DataGetrows(conn_store,sql_select_orders,mydata,myfields,noRecords)
  
	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			sql_select_order_details = "select * from Store_Customers where CID=" & mydata(myfields("cid"),rowcounter) & " and store_id="&store_id&" and record_type=0"
			set myfields2=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_select_order_details,mydata2,myfields2,noRecords2)
			rowcounter2 = 0


			sql_select_order_details = "select * from Store_Transactions where OID=" & mydata(myfields("oid"),rowcounter) & " and store_id="&store_id
			set myfields1=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_select_order_details,mydata1,myfields1,noRecords1)
			if mydata(myfields("verified"),rowcounter) = -1 then
				Verified = "Yes"
			Else
				Verified = "No"
			end if
			if mydata(myfields("shipstate"),rowcounter) = "AA" then
				shipstate = "Outside US & Canada"
			else
				shipstate = mydata(myfields("shipstate"),rowcounter)
			end if
			if mydata2(myfields2("state"),rowcounter2) = "AA" then
				state= "Outside US & Canada"
			else
				state= mydata2(myfields2("state"),rowcounter2)
			end if

			TxtExpOrderBasic=Verified & chr(9)& mydata(myfields("purchase_date"),rowcounter) & chr(9)&_
				mydata(myfields("oid"),rowcounter) & chr(9)& mydata(myfields("ccid"),rowcounter) & chr(9)&_
				mydata(myfields("shipping_method_name"),rowcounter) & chr(9)&_
				mydata(myfields("shippedpr"),rowcounter) & chr(9)&_
				mydata(myfields("payment_method"),rowcounter) & chr(9)&_
				Currency_Format_Function(mydata(myfields("grand_total"),rowcounter)) & chr(9)&_
				mydata2(myfields2("first_name"),rowcounter2) & " " & mydata2(myfields2("last_name"),rowcounter2) & chr(9)&_
				mydata2(myfields2("phone"),rowcounter2) & chr(9)&_
				mydata2(myfields2("email"),rowcounter2) & chr(9)
			TxtExpOrderShip=mydata(myfields("shiplastname"),rowcounter)&chr(9)&_
				mydata(myfields("shipfirstname"),rowcounter)&chr(9)&_
				mydata(myfields("shipcompany"),rowcounter) &chr(9)&_
				mydata(myfields("shipaddress1"),rowcounter)&chr(9)&_
				mydata(myfields("shipaddress2"),rowcounter) &chr(9)&_
				mydata(myfields("shipcity"),rowcounter)&chr(9)&_
				shipstate&chr(9)&_
				mydata(myfields("shipzip"),rowcounter)&chr(9)&_
				mydata(myfields("shipcountry"),rowcounter)&chr(9)&_
				mydata(myfields("shipphone"),rowcounter)&chr(9)&_
				mydata(myfields("shipfax"),rowcounter)&chr(9)&_
				mydata(myfields("shipemail"),rowcounter)&chr(9)
			TxtExpOrderShip=replace(TxtExpOrderShip,chr(9)&chr(9),chr(9)&"{}"&chr(9))
					
			'billing information
			TxtExpOrderBill=mydata2(myfields2("last_name"),rowcounter2)&chr(9)&_
				mydata2(myfields2("first_name"),rowcounter2)&chr(9)&_
				mydata2(myfields2("company"),rowcounter2)&chr(9)&_
				mydata2(myfields2("address1"),rowcounter2)&chr(9)&_
				mydata2(myfields2("address2"),rowcounter2)&chr(9)&_
				mydata2(myfields2("city"),rowcounter2)&chr(9)&_
				state&chr(9)&_
				mydata2(myfields2("zip"),rowcounter2)&chr(9)&_
				mydata2(myfields2("country"),rowcounter2)&chr(9)&_
				mydata2(myfields2("phone"),rowcounter2)&chr(9)&_
				mydata2(myfields2("fax"),rowcounter2)&chr(9)&_
				mydata2(myfields2("email"),rowcounter2)&chr(9)&_
				mydata(myfields("cust_notes"),rowcounter)
			TxtExpOrderBill=replace(TxtExpOrderBill,chr(9)&chr(9),chr(9)&"{}"&chr(9))
								
			if noRecords1 = 0 then
				FOR rowcounter1= 0 TO myfields1("rowcount")
					sAttrSku=mydata1(myfields1("item_attribute_skus"),rowcounter1)
					if isnull(sAttrSku) then
					   sAttrSku=""
					end if
                    if replace(replace(sAttrSku,",","")," ","")<>"" then
						mysku=sAttrSku 
					else
						mysku=mydata1(myfields1("item_sku"),rowcounter1)
					end if
					TxtExpOrderDetail=mydata1(myfields1("item_id"),rowcounter1)&chr(9)&_
						mysku&chr(9)&_
						mydata1(myfields1("item_name"),rowcounter1)&chr(9)&_
						mydata1(myfields1("u_d_1"),rowcounter1)&" "&mydata1(myfields1("u_d_2"),rowcounter1)&" "&mydata1(myfields1("u_d_3"),rowcounter1)&" "&mydata1(myfields1("u_d_4"),rowcounter1)&_
						" "&chr(9)& mydata1(myfields1("quantity"),rowcounter1)&chr(9)&_
						Currency_Format_Function(mydata1(myfields1("sale_price"),rowcounter1))&chr(9)
					TxtExpOrder=TxtExpOrderBasic&TxtExpOrderDetail&TxtExpOrderShip&TxtExpOrderBill
					TxtExpOrder=replace(TxtExpOrder,vbcrlf,"")
					TxtExpOrder=replace(TxtExpOrder,chr(13),"")
					TxtExpOrder=replace(TxtExpOrder,chr(10),"")
					TxtExpOrder=TxtExpOrder & chr(13) & chr(10)
					TxtExp=TxtExp&TxtExpOrder
					TxtExpOrder=""

				Next
			end if


				
				TxtExpOrderDetail="Taxes"&chr(9)& "Taxes"&chr(9)& "Taxes"&chr(9)& chr(9)& "1"&chr(9)&_
					Currency_Format_Function(mydata(myfields("tax"),rowcounter))&chr(9)
				TxtExpOrder=TxtExpOrderBasic&TxtExpOrderDetail&TxtExpOrderShip&TxtExpOrderBill		
				TxtExpOrder=replace(TxtExpOrder,vbcrlf,"")
				TxtExpOrder=replace(TxtExpOrder,chr(13),"")
				TxtExpOrder=replace(TxtExpOrder,chr(10),"")
				TxtExpOrder=TxtExpOrder & chr(13) & chr(10)
				TxtExp=TxtExp&TxtExpOrder
				TxtExpOrder=""


			TxtExpOrderDetail="Shipping"&chr(9)& "Shipping"&chr(9)& "Shipping"&chr(9)& chr(9)& "1"&chr(9)&_
					Currency_Format_Function(mydata(myfields("shipping_method_price"),rowcounter))&chr(9)
				TxtExpOrder=TxtExpOrderBasic&TxtExpOrderDetail&TxtExpOrderShip&TxtExpOrderBill		
				TxtExpOrder=replace(TxtExpOrder,vbcrlf,"")
				TxtExpOrder=replace(TxtExpOrder,chr(13),"")
				TxtExpOrder=replace(TxtExpOrder,chr(10),"")
				TxtExpOrder=TxtExpOrder & chr(13) & chr(10)
				TxtExp=TxtExp&TxtExpOrder
				TxtExpOrder=""
			 

			if mydata(myfields("coupon_amount"),rowcounter) > 0 and myfields1("rowcount") =>0 then
				TxtExpOrderDetail="Coupon"&chr(9)& mydata(myfields("coupon_id"),rowcounter)&chr(9)& mydata(myfields("coupon_id"),rowcounter)&chr(9)& chr(9)& "1"&chr(9)&_
					Currency_Format_Function(-mydata(myfields("coupon_amount"),rowcounter))&chr(9)
				TxtExpOrder=TxtExpOrderBasic&TxtExpOrderDetail&TxtExpOrderShip&TxtExpOrderBill		
				TxtExpOrder=replace(TxtExpOrder,vbcrlf,"")
				TxtExpOrder=replace(TxtExpOrder,chr(13),"")
				TxtExpOrder=replace(TxtExpOrder,chr(10),"")
				TxtExpOrder=TxtExpOrder & chr(13) & chr(10)
			end if
			TxtExp=TxtExp&TxtExpOrder
			TxtExpOrder=""
		Next
	end if
	TxtExp=replace(replace(TxtExp,"&#34;",""""),"&#39;","'")

	MyFile.Write TxtExp
	MyFile.Close

	Response.Redirect "Download_Exported_Files.asp"


end if

if Request.Form("Export_To_QuickBooks_File") <> "" then

	
	'for selected orders 21june2007
	
		selected_orders = Request.Form("DELETE_IDS")
	
	'completed

	'CODE FOR SAVING THE RESULTS IN A QUICKBOOKS FORMAT FILE, FOR
	'USING IN QUICKBOOKS SOFTWARE

	Set FileObject = CreateObject("Scripting.FileSystemObject")
	Set rs_Details = Server.CreateObject("ADODB.Recordset")
	
	Export_Folder = fn_get_sites_folder(Store_Id,"Export")
	File_Name = File_Name&".iif"
	File_Full_Name = Export_Folder&File_Name	

	TxtExp="!TRNS" & chr(9)& "TRNSID" & chr(9)& "TRANSTYPE" & chr(9)& "DATE" & chr(9)& "ACCNT" & chr(9)&_
		"NAME" & chr(9)& "CLASS" & chr(9)& "AMOUNT" & chr(9)& "DOCNUM" & chr(9)& "MEMO" & chr(9)&_
		"CLEAR" & chr(9)& "TOPRINT" & chr(9)& "NAMEISTAXABLE" & chr(9)& "ADDR1" & chr(9)& "ADDR3" & chr(9)&_
		"TERMS" & chr(9)& "SHIPVIA" & chr(9)& "SHIPDATE" & chr(9)& chr(13) & chr(10)&_
		"!SPL" & chr(9)& "SPLID" & chr(9)& "TRANSTYPE" & chr(9)& "DATE" & chr(9)& "ACCNT" & chr(9)&_
		"NAME" & chr(9)& "CLASS" & chr(9)& "AMOUNT" & chr(9)& "DOCNUM" & chr(9)& "MEMO" & chr(9)&_
		"CLEAR" & chr(9)& "QNTY" & chr(9)& "PRICE" & chr(9)& "INVITEM" & chr(9)& "TAXABLE" & chr(9)&_
		"OTHER2" & chr(9)& "YEARTODATE" & chr(9)& "WAGEBASE" & chr(9)& chr(13) & chr(10)&_
		"!ENDTRNS" & chr(9)& chr(13) & chr(10)

	Set MyFile = FileObject.OpenTextFile(File_Full_Name, 8,true)
	set myfields=server.createobject("scripting.dictionary")

	Call DataGetrows(conn_store,sql_select_orders,mydata,myfields,noRecords)

	if noRecords = 0 then
		FOR rowcounter= 0 TO myfields("rowcount")
			If Response.IsClientConnected=false Then
				exit for
			end if

			sql_select_order_details = "select * from Store_Transactions where OID=" & mydata(myfields("oid"),rowcounter) & " and store_id ="&store_id
			set myfields1=server.createobject("scripting.dictionary")
			Call DataGetrows(conn_store,sql_select_order_details,mydata1,myfields1,noRecords1)

			if noRecords1=0 then
				sSerialDate = dateserial(year(mydata(myfields("sys_created"),rowcounter)),month(mydata(myfields("sys_created"),rowcounter)),day(mydata(myfields("sys_created"),rowcounter)))
				TxtExpOrder="TRNS" & chr(9)& mydata(myfields("oid"),rowcounter) & chr(9)& "INVOICE" & chr(9)&_
   					sSerialDate & chr(9)& QB_ACCOUNT & chr(9)& mydata(myfields("shipfirstname"),rowcounter) &_
                   " " & mydata(myfields("shiplastname"),rowcounter)& chr(9)&_
   					chr(9)& formatnumber(mydata(myfields("grand_total"),rowcounter),2) & chr(9)&_
   					mydata(myfields("oid"),rowcounter) & chr(9)& chr(9)& "N" & chr(9)& "Y" & chr(9)&"N" & chr(9)&_
					mydata(myfields("shipaddress1"),rowcounter) & " " & mydata(myfields("shipcity"),rowcounter) & " " & mydata(myfields("shipstate"),rowcounter) & " " & mydata(myfields("shipzip"),rowcounter) & " " & mydata(myfields("shipcountry"),rowcounter)& chr(9)&_
   					chr(9)& chr(9)& chr(9)& chr(9)& chr(13) & chr(10)

   				thistotal = cdbl(formatnumber(mydata(myfields("grand_total"),rowcounter),2))
				FOR rowcounter1= 0 TO myfields1("rowcount")
					itemtotal = cdbl(formatnumber(-cdbl(mydata1(myfields1("quantity"),rowcounter1)*mydata1(myfields1("sale_price"),rowcounter1)),2))
					thistotal=thistotal+itemtotal
					sAttrSku=mydata1(myfields1("item_attribute_skus"),rowcounter1)
					if isnull(sAttrSku) then
					   sAttrSku=""
					end if
               if replace(replace(sAttrSku,",","")," ","")<>"" then
						mysku=sAttrSku
						  
					else
						mysku=mydata1(myfields1("item_sku"),rowcounter1)
					end if

					TxtExpOrder=TxtExpOrder & ("SPL" & chr(9)& mydata1(myfields1("transaction_id"),rowcounter1) & chr(9)&_
						"INVOICE" & chr(9)& sDateSerial & chr(9)& QB_ACCOUNT2 & chr(9)& chr(9)& chr(9)&_
						formatnumber(-cdbl(mydata1(myfields1("quantity"),rowcounter1)*mydata1(myfields1("sale_price"),rowcounter1)),2) & chr(9)&_
						mydata(myfields("oid"),rowcounter) & chr(9)& mydata1(myfields1("item_name"),rowcounter1) & chr(9)&_
						"N" & chr(9)& formatnumber(-cdbl(mydata1(myfields1("quantity"),rowcounter1)),2) & chr(9)&_
						formatnumber(mydata1(myfields1("sale_price"),rowcounter1),2) & chr(9)&_
						left(replace(replace(mysku,"<BR>",""),"=",""),25) & chr(9)& "Y" & chr(9)& chr(9)& "0" & chr(9)&_
						"0" & chr(9)&  chr(13)&chr(10))
				Next
			end if

			if mydata(myfields("tax"),rowcounter) > 0 and noRecords1=0 then
				itemtotal = cdbl(formatnumber(mydata(myfields("tax"),rowcounter),2))
				thistotal=thistotal-itemtotal
				TxtExpOrder=TxtExpOrder & ("SPL" & chr(9)& mydata(myfields("oid"),rowcounter)+10000 & chr(9)&_
					"INVOICE" & chr(9)& sDateSerial & chr(9)& "TAX" & chr(9)& chr(9)& chr(9)&_
					formatnumber(-1*mydata(myfields("tax"),rowcounter),2) & chr(9)&_
					mydata(myfields("oid"),rowcounter) & chr(9)& "Tax" & chr(9)& "N" & chr(9)&_
					formatnumber(-1,2) & chr(9)& formatnumber(mydata(myfields("tax"),rowcounter),2) & chr(9)&_
					"Tax" & chr(9)& "N" & chr(9)& "" & chr(9)& "0" & chr(9)& "0" & chr(9)&  chr(13)&chr(10))
			end if

			if mydata(myfields("shipping_method_price"),rowcounter) > 0 and noRecords1=0 then
				itemtotal = formatnumber(mydata(myfields("shipping_method_price"),rowcounter),2)
				thistotal=thistotal-itemtotal
				TxtExpOrder=TxtExpOrder & ("SPL" & chr(9)& mydata(myfields("oid"),rowcounter)+100000 & chr(9)&_ 
					"INVOICE" & chr(9)& sDateSerial & chr(9)& QB_ACCOUNT2 & chr(9)& chr(9)& chr(9) &_
					formatnumber(-1*mydata(myfields("shipping_method_price"),rowcounter),2) & chr(9)&_ 
					mydata(myfields("oid"),rowcounter) & chr(9)& "Shipping" & chr(9)& "N" & chr(9)&_
					formatnumber(-1,2) & chr(9)&_
					formatnumber(mydata(myfields("shipping_method_price"),rowcounter),2) & chr(9)&_
					"Shipping" & chr(9)& "N" & chr(9)& "" & chr(9)& "0" & chr(9)& "0" & chr(9)&  chr(13)&chr(10))
			end if

			if mydata(myfields("coupon_amount"),rowcounter) > 0 and noRecords1=0 then
				itemtotal = cdbl(formatnumber(mydata(myfields("coupon_amount"),rowcounter),2))
				thistotal=thistotal+itemtotal
				TxtExpOrder=TxtExpOrder & ("SPL" & chr(9)& mydata(myfields("oid"),rowcounter)+1000000 & chr(9)&_
					"INVOICE" & chr(9)& sDateSerial & chr(9)& QB_ACCOUNT2 & chr(9)& chr(9)& chr(9)&_
					formatnumber(mydata(myfields("coupon_amount"),rowcounter),2) & chr(9)&_
					mydata(myfields("oid"),rowcounter) & chr(9)& "Coupon" & chr(9)& "N" & chr(9)&_
					formatnumber(1,2) & chr(9)&_
					formatnumber(mydata(myfields("coupon_amount"),rowcounter),2) & chr(9)&_
					"Coupon" & chr(9)& "Y" & chr(9)& "" & chr(9)& "0" & chr(9)& "0" & chr(9)&  chr(13)&chr(10))
			end if

			if mydata(myfields("promotion_amount"),rowcounter) > 0 and noRecords1=0 then
				itemtotal = cdbl(formatnumber(mydata(myfields("promotion_amount"),rowcounter),2))
				thistotal=thistotal+itemtotal
				TxtExpOrder=TxtExpOrder & ("SPL" & chr(9)& mydata(myfields("oid"),rowcounter)+1000000 & chr(9)&_
					"INVOICE" & chr(9)& sDateSerial & chr(9)& QB_ACCOUNT2 & chr(9)& chr(9)& chr(9)&_
					formatnumber(mydata(myfields("promotion_amount"),rowcounter),2) & chr(9)&_
					mydata(myfields("oid"),rowcounter) & chr(9)& "Promotion" & chr(9)& "N" & chr(9)&_ 
					formatnumber(1,2) & chr(9)& formatnumber(mydata(myfields("promotion_amount"),rowcounter),2) & chr(9)&_
					"Promotion" & chr(9)& "Y" & chr(9)& chr(9)& "0" & chr(9)& "0" & chr(9)&  chr(13)&chr(10))
			end if

			if noRecords1=0 and thistotal<>0 then
				TxtExpOrder=TxtExpOrder & ("SPL" & chr(9)& mydata(myfields("oid"),rowcounter)+10000000 & chr(9)&_
					"INVOICE" & chr(9)& sDateSerial & chr(9)& QB_ACCOUNT2 & chr(9)& chr(9)& chr(9)&_
					-formatnumber(thistotal,2) & chr(9)& mydata(myfields("oid"),rowcounter) & chr(9)&_
					"Rounding" & chr(9)& "N" & chr(9)& formatnumber(1,2) & chr(9)& -formatnumber(thistotal,2) & chr(9)&_
					"Rounding" & chr(9)& "Y" & chr(9)& "" & chr(9)& "0" & chr(9)& "0" & chr(9)&  chr(13)&chr(10))
			end if

			if noRecords1=0 then
				TxtExpOrder=TxtExpOrder & ("ENDTRNS" & chr(9)& chr(13) & chr(10))
			end if
			TxtExp=TxtExp&TxtExpOrder
			TxtExpOrder=""
		Next
	End If
	TxtExp=replace(replace(TxtExp,"&#34;",""""),"&#39;","'")
   TxtExp=replace(replace(TxtExp,"&nbsp;"," "),"<BR>","")

	MyFile.Write TxtExp
	MyFile.Close


end if

%>
<script langauge="JavaScript">

	function goPrintSeleted(){
		tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to use batch processing.");
                }
                         else
                         {
		document.forms[0].Ship_Labels.value="SEL";
		document.forms[0].Print_Invoices.value="";
      document.forms[0].Ship.value="";
		document.forms[0].Delete_Item.value="";
		document.forms[0].Print_Orders.value="";
		document.forms[0].Print_Packing.value="";
		document.forms[0].Pick_List.value="";
			document.forms[0].submit(); }
		}
	function goPickList(){
		tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to use batch processing.");
                }
                else
                         {
		document.forms[0].Pick_List.value="SEL";
		document.forms[0].Print_Invoices.value="";
      document.forms[0].Ship_Labels.value="";
		document.forms[0].Print_Orders.value="";
		document.forms[0].Print_Packing.value="";
		document.forms[0].Ship.value="";
		document.forms[0].submit(); }
		}
					
	function goShipSeleted(){
		tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to use batch processing.");
                }
                         else
                         {
		document.forms[0].Ship.value="SEL";
		document.forms[0].Print_Invoices.value="";
                document.forms[0].Delete_Item.value="";
		document.forms[0].Ship_Labels.value="";
		document.forms[0].Print_Orders.value="";
		document.forms[0].Print_Packing.value="";
		document.forms[0].Pick_List.value="";
			document.forms[0].submit(); }
                       }

	function goPrintOrdersSeleted()
	{
                tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to use batch processing.");
                }
		else
                {
		document.forms[0].Print_Invoices.value="";
     	document.forms[0].Print_Orders.value="SEL";
		document.forms[0].Print_Packing.value="";
		document.forms[0].Ship.value="";
		document.forms[0].Delete_Item.value="";
		document.forms[0].Ship_Labels.value="";
		document.forms[0].Pick_List.value="";
		document.forms[0].submit();
		}
	}
	function goPrintPackingSelected()
	{
                tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to use batch processing.");
                }
		else
                {
		document.forms[0].Print_Invoices.value="";
          document.forms[0].Print_Packing.value="SEL";
		document.forms[0].Print_Orders.value="";
		document.forms[0].Ship.value="";
		document.forms[0].Delete_Item.value="";
		document.forms[0].Ship_Labels.value="";
		document.forms[0].Pick_List.value="";
		document.forms[0].submit();
		}
	}
	function goPrintInvoicesSelected()
	{
	        tsel=0;
                var mylength=document.forms[0].DELETE_IDS.length;
                if(typeof mylength=="undefined")
                {
                // no array exists, examine 'checked' property of checkbox directly
                if(document.forms[0].DELETE_IDS.checked==true) tsel++;
                }
                else
                {
                // an array has been defined, step through each array element
                for(i=0;i<mylength;i++)
                {
                if (document.forms[0].DELETE_IDS[i].checked) tsel++;
                }
                }

// check to see if any checkboxes have been selected
                if(!tsel)
                {
                 alert("You must select at least one order to use batch processing.");
                }
                         else
                         {
		document.forms[0].Print_Invoices.value="SEL";
		document.forms[0].Print_Orders.value="";
		document.forms[0].Print_Packing.value="";
		document.forms[0].Ship.value="";
		document.forms[0].Delete_Item.value="";
		document.forms[0].Ship_Labels.value="";
		document.forms[0].Pick_List.value="";
		document.forms[0].submit(); 
		}
	}



</script>
<input type=hidden name=records value=<%= RowsPerPage %>>
<input type="Hidden" name="Delete_Item" value="">
	<input type="Hidden" name="Ship_Labels" value="">
	<input type="Hidden" name="Ship" value="">
	<input type="Hidden" name="Print_Orders" value="">
	<input type="Hidden" name="Print_Packing" value="">
	<input type="Hidden" name="Print_Invoices" value="">
	<input type="Hidden" name="Pick_List" value="">

        <SCRIPT LANGUAGE="VBScript">

	sub display()

	

		Dim WshShell
		dim FValue
		dim SValue		
		dim temp
		dim ShowVal
		dim VeryVal		
		
		Set WshShell = CreateObject("wscript.shell") 



		temp= "<%=sql_select_orders%>"

'msgbox temp
'exit sub


		if temp="" then
			msgbox "Please click Export to Quickbooks button First",0,"QuickBooks"
			exit Sub
		end if

			
		temp=temp & "|" & "TABLE=PURCHASES" & "|" & "<%=store_id%>"& "|" &"<%=selected_orders%>"
		


		
		WshShell.Run "C:\bin\Quick.exe" & " " & temp





exit sub





		
		fvalue= "<%=col1_value%>"
		svalue="<%=col2_value%>"
		ShowVal= "<%=ShippingStatus%>"
		VeryVal="<%=Show%>"
		
		
		
		
		'Both Text boxes are empty then do nothing
		If fvalue="" and svalue="" then	
			if ShowVal="0" then
				if VeryVal=0 then
					
				else
					temp="WHERE " & "Verified=" & "<%=show%>"
				end if				
			ElseIf ShowVal="1" then
				ShowVal="Yes"
				if VeryVal=0 then
					temp="WHERE " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					temp="WHERE " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			Elseif ShowVal="2" then
				ShowVal="No"
				if VeryVal=0 then
					temp="WHERE " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					temp="WHERE " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			else
				ShowVal="Partial"
				if VeryVal=0 then
					temp="WHERE " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
				temp="WHERE " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			end if	
		
		'First text box not empty then pass it as argument with Where clause
		ElseIf FValue<>"" and SValue="" then
			if ShowVal=0 then
				if VeryVal=0 then
					Temp="WHERE " & "<%=col1_new%>"
				else
					Temp="WHERE " & "<%=col1_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" 
				end if				
			ElseIf ShowVal=1 then
				ShowVal="Yes"
				if VeryVal=0 then
					Temp="WHERE " & "<%=col1_new%>" & " " & "AND " &  " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					Temp="WHERE " & "<%=col1_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			Elseif ShowVal=2 then
				ShowVal="No"
				if VeryVal=0 then
					Temp="WHERE " & "<%=col1_new%>" & " " &  "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					Temp="WHERE " & "<%=col1_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if
				
			else
				ShowVal="Partial"
				if VeryVal=0 then
					Temp="WHERE " & "<%=col1_new%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"			
				else
					Temp="WHERE " & "<%=col1_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			end if		
		
			
		'Second text box not empty then pass it as argument with Where clause	
		ElseIf Fvalue="" and SValue<>"" then
			if ShowVal=0 then
				if VeryVal=0 then
					Temp="WHERE " & "<%=col2_new%>"
				else
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" 
				end if				
			ElseIf ShowVal=1 then
				ShowVal="Yes"
				if VeryVal=0 then
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " "  & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			Elseif ShowVal=2 then
				ShowVal="No"
				if VeryVal=0 then
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " "  & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"	
				end if				
			else
				ShowVal="Partial"
				if VeryVal=0 then
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				else
					Temp="WHERE " & "<%=col2_new%>" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " " & "'" & ShowVal & "'"
				end if				
			end if		
			
		Else
		'both text box not empty then pass them as argument with Where clause	
			if ShowVal=0 then
				if VeryVal=0 then
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)"
				else
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "Verified=" & "<%=show%>"
				end if				
			ElseIf ShowVal=1 then
				ShowVal="Yes"
				if VeryVal=0 then
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "ShippedPr LIKE " & " "  & "'" & ShowVal & "'"
				else
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " "  & "'" & ShowVal & "'"
				end if				
			Elseif ShowVal=2 then
				ShowVal="No"
				if VeryVal=0 then
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "ShippedPr LIKE " & " "  & "'" & ShowVal & "'" 			
				else
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " "  & "'" & ShowVal & "'" 				
				end if				
			else
				ShowVal="Partial"
				if VeryVal=0 then
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "ShippedPr LIKE " & " "  & "'" & ShowVal & "'"			
				else
					temp="WHERE " & "(<%=col1_new%>" &  " " & "<%=col1_oper%>" & " " & "<%=col2_new%>)" & " " & "AND " & " " & "Verified=" & "<%=show%>" & " " & "AND " & " " & "ShippedPr LIKE " & " "  & "'" & ShowVal & "'"			
				end if
				
			end if		
			
		end if


			
temp=temp & "|" & "TABLE=ITEMS"
		
msgbox temp
exit sub		
		
	WshShell.Run "C:\bin\Quick.exe" & " " & temp

end sub
		
</script>


<TR bgcolor='#D4DEE5'><td width='100%' colspan='3' height='13'><table width='100%' border='0' cellspacing='1' cellpadding=2 class='list'>
<tr bgcolor='#FFFFFF'>
		<td class="inputname" width="40%">Verified</td>
		<td class="inputvalue" width=60%">
			<select name="Show" size="1" ID="Select1">
					<option value="0"
					<% If Show="0" then %>
						selected
					<% End If %>
					>All</option>
					<option value="1"
					<% If Show="" or Show="1" then %>
						selected
					<% End If %>
					>Yes</option>
				<option value="2"
					<% If Show="2" then %>
						selected
					<% End If %>
					>No</option>
				<option value="3"
					<% If Show="3" then %>
						selected
					<% End If %>
					>Returned</option>
				<option value="4"
					<% If Show="4" then %>
						selected
					<% End If %>
					>Abandoned</option>
			</select>
			 </td></tr>
			 <tr bgcolor='#FFFFFF'><td class="inputname">Ship Status</td>
			<td class="inputvalue">
			<select name="ShippingStatus" size="1" ID="Select2">
					<option value="0"
					<% If ShippingStatus="0" then %>
						selected
					<% End If %>
					>All</option>
					<option value="1"
					<% If ShippingStatus="1" then %>
						selected
					<% End If %>
					>Yes</option>
					<option value="2"
					<% If ShippingStatus="" or ShippingStatus="2" then %>
						selected
					<% End If %>
					>No</option>
					<option value="3"
					<% If ShippingStatus="3" then %>
						selected
					<% End If %>
					>Partial</option> 		
			</select>
			 </td>
	</tr>
			<TR bgcolor='#FFFFFF'><td class=inputname width=40%>
		<select name=col1 ID="Select3">
		<% if col1<>"" then %>
			<option value="<%= col1 %>"><%= replace(replace(replace(replace(replace(col1,"**",""),"%%",""),"<>0","="),"1=1",""),"''","") %></option>
			<option value="">Select column to search on</option>
		<% else %>
			<option value="">Select column to search on</option>
		<% end if %>
		<option value="shipfirstname like '%**%'">First Name like</option>
		<option value="shiplastname like '%**%'">Last Name like</option>
		<option value="purchase_date > '**'">Purchase After (Date)</option>
		<option value="purchase_date < '**'">Purchase Before (Date)</option>
		<option value="shipcompany like '%**%'">Company like</option>
		<option value="shipaddress1 like '%**%'">Address like</option>
		<option value="shipCity like '%**%'">City like</option>
		<option value="shipState like '%**%'">State Abbrev like</option>
		<option value="shipZip like '%**%'">Zip Code like</option>
		<option value="shipCountry like '%**%'">Country like</option>
		<option value="shipPhone like '%**%'">Phone like</option>
		<option value="shipFax like '%**%'">Fax like</option>
		<option value="shipEmail like '%**%'">Email like</option>
		<option value="Tax >= **">Taxes >= (Number)</option>
		<option value="Tax <= **">Taxes <= (Number)</option>
		<option value="Shipping_Method_Price >= **">Shipping >= (Number)</option>
		<option value="Shipping_Method_Price <= **">Shipping <= (Number)</option>
		<option value="Grand_total >= **">Total >= (Number)</option>
		<option value="Grand_total <= **">Total <= (Number)</option>
		<option value="coupon_id like '%**%'">Coupon id like</option>
		<option value="payment_method like '%**%'">Payment Method like</option>
		<option value="shipping_method_name like '%**%'">Shipping Method like</option>
		<option value="verified_ref like '%**%'">Verification Id like</option>
		<option value="oid = **">Order ID = (Number)</option>
		<option value="ccid = **">Customer ID = (Number)</option>
		</td><td class=inputvalue><input type=text name=col1_value value="<%= col1_value %>" ID="Text1" size=49>
		<select name=col1_oper ID="Select4">
		<option value=AND>AND</option>
		<option value=OR>OR</option>
		</select></td></tr>
		<TR bgcolor='#FFFFFF'><td class=inputname width=40%>
		<select name=col2 ID="Select5">
        <% if col2<>"" then %>
			<option value="<%= col2 %>"><%= replace(replace(replace(replace(replace(col2,"**",""),"%%",""),"<>0","="),"1=1",""),"''","") %></option>
			<option value="">Select column to search on</option>
        <% else %>
			<option value="">Select column to search on</option>
        <% end if %>
        <option value="shipfirstname like '%**%'">First Name like</option>
		<option value="shiplastname like '%**%'">Last Name like</option>
		<option value="sys_created > '**'">Purchase After (Date)</option>
		<option value="sys_created < '**'">Purchase Before (Date)</option>
		<option value="shipcompany like '%**%'">Company like</option>
		<option value="shipaddress1 like '%**%'">Address like</option>
		<option value="shipCity like '%**%'">City like</option>
		<option value="shipState like '%**%'">State Abbrev like</option>
		<option value="shipZip like '%**%'">Zip Code like</option>
		<option value="shipCountry like '%**%'">Country like</option>
		<option value="shipPhone like '%**%'">Phone like</option>
		<option value="shipFax like '%**%'">Fax like</option>
		<option value="shipEmail like '%**%'">Email like</option>
		<option value="Tax >= **">Taxes >= (Number)</option>
		<option value="Tax <= **">Taxes <= (Number)</option>
		<option value="Shipping_Method_Price >= **">Shipping >= (Number)</option>
		<option value="Shipping_Method_Price <= **">Shipping <= (Number)</option>
		<option value="Grand_total >= **">Total >= (Number)</option>
		<option value="Grand_total <= **">Total <= (Number)</option>
		<option value="coupon_id like '%**%'">Coupon id like</option>
		<option value="payment_method like '%**%'">Payment Method like</option>
		<option value="shipping_method_name like '%**%'">Shipping Method like</option>
		<option value="verified_ref like '%**%'">Verification Id like</option>
		<option value="oid = **">Order ID = (Number)</option>
		<option value="ccid = **">Customer ID = (Number)</option>

			</td><td class=inputvalue><input type=text name=col2_value value="<%= col2_value %>" ID="Text2" size=60></td></tr>

			<tr bgcolor='#FFFFFF'>
					<td colspan="7" align="center">
							<input class="Buttons" name="SearchOrders" type="submit" value="Search Orders" ID="Submit1">
							<% if Service_Type>=7 then %>
								<BR><input class="Buttons" name="Export_To_Text_File" type="submit" value="Export To Text" ID="Submit2">
								<input class="Buttons" name="Export_To_QuickBooks_File" type="submit" value="Export To Quickbooks" ID="Submit3">
							<% end if %>
				</td>
	</tr>

<% if Request.Form("Export_To_QuickBooks_File") <> "" then %>
		<TR bgcolor='#FFFFFF'><TD colspan="7" width="35%" align="center">Please read the <a href=quickbooks.asp class=link>Quickbooks instructions</a> and then Click on <b>Insert in QuickBook</b> button to Import data!!
                <BR><input class="Buttons" name="Insert In to Quickbooks" type="button"  value="Insert in Quick Book" onclick ="Display()" ID="Button1"></td></tr>
<% end if %> 


		</table></td></tr>






<% if (request.querystring<>"" or request.form<>"") AND Request.Form("Export_To_QuickBooks_File")= "" then %>
<!--#include file="list_view.asp"-->
<%
end if

if Request.QueryString("Delete_Id") <> "" then
	Delete_Id = request.querystring("Delete_Id")
elseif Request.Form("Delete_Id") ="SEL" and request.form("DELETE_IDS") <> "" then
	Delete_Id = request.form("DELETE_IDS")
else
	Delete_Id = ""
end if

if Delete_Id<>"" then

	sql_delete = "delete from store_purchases where store_Id="&Store_Id&" and Oid in ("&Delete_Id&");"&_
		"delete from store_transactions where store_Id="&Store_Id&" and Oid in ("&Delete_Id&");"&_
		"delete from store_purchases_returns where store_Id="&Store_Id&" and Oid in ("&Delete_Id&");"&_
		"delete from store_gift_certificates_dets where store_Id="&Store_Id&" and Order_id in ("&Delete_Id&");"&_
        "delete from store_purchases_shippments where store_Id="&Store_Id&" and Oid in ("&Delete_Id&"); "&_
		"delete from store_maxmind where store_Id="&Store_Id&" and Oid in ("&Delete_Id&"); "&_
		"delete from store_purchases_notes where store_Id="&Store_Id&" and Order_id in ("&Delete_Id&")"
		
	conn_store.Execute sql_delete
end if

createFoot thisRedirect, 2
%>

