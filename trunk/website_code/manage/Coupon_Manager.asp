<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "marketing/coupons.htm"
sInstructions="Coupon codes can be sent to your customers.  The customer will enter the unique coupon code at checkout to receive the discount you have specified."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Coupons"
myStructure("ColumnList") = "coupon_line_id,coupon_name,coupon_start_date,coupon_end_date,coupon_id,coupon_type,coupon_amount"
myStructure("HeaderList") = "coupon_name,coupon_start_date,coupon_end_date,coupon_id,coupon_type,coupon_amount"
myStructure("DefaultSort") = "coupon_name"
myStructure("PrimaryKey") = "coupon_line_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "coupon_manager.asp"
myStructure("FormAction") = "coupon_manager.asp"
myStructure("Title") = "Coupons"
myStructure("FullTitle") = "Marketing > Coupons"
myStructure("CommonName") = "Coupon"
myStructure("NewRecord") = "new_coupon.asp"
myStructure("Heading:coupon_line_id") = "PK"
myStructure("Heading:coupon_name") = "Name"
myStructure("Heading:coupon_start_date") = "Start"
myStructure("Heading:coupon_end_date") = "End"
myStructure("Heading:coupon_id") = "ID"
myStructure("Heading:coupon_type") = "Type"
myStructure("Heading:coupon_amount") = "Amount"
myStructure("Heading:coupon_total_left") = "Left"
myStructure("Format:coupon_name") = "STRING"
myStructure("Format:coupon_start_date") = "DATESHORT"
myStructure("Format:coupon_end_date") = "DATESHORT"
myStructure("Format:coupon_id") = "STRING"
myStructure("Format:coupon_type") = "LOOKUP"
myStructure("Format:coupon_amount") = "INT"
myStructure("Format:coupon_total_left") = "INT"
myStructure("Length:coupon_amount") = 2
myStructure("Length:coupon_total_left") = 0
myStructure("Lookup:coupon_type") = "0:%^-1:"&Store_Currency
%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
		Coupon_Line_Id = Request.QueryString("Delete_Id")
		' IF THERE ARE NO OTHER COUPONS IN THE DATABASE
		sql_check = "select coupon_id from store_Coupons where Store_id = "&Store_id
		rs_store.open sql_check,conn_store
		if rs_store.bof = true then
			' UPDATING STORE_SETTINGS - SET ENABLE_COUPON=FALSE
			sql_update = "update store_settings set Enable_Coupon = 0 where Store_id = "&Store_id
			conn_store.Execute sql_update
		end if
	end if
	
	if Request.Form("Delete_Id") ="SEL" then
  	if request.form("DELETE_IDS") <> "" then
      Coupon_Line_Id = request.form("DELETE_IDS")
      ' IF THERE ARE NO OTHER COUPONS IN THE DATABASE
  		sql_check = "select coupon_id from store_Coupons where Store_id = "&Store_id
  		rs_store.open sql_check,conn_store
  		if rs_store.bof = true then
  			' UPDATING STORE_SETTINGS - SET ENABLE_COUPON=FALSE
  			sql_update = "update store_settings set Enable_Coupon = 0 where Store_id = "&Store_id
  			conn_store.Execute sql_update
  		end if
	  end if
  end if
createFoot thisRedirect, 0
%>

