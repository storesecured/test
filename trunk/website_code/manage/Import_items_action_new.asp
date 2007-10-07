<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

sImportType="Items"
sTitle = "Import Items"
sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > Import"
sHeader="item"
sMenu="inventory"
sEmptytoNull=1

Dim arrColumns(52)
arrColumns(0) = "Item_Sku:Re:Text:200"
arrColumns(1) = "Department_Name:Op:Text:3000"
arrColumns(2) = "Item_Name:Op:Text:255"
arrColumns(3) = "Retail_Price:Op:Integer"
arrColumns(4) = "Wholesale_Price:Op:Integer"
arrColumns(5) = "ImageS_Path:Op:Text:100"
arrColumns(6) = "ImageL_Path:Op:Text:100"
arrColumns(7) = "Taxable:Op:Boolean"
arrColumns(8) = "Description_S:Op:Html:4000"
arrColumns(9) = "Description_L:Op:Html:8000"
arrColumns(10) = "Item_Weight:Op:Integer"
arrColumns(11) = "Quantity_in_stock:Op:Integer"
arrColumns(12) = "Quantity_Control:Op:Boolean"
arrColumns(13) = "Quantity_Control_Number:Op:Integer"
arrColumns(14) = "Retail_Price_special_Discount:Op:Integer"
arrColumns(15) = "Special_start_date:Op:Date"
arrColumns(16) = "Special_end_date:Op:Date"
arrColumns(17) = "Special_Description:Op:Text:100"
arrColumns(18) = "Shipping_Fee:Op:Integer:"
arrColumns(19) = "File_Location:Op:Text:250"
arrColumns(20) = "Item_Remarks:Op:Html:4000"
arrColumns(21) = "Fractional:Op:Boolean"
arrColumns(22) = "Show:Op:Boolean"
arrColumns(23) = "Filename:Op:Text:100::@,$,%, ,',&,.,/,(,),`,;,#,!,?,^"
arrColumns(24) = "Item_Attributes:Op:Text:4000"
arrColumns(25) = "View_Order:Op:Integer"
arrColumns(26) = "Meta_Keywords:Op:Text:1500"
arrColumns(27) = "Meta_Description:Op:Text:1500"
arrColumns(28) = "Item_Accessories:Op:Text:4000"
arrColumns(29) = "Meta_Title:Op:Text:100"
arrColumns(30) = "Ship_Location:Op:Html:3000"
arrColumns(31) = "M_d_1:Op:Html:4000"
arrColumns(32) = "M_d_2:Op:Html:4000"
arrColumns(33) = "M_d_3:Op:Html:4000"
arrColumns(34) = "M_d_4:Op:Html:4000"
arrColumns(35) = "M_d_5:Op:Html:4000"
arrColumns(36) = "Custom_Link:Op:Text:255"
arrColumns(37) = "U_d_1_name:Op:Text:250"
arrColumns(38) = "U_d_1:Op:Boolean"
arrColumns(39) = "U_d_2_name:Op:Text:250"
arrColumns(40) = "U_d_2:Op:Boolean"
arrColumns(41) = "U_d_3_name:Op:Text:250"
arrColumns(42) = "U_d_3:Op:Boolean"
arrColumns(43) = "U_d_4_name:Op:Text:250"
arrColumns(44) = "U_d_4:Op:Boolean"
arrColumns(45) = "U_d_5_name:Op:Text:250"
arrColumns(46) = "U_d_5:Op:Boolean"
arrColumns(47) = "sHandling_fee:Op:Integer"
arrColumns(48) = "Brand:Op:Text:100"
arrColumns(49) = "Condition:Op:Text:20"
arrColumns(50) = "Product_Type:Op:Text:20"
arrColumns(51) = "Item_Matrix:Op:Text:4000"
arrColumns(52) = "Item_Group:Op:Text:4000"



sProcName = "wsp_item_import"

%>
<!--#include file="include/import_action_include_new.asp"-->

	
	

	

