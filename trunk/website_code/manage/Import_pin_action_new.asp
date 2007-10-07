<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sImportType="Pin"
sHeader="pin"
sMenu="inventory"
sTitle="Import Inventory Pins" 
sFullTitle = "Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=edit_pin1.asp class=white>Pins</a> > Import"
sMenu="inventory"

Dim arrColumns(2)
arrColumns(0) = "Item_Sku:Re:Text:50"
arrColumns(1) = "Pin_No:Re:Text:20:"
arrColumns(2) = "Other_Info:Op:Text:100"

sProcName = "wsp_pin_import"

%>
<!--#include file="include/import_action_include_new.asp"-->
