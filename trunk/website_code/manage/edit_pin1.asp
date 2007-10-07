<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sql_select_pin="SELECT item_sku,pin,other_info,pin_used,cid,use_date,oid FROM Store_pin WHERE Store_id="&Store_id
set deptfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select_pin,deptdata,deptfields,noRecordsDept)

sQuestion_Path = "inventory/edit_pins.htm"
sInstructions="A pin is a unique code that can be delivered to your customers upon successfull purchase.  Ie if you sold calling cards a pin number could be delivered  for using the calling card.  Or if you sold software a registration key could be delivered."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "store_pin"
myStructure("ColumnList") = "itemdown_id,item_sku,pin,other_info,pin_used"
myStructure("HeaderList") = "itemdown_id,item_sku,pin,other_info,pin_used"
myStructure("DefaultSort") = "item_sku"
myStructure("PrimaryKey") = "itemdown_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "inventory"
myStructure("FileName") = "edit_pin1.asp"
myStructure("FormAction") = "edit_pin1.asp"
myStructure("Title") = "Pins"
myStructure("FullTitle") = "Inventory > <a href=edit_items.asp class=white>Items</a> > Pins"
myStructure("CommonName") = "Pin"
myStructure("NewRecord") = "pin_add.asp"

myStructure("Heading:itemdown_id") = "PK"
myStructure("Heading:pin") = "Pin"
myStructure("Heading:item_sku") = "Sku"
myStructure("Heading:other_info") = "Other"
myStructure("Heading:pin_used") = "Used"
'myStructure("Heading:use_date") = "Date"

myStructure("Format:pin") = "STRING"
myStructure("Format:item_sku") = "STRING"
myStructure("Format:other_info") = "STRING"
'myStructure("Format:use_date") = "STRING"
myStructure("Format:pin_used") = "LOOKUP"
myStructure("Lookup:pin_used") = "0:No^-1:Yes"
%>

<!--#include file="head_view.asp"-->

<!--#include file="list_view.asp"-->
<% createFoot thisRedirect,0 %>

