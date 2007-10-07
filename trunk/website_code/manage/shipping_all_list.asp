<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 

sql_region = "SELECT Country, Country_Id FROM Sys_Countries ORDER BY Country;"
set myfields1=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_region,mydata1,myfields1,noRecords1)

sql_location = "SELECT location_name,ship_location_id FROM store_ship_location where Store_Id="&Store_Id&" ORDER BY location_name;"
set myfields2=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_location,mydata2,myfields2,noRecords2)

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Shippers_class_all"
myStructure("ColumnList") = "shipping_method_id,shippers_class,shipping_method_name,base_fee,matrix_low,matrix_high,countries"
if myfields2("rowcount") => 0 then
   myStructure("ColumnList") = myStructure("ColumnList") & ",ship_location_id"
   myStructure("Heading:ship_location_id") = "Location"
   myStructure("Format:ship_location_id") = "SQL"
   myStructure("Sql:ship_location_id") = "select location_name as ship_location_id from store_ship_location where store_Id="&Store_Id&" and Ship_Location_Id=THISFIELD"
   myStructure("Default:ship_location_id") = "Default"
   myStructure("Link:ship_location_id") = "new_location.asp?Op=edit&Id=THISFIELD&"&sAddString
end if

myStructure("DefaultSort") = "Shipping_method_name"
myStructure("PrimaryKey") = "shipping_method_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "shipping"
myStructure("FileName") = "shipping_all_list.asp"
myStructure("FormAction") = "shipping_all_list.asp"
myStructure("Title") = "Shipping Methods"
myStructure("FullTitle") = "Shipping > Methods"
myStructure("CommonName") = "Shipping Method"
myStructure("NewRecord") = "shipping_all_class.asp"
myStructure("Heading:shipping_method_id") = "PK"
myStructure("Heading:shippers_class") = "Type"
myStructure("Heading:shipping_method_name") = "Name"
myStructure("Heading:base_fee") = "Base Fee"
myStructure("Heading:weight_fee") = "Wt Fee"
myStructure("Heading:countries") = "Countries"
myStructure("Heading:zip_start") = "Zip Start"
myStructure("Heading:zip_end") = "Zip End"
myStructure("Heading:matrix_low") = "Matrix Low"
myStructure("Heading:matrix_high") = "Matrix High"
myStructure("Format:zip_start") = "STRING"
myStructure("Format:zip_end") = "STRING"
myStructure("Format:shippers_class") = "LOOKUP"
myStructure("Format:shipping_method_name") = "STRING"
myStructure("Format:base_fee") = "CURR"
myStructure("Format:weight_fee") = "CURR"
myStructure("Format:matrix_low") = "INT"
myStructure("Format:matrix_high") = "INT"
myStructure("Format:countries") = "COLLECTION"
myStructure("Collection:countries") = "country_id,country"
myStructure("Lookup:shippers_class") = "1:Flat Fee^2:Flat Fee + Weight^3:Per Item^4:Total Order Matrix^5:% Total Order^6:Weight Matrix"
sQuestion_Path = "general/shipping.htm"


%>

<!--#include file="head_view.asp"-->
<tr><td bgcolor='#FFFFFF'>
<input type="button" class="Buttons" value="Shipping Settings" name="Add" OnClick='JavaScript:self.location="shipping_class_realtime.asp"'>
<input OnClick='JavaScript:self.location="location_manager.asp"' name="Shipping_Method_List" type="button" class="Buttons" value="Location List">
<input OnClick='JavaScript:self.location="insurance_class.asp"' name="Shipping_Method_List" type="button" class="Buttons" value="Shipping Insurance List">
<br><br>
</td></tr>
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
	end if
	

createFoot thisRedirect, 0
%>
