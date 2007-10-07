<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "advanced/location_manager.htm"
sInstructions="Locations can be used to specify different items coming from different physical locations.  For realtime shipping this enables actual pricing based on real location of the product.  For other shipping methods the locations can be used to define different costs for different locations different wholesale shipping costs.  By default a store has 1 shipping method which is your base location and all items come from that location.  Once you add additional shipping locations you will be prompted on each item entry to define which location that item comes from."
sTextHelp="shipping/shipping-locations.doc"


set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Ship_Location"
myStructure("ColumnList") = "ship_location_id,location_name,location_zip,location_country"
myStructure("DefaultSort") = "location_name"
myStructure("PrimaryKey") = "ship_location_id"
myStructure("Level") = 7
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "shipping"
myStructure("FileName") = "location_Manager.asp"
myStructure("FormAction") = "location_Manager.asp"
myStructure("Title") = "Shipping Locations"
myStructure("FullTitle") = "Shipping > Locations"
myStructure("CommonName") = "Shipping Location"
myStructure("NewRecord") = "new_location.asp"
myStructure("Heading:ship_location_id") = "PK"
myStructure("Heading:location_name") = "Name"
myStructure("Heading:location_zip") = "Zip"
myStructure("Heading:location_country") = "Country"
myStructure("Format:location_name") = "STRING"
myStructure("Format:location_zip") = "STRING"
myStructure("Format:location_country") = "STRING"
%>

<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
		Delete_id = Request.QueryString("Delete_Id")
		sql_delete = "delete from store_shippers_class_all where Store_id = "&Store_id&" and Ship_location_id="&Delete_Id
		conn_store.Execute sql_update
		Session("Location_List:"&Store_Id)=""
	end if
	
	if  Request.Form("Delete_Id") ="SEL" then
  	if request.form("DELETE_IDS") <> "" then
      Delete_Id = request.form("DELETE_IDS")
      sql_delete = "delete from store_shippers_class_all where Store_id = "&Store_id&" and Ship_location_id in ("&Delete_Id&")"
		  conn_store.Execute sql_update
		  Session("Location_List:"&Store_Id)=""
	  end if
  end if
createFoot thisRedirect, 0
%>

