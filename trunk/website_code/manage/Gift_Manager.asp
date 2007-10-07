<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "marketing/gift_certificates.htm"
sInstructions="Gift certificates are electronic codes which can be purchased by your customers and used for future purchases.  The gift certificate code is emailed to the recipient who can use the certificate until it has either expired or there is no remaining balance.<BR><BR>You may use user defined field #1 for recipients email and user defined field #2 for a gift message on gift certificate inventory items."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Gift_Certificates"
myStructure("ColumnList") = "gift_id,gift_name,amount,validity_time,item,restricted_items"
myStructure("HeaderList") = "gift_name,amount,validity_time,item,restricted_items,purchases"
myStructure("DefaultSort") = "gift_name"
myStructure("PrimaryKey") = "gift_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "Gift_Manager.asp"
myStructure("FormAction") = "Gift_Manager.asp"
myStructure("Title") = "Gift Certificates"
myStructure("FullTitle") = "Marketing > Gift Certificates"
myStructure("CommonName") = "Gift Certificate"
myStructure("NewRecord") = "new_gift.asp"
myStructure("Heading:gift_id") = "PK"
myStructure("Heading:gift_name") = "Name"
myStructure("Heading:amount") = "Amount"
myStructure("Heading:validity_time") = "Valid Days"
myStructure("Heading:item") = "Item"
myStructure("Heading:restricted_items") = "Restricted"
myStructure("Heading:purchases") = "Purchases"
myStructure("Format:gift_name") = "STRING"
myStructure("Format:amount") = "CURR"
myStructure("Format:validity_time") = "INT"
myStructure("Format:item") = "STRING"
myStructure("Format:restricted_items") = "STRING"
myStructure("Format:purchases") = "TEXT"
myStructure("Link:purchases") = "gift_list.asp?Id=PK"
myStructure("Link:item") = "item_edit.asp?Id=THISFIELD"
myStructure("Length:validity_time") = 0
%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
    Delete_id = Request.QueryString("Delete_Id")
    sql_del = "delete from store_gift_certificates_dets where store_id="&store_id&" and gift_id="&Delete_Id
		conn_store.Execute sql_del
	end if

	if Request.Form("Delete_Id") ="SEL" then
  	if request.form("DELETE_IDS") <> "" then
  	   Delete_Id = request.form("DELETE_IDS")
       sql_del = "delete from store_gift_certificates_dets where store_id="&store_id&" and gift_id in ("&Delete_Id&")"
		   conn_store.Execute sql_del
	  end if
  end if
createFoot thisRedirect, 0
%>

