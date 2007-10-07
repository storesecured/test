<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
sQuestion_Path = "advanced/insurance.htm"
sTextHelp="insurance/flat_insurance.doc"

sInstructions="Flat fee insurance can be used to charge all customers the same rate regardless of how many items they have purchased or the total weight of the items.  This is the simplest of all insurance methods."


set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Insurance_class_all"
myStructure("TableWhere") = "Insurance_class=1"
myStructure("ColumnList") = "insurance_method_id,insurance_method_name,base_fee"

myStructure("DefaultSort") = "insurance_method_name"
myStructure("PrimaryKey") = "insurance_method_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "shipping"
myStructure("FileName") = "insurance_class1_list.asp"
myStructure("FormAction") = "insurance_class1_list.asp"
myStructure("Title") = "Flat Fee Shipping Insurance"
myStructure("FullTitle") = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > View/Edit Flat Fee"
myStructure("CommonName") = "Insurance Method"
myStructure("NewRecord") = "insurance_class1_add.asp"
myStructure("Heading:insurance_method_id") = "PK"
myStructure("Heading:insurance_method_name") = "Name"
myStructure("Heading:base_fee") = "Flat Fee"
myStructure("Format:insurance_method_name") = "STRING"
myStructure("Format:base_fee") = "CURR"

%>

<!--#include file="head_view.asp"-->

<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
	end if
	
createFoot thisRedirect, 0
%>

