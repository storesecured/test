<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
sQuestion_Path = "advanced/insurance.htm"
sTextHelp="insurance/percentage_insurance.doc"
sInstructions="Percentage of total order insurance can be used to charge all customers a percentage of their order total as a insurance charge."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Insurance_class_all"
myStructure("TableWhere") = "Insurance_class=5"
myStructure("ColumnList") = "insurance_method_id,insurance_method_name,base_fee,matrix_low,matrix_high"

myStructure("DefaultSort") = "insurance_method_name"
myStructure("PrimaryKey") = "insurance_method_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "shipping"
myStructure("FileName") = "insurance_class5_list.asp"
myStructure("FormAction") = "insurance_class5_list.asp"
myStructure("Title") = "% Total Order Shipping Insurance"
myStructure("FullTitle") = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > View/Edit % Total Order"
myStructure("CommonName") = "Insurance Method"
myStructure("NewRecord") = "insurance_class5_add.asp"
myStructure("Heading:insurance_method_id") = "PK"
myStructure("Heading:insurance_method_name") = "Name"
myStructure("Heading:base_fee") = "Fee"
myStructure("Format:insurance_method_name") = "STRING"
myStructure("Format:base_fee") = "PERCENT"

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
	end if
createFoot thisRedirect, 0
%>

