<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
sQuestion_Path = "advanced/insurance.htm"
sTextHelp="insurance/matrix_insurance.doc"

sInstructions="Insurance based on total order amount can be used to setup different insurance amounts for different order values."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Insurance_class_all"
myStructure("TableWhere") = "Insurance_class=4"
myStructure("ColumnList") = "insurance_method_id,insurance_method_name,base_fee,matrix_low,matrix_high"

myStructure("DefaultSort") = "insurance_method_name"
myStructure("PrimaryKey") = "insurance_method_id"
myStructure("Level") = 0
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("Menu") = "shipping"
myStructure("FileName") = "insurance_class4_list.asp"
myStructure("FormAction") = "insurance_class4_list.asp"
myStructure("Title") = "Total Order Matrix Shipping Insurance"
myStructure("FullTitle") = "Shipping > <a href=insurance_class.asp class=white>Insurance</a> > View/Edit Total Order Matrix"
myStructure("CommonName") = "Insurance Method"
myStructure("NewRecord") = "insurance_class4_add.asp"
myStructure("Heading:insurance_method_id") = "PK"
myStructure("Heading:insurance_method_name") = "Name"
myStructure("Heading:base_fee") = "Base Fee"
myStructure("Format:insurance_method_name") = "STRING"
myStructure("Heading:matrix_low") = "Matrix Low"
myStructure("Heading:matrix_high") = "Matrix High"
myStructure("Format:matrix_low") = "CURR"
myStructure("Format:matrix_high") = "CURR"
myStructure("Format:base_fee") = "CURR"


%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
  if Request.QueryString("Delete_Id") <> "" then
												end if
createFoot thisRedirect, 0
%>

