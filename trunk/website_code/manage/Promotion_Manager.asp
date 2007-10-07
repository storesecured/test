<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sQuestion_Path = "marketing/promotions.htm"
sTextHelp="promotion/promotions.doc"

sInstructions="Promotions are automatically applied at checkout to customers whose purchases meet your requirements."

set myStructure=server.createobject("scripting.dictionary")
myStructure("TableName") = "Store_Promotions_rules"
myStructure("ColumnList") = "promotion_id,promotion_name,promotion_start_date,promotion_end_date,discounted_items_amount,customer_group,total_from,total_to"
myStructure("HeaderList") = "promotion_name,promotion_start_date,promotion_end_date,discounted_items_amount,customer_group,total_from,total_to"
myStructure("DefaultSort") = "promotion_name"
myStructure("PrimaryKey") = "promotion_id"
myStructure("Level") = 5
myStructure("EditAllowed") = 1
myStructure("AddAllowed") = 1
myStructure("DeleteAllowed") = 1
myStructure("BackTo") = ""
myStructure("Menu") = "marketing"
myStructure("FileName") = "promotion_manager.asp"
myStructure("FormAction") = "promotion_manager.asp"
myStructure("Title") = "Promotions"
myStructure("FullTitle") = "Marketing > Promotions"
myStructure("CommonName") = "Promotion"
myStructure("NewRecord") = "new_promotion.asp"
myStructure("Heading:promotion_id") = "PK"
myStructure("Heading:promotion_name") = "Name"
myStructure("Heading:promotion_start_date") = "Start"
myStructure("Heading:promotion_end_date") = "End"
myStructure("Heading:discounted_items_amount") = "Discount"
myStructure("Heading:customer_group") = "Customer Group"
myStructure("Link:customer_group") = "create_new_customer_group.asp?op=edit&Id=THISFIELD"
myStructure("Heading:total_from") = "Total From"
myStructure("Heading:total_to") = "Total To"
myStructure("Format:promotion_name") = "STRING"
myStructure("Format:promotion_start_date") = "DATESHORT"
myStructure("Format:promotion_end_date") = "DATESHORT"
myStructure("Format:discounted_items_amount") = "PERCENT"
myStructure("Format:customer_group") = "SQL"
myStructure("Sql:customer_group") = "Select Group_name as customer_group from Store_Customers_Groups where group_id=THISFIELD"
myStructure("Default:customer_group") = "All Customers"
myStructure("Format:total_from") = "CURR"
myStructure("Format:total_to") = "CURR"
myStructure("Length:discounted_items_amount") = 0

%>
<!--#include file="head_view.asp"-->
<!--#include file="list_view.asp"-->

<%
createFoot thisRedirect, 0
%>


