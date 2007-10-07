<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
sInstructions="Click on a button below to configure insurance rates.  Please note that insurance is automatically added to shipping prices at checkout. You should define the name of the insurance method to match the shipping method name to which the insurance should be added."

'LOAD CURRENT VALUES FROM THE DATABASE

sFormAction = "Store_Settings.asp"
sName = "Store_Insurance_class"
sFormName = "Insurance_Class"
sTitle = "Shipping > Insurance"
sSubmitName = "Update_Store_Insurance_class"
thisRedirect = "insurance_class.asp"
sTopic="Insurance"
sMenu = "shipping"
sQuestion_Path = "advanced/insurance.htm"
createHead thisRedirect

%>
    <TR bgcolor='#FFFFFF'><td colspan=3>
    <input type="button" class="Buttons" value="Flat Fee" name="Add_Tax" OnClick=JavaScript:self.location="insurance_class1_list.asp?<%= sAddString %>">
    <input type="button" class="Buttons" value="Total Order Matrix" name="Add_Tax" OnClick=JavaScript:self.location="insurance_class4_list.asp?<%= sAddString %>">
    <input type="button" class="Buttons" value="% of Total Order" name="Add_Tax" OnClick=JavaScript:self.location="insurance_class5_list.asp?<%= sAddString %>"></td></tr>


<% createFoot thisRedirect,0 %>
