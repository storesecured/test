<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

checked2="checked"
op=Request.QueryString("op")
IItem_ID=0
IItem_Quantity=0

Item_Id = Request.QueryString("Item_Id")
Attr_Id = Request.QueryString("Id")
if Attr_Id="" then
   Attr_Id = Request.QueryString("Attribute_Id")
end if

if not isNumeric(Attr_Id) or not isNumeric(Item_Id) then
	Response.Redirect "admin_error.asp?message_id=1"
end if

checked1=""
checked2="checked"
sDiv="none"
sDivDont="block"

sFlashHelp="attributes/attributes.htm"
sMediaHelp="attributes/attributes.wmv"
sZipHelp="attributes/attributes.zip"
sArticleHelp="AddingAttributes.htm"


if op="edit" then
	' IF EDIT, THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlAttribute="select * from store_items_attributes where attribute_id=" & Attr_Id & " and store_id=" & store_id
	rs_Store.open sqlAttribute,conn_store,1,1
	if rs_store.bof or rs_store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	AttributeClass=trim(rs_store("Attribute_Class"))
	AttributeValue=rs_Store("Attribute_Value")
	AttributePriceDifference=rs_store("Attribute_Price_Difference")
	AttributeWeightDifference=rs_store("Attribute_Weight_Difference")
	AttributeSKU=rs_store("Attribute_SKU")
	AttributeHidden=rs_store("Attribute_Hidden")
	DefaultValue=rs_store("Default")
	Display_Order=rs_store("Display_Order")
	IItem_ID=rs_store("IItem_ID")
	IItem_Quantity=rs_store("IItem_Quantity")
	Use_Item=rs_store("Use_Item")
	if Use_Item=-1 then
		checked1="checked"
		sDiv="block"
		sDivDont="none"
		checked2=""
	else
		checked1=""
		checked2="checked"
		sDiv="none"
		sDivDont="block"

	end if
	Attribute_Type=rs_store("Attribute_Type")
	rs_store.close
end if

if Attribute_Type="radio" then
   radio_checked="checked"
else
    dropdown_checked="checked"
end if


sql_select="SELECT Item_Name FROM Store_Items WHERE Store_id="&Store_id&" and Item_Id = "&Item_Id&" ORDER BY Item_Name;"
rs_store.open sql_select,conn_store,1,1 
Item_Name = Rs_store("Item_Name")
rs_store.close

sCommonName="Attribute"
sFormAction = "Inventory_Action.asp"
sFormName = "Add_Attribute"
sName = "Add_Attribute"

sCancel="Item_Attributes.asp?Item_Id="&Item_Id&"&"&sAddString
sFullTitle ="Inventory > <a href=edit_items.asp class=white>Items</a> > <a href=item_edit.asp?Id="&Item_Id&" class=white>Edit</a> > <a href="&sCancel&" class=white>Attributes</a> > "

if op="edit" then
        sFullTitle=sFullTitle&"Edit - "&item_name&" - "&AttributeClass&" "&AttributeValue
        sTitle="Edit Item Attribute - "&item_name&" - "&AttributeClass&" "&AttributeValue
else
	sFullTitle=sFullTitle&"Add - "&item_name
        sTitle ="Add Item Attribute - "&item_name
end if

addPicker = 1
sSubmitName = "Add_Attribute"
thisRedirect = "item_add_attribute_value.asp"
sMenu="inventory"
createHead thisRedirect

if Service_Type < 5	then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=2>
		This feature is not available at your current level of service.<BR><BR>
		SILVER Service or higher is required.
		<a href=billing.asp class=link>Click here to upgrade now.</a>
	</td></tr>
	<% createFoot thisRedirect, 0%>
<% else

'sql_attribute_group =  "SELECT Attribute_class FROM store_items_attributes WHERE (((Store_id)="&Store_id&") AND ((Item_Id)="&Item_Id&")) GROUP BY Attribute_class ORDER BY store_items_attributes.Attribute_class;"

sql_attribute_group =  "SELECT Attribute_class, Attribute_type FROM store_items_attributes WHERE (((Store_id)="&Store_id&") AND ((Item_Id)="&Item_Id&")) GROUP BY Attribute_class, Attribute_type ORDER BY store_items_attributes.Attribute_class;"

set attfields=server.createobject("scripting.dictionary")

Call DataGetrows(conn_store,sql_attribute_group,attdata,attfields,noRecords)

%>
<script language=javascript>
function showHideForm2(div,field,div2,field2,nest){

	e = eval("document.forms[0]."+field+"[0].checked")
	
	 //if (document.forms[0].elements(field)[0].checked)
  
  if (e)
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	  obj.display='block'
	  obj=bw.dom?document.getElementById(div2).style:bw.ie4?document.all[div2].style:bw.ns4?nest?document[nest].document[div2]:document[div2]:0;
	  obj.display='none'
  }
  	else
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	  obj.display='none'
	  obj=bw.dom?document.getElementById(div2).style:bw.ie4?document.all[div2].style:bw.ns4?nest?document[nest].document[div2]:document[div2]:0;
    obj.display='block'
  }
}
function showHideForm3(div,field,div2,field2,nest){

	e = eval("document.forms[0]."+field+"[0].checked")
	// document.forms[0].elements(field)[0].checked
	
  if (e)
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	  obj.display='none'
	  obj=bw.dom?document.getElementById(div2).style:bw.ie4?document.all[div2].style:bw.ns4?nest?document[nest].document[div2]:document[div2]:0;
	  obj.display='block'
  }
  	else
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	  obj.display='block'
	  obj=bw.dom?document.getElementById(div2).style:bw.ie4?document.all[div2].style:bw.ns4?nest?document[nest].document[div2]:document[div2]:0;
    obj.display='none'
  }
}

</script>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Attribute_Id" value="<%=Attr_Id%>">
<input type="Hidden" name="Item_Id" value="<%= Item_Id %>">

   
	<tr bgcolor='#FFFFFF'>
		<td width="100%" colspan="4" height="23">
			<script language="javascript">
					var attr_type=new Array();
			<% if noRecords=0 then %>

					<% FOR attrowcounter= 0 TO attfields("rowcount") %>
					 attr_type["<%= trim(attdata(attfields("attribute_class"),attrowcounter)) %>"]="<%=attdata(attfields("attribute_type"),attrowcounter) %>";
					<% Next %>

			<% end if %>
			</script>

				<tr bgcolor='#FFFFFF'>
					<td width="30%" class="inputname" rowspan=2>Class</td>
				<td width="70%" class="inputvalue" colspan=2>
						Choose one<select size="1" name="Attribute_Class_Pick" onchange="fnType()">
						<% if noRecords=0 then %>
								<% FOR attrowcounter= 0 TO attfields("rowcount") %>
								<option value="<%= trim(attdata(attfields("attribute_class"),attrowcounter)) %>"><%= trim(attdata(attfields("attribute_class"),attrowcounter)) %></option>
								<% Next %>
							<% end if %>
						</select><% small_help "Class" %></td>
						<tr bgcolor='#FFFFFF'><td colspan=2>OR NEW

					<input	name="Attribute_Class_New" size="60"
							<% if op="edit" then %>
								value="<%= AttributeClass %>"
							<% end if %>			 
						><INPUT type="hidden" name="Attribute_Class_New_C" value="Op|String|||||New Class">
						<% small_help "Class" %></td>
				</tr> 
		
			<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Value</td>
				<td width="70%" class="inputvalue" colspan=2>
					<input	name="Attribute_Value" size="60"
							<% if op="edit" then %>
								 value="<%= AttributeValue %>"
							<% end if %>
							>
						<INPUT type="hidden"  name="Attribute_Value_C" value="Re|String|||||Value">
						<% small_help "Value" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname" rowspan=2>Type</td>
				<td width="70%" class="inputvalue" colspan=2>
					<input class="image" type="radio" name="Attribute_Type" value="dropdown" <%=dropdown_checked%>>Dropdown<% small_help "Type" %></td></tr>
				<tr bgcolor='#FFFFFF'><td width="70%" class="inputvalue" colspan=2>
				<input class="image" type="radio" name="Attribute_Type" value="radio" <%=radio_checked%>>Radio Buttons</td>
						<!--<INPUT type="hidden"  name="Attribute_Value_C" value="Re|String|||||Value">-->
						<% small_help "Type" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td class="inputname" rowspan=5>Use inventory item</td><td class="inputvalue" rowspan=2><input class="image" type="radio" OnClick="showHideForm2('Use','Use_Item','DontUse','Use_Item');" name="Use_Item" value="-1" <%= checked1 %>> Yes
					
				<td class="inputname">Item ID to add to cart
						<input name="IItem_ID" size="5" value="<%= IItem_ID %>">
						<INPUT type="hidden" name="IItem_ID_C" value="Re|Integer|||||Item ID">
						<a class="link" href="JavaScript:goItemPicker('IItem_ID');">Item Picker</a>
						<% small_help "Item ID" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td class="inputname">Quantity to add to cart
						<input name="IItem_Quantity" size="5" value="<%= IItem_Quantity %>" onKeyPress="return goodchars(event,'0123456789.')">
						<INPUT type="hidden" name="IItem_Quantity_C" value="Re|Integer|||||Quantity">
						<% small_help "Quantity" %></td>
				</tr>
				
				<tr bgcolor='#FFFFFF'>
					<td class="inputvalue" rowspan=3><input class="image" type="radio" OnClick="showHideForm3('DontUse','Use_Item','Use','Use_Item');" name="Use_Item" value="0" <%= checked2 %>> No
					<td class="inputname">Price Diff
					<%= Store_Currency %><input	name="Attribute_Price_Difference" size="5" onKeyPress="return goodchars(event,'0123456789.-')"
							<% if op="edit" then %>
								value="<%= AttributePriceDifference %>"
							<% else %>
								value="0"
							<% end if %>
						>
						<INPUT type="hidden"  name="Attribute_Price_Difference_C" value="Re|Integer|||||Price Diff">
						<% small_help "Price Diff" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td class="inputname">Weight Diff
					<input	name="Attribute_Weight_Difference" size="5" onKeyPress="return goodchars(event,'0123456789.-')"
							<% if op="edit" then %>
								value="<%= AttributeWeightDifference %>"
							<% else %>
								value="0"
							<% end if %>
						>
					<INPUT type="hidden"  name="Attribute_Weight_Difference_C" value="Re|Integer|||||Weight Diff">
					<% small_help "Weight Diff" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td class="inputname">SKU
					<input	name="Attribute_sku" size="60"
						<% if op="edit" then %>
								value="<%= AttributeSKU %>"
						<% end if %>
						>
					<INPUT type="hidden" name="Attribute_sku_C" value="Op|String|0|50|||Sku">
					<% small_help "SKU" %>
				</td>
				</tr>
				

				<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Hidden value - not seen by customers</td>
				<td width="70%" class="inputvalue" colspan=2>
					<input name="Attribute_hidden" size="60"
							<% if op="edit" then %> 
								value="<%= AttributeHidden %>"
							<% end if %>
						>
					<INPUT type="hidden" name="Attribute_hidden_C" value="Op|String|||||Hidden Value">
					<% small_help "Hidden value" %></td>
				</tr>
				
				<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Default</td>
				<td width="70%" class="inputvalue" colspan=2>
					<input class="image" type="checkbox" name="Default" size="5"
							<% if op="edit" then %>
								<% if DefaultValue<>0 then %>
									checked
								<% end if %>
							<% end if %>
						><% small_help "Default" %>
					</td></tr>
							<tr bgcolor='#FFFFFF'>
				<td width="30%" class="inputname">Display Order</td>
				<td width="70%" class="inputvalue" colspan=2>
					<input name="Display_Order" size="5" onKeyPress="return goodchars(event,'0123456789.-')"
							<% if op="edit" then %> 
								value="<%= Display_Order %>"
							<% else %>
								  value="0"
							<% end if %>
						>
					<INPUT type="hidden" name="Display_Order_C" value="Op|Number|||||Display Order">
					<% small_help "Display Order" %></td>
				</tr>
<% createFoot thisRedirect, 1%>

<SCRIPT language="JavaScript">
fnType(); 
 function fnType()
 {
 	var cls_type=attr_type[document.forms[0].Attribute_Class_Pick.value];
		for (i=0;i<document.forms[0].Attribute_Type.length;i++)
			if (document.forms[0].Attribute_Type[i].value==cls_type)
				document.forms[0].Attribute_Type[i].checked=true;
 }
 
 
 
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Attribute_Value","req","Please enter a value.");
 frmvalidator.addValidation("IItem_ID","req","Please enter a item id.");
 frmvalidator.addValidation("IItem_Quantity","req","Please enter a quantity.");
 frmvalidator.addValidation("Attribute_Price_Difference","req","Please enter a price difference.");
 frmvalidator.addValidation("Attribute_Weight_Difference","req","Please enter a weight difference.");
 frmvalidator.addValidation("Display_Order","req","Please enter a display order.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");

function DoCustomValidation()
{
  if (document.forms[0].Attribute_Class_Pick.value == "")
  {
     if (document.forms[0].Attribute_Class_New.value == "")
      {
        alert('Please choose a class from the dropdown or enter a new class name.');
        return false;
      }
  }
  if (document.forms[0].Item_Id.value == document.forms[0].IItem_ID.value)
  {
    alert('An item cannot be an attribute of itself.  Please choose a new item id.');
    return false;
  }
  else
  {
      return true;
      

  }
}

</script>
<% end if %>
