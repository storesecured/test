Please note:&nbsp;<br>
		Item_Sku field (i.e. SKU=Stock Keeping Unit Number) Is the unique item identifier:<br>
			If Item_Sku does not exist in the DB a new item \ record will be inserted into store_items table.<br>
			otherwise if Item_Sku already exists in the DB the item \ record will be updated.&nbsp;<br>
		<BR><BR>Even if a particular column is not required you must leave the empty column there as a placeholder.
		<BR><BR>
      <table width="100%" border="1" cellspacing="0" cellpadding=2>
				
								<tr bgcolor=#DDDDDD>
<td class=0><b>Field name</b></td>
<td><b>Type/Size</b></td>
<td><b>Required?</b></td>
<td><b>Excel Column</b></td>
<td><b>Example value</b></td>
<td><b>Description</b></td></tr>

<tr class=1><td>Item Sku</td>
<td>TEXT/200</td>
<td>Y</td>
<td>A</td>
<td>1SN34TRQ</td>
<td>The item Stock Keeping Unit</td></tr>

<tr class=0><td>Departments</td>
<td>TEXT/3000</td>
<td>N</td>
<td>B</td>
<td>Housewares:Silver</td>
<td>The item Departments. Use a colon to separate multiple departments.</td></tr>

<tr class=1><td>Item Name</td>
<td>TEXT/250</td>
<td>N</td>
<td>C</td>
<td>3-Piece Knife Set</td>
<td>The item name</td></tr>

<tr class=0><td>Retail Price</td>
<td>Number</td>
<td>N</td>
<td>d</td>
<td>9.99</td>
<td>Item Retail price (the price that we are going to display in store ...)</td></tr>

<tr class=1><td>Cost</td>
<td>Number</td>
<td>N</td>
<td>e</td>
<td>8.99</td>
<td>Price that item cost us</td></tr>

<tr class=0><td>Small Image file name</td>
<td>TEXT/100</td>
<td>N</td>
<td>F</td>
<td>small_knife.gif</td>
<td>Name of gif or jpeg image file</td>

<tr class=1><td>Large Image file name</td>
<td>TEXT/100</td>
<td>N</td>
<td>G</td>
<td>big_knife.gif</td>
<td>Name of gif or jpeg image file for product detail page</td>

<tr class=0><td>Taxable</td>
<td>BIT</td>
<td>N</td>
<td>H</td>
<td>Y</td>
<td>Is item taxable ? if no then we will not take it into account when calculating tax for order total</td></tr>

<tr class=1><td>Small Description</td>
<td>MEMO</td>
<td>N</td>
<td>I</td>
<td>Small Description</td>
<td>The item description</td></tr>

<tr class=0><td>Large Description</td>
<td>MEMO</td>
<td>N</td>
<td>J</td>
<td>Large Description</td>
<td>The item description for the more product detail page</td></tr>

<tr class=1><td>Item Weight</td>
<td>Number</td>
<td>N</td>
<td>K</td>
<td>6.5</td>
<td>Weight In Any unit (as long it's consistent)</td></tr>

<tr class=0><td>Quantity in stock</td>
<td>Number</td>
<td>N</td>
<td>L</td>
<td>55</td>
<td>how many items in stock?</td></tr>

<tr class=1><td>Quantity Control</td>
<td>BIT</td>
<td>N</td>
<td>M</td>
<td>Y</td>
<td>Enter Y if you want to disable ordering for this item when it's quantity is below threshold</td></tr>

<tr class=0><td>Quantity Control Number</td>
<td>Number</td>
<td>N</td>
<td>N</td>
<td>3</td>
<td>The threshold for quantity control</td></tr>

<tr class=1><td>Retail Price special Discount</td>
<td>Number</td>
<td>N</td>
<td>O</td>
<td>20</td>
<td>Item Retail price discount (0-100 0=full price;10 = 10% discount)</td></tr>

<tr class=0><td>Special start date</td>
<td>Date/Time</td>
<td>N</td>
<td>P</td>
<td>05/09/2001</td>
<td>When this special starts</td></tr>

<tr class=1><td>Special end date</td>
<td>Date/Time</td>
<td>N</td>
<td>Q</td>
<td>09/09/2001</td>
<td>When this special ends</td></tr>

<tr class=0><td>Unused currently</td>
<td>MEMO</td>
<td>N</td>
<td>R</td>
<td></td>
<td>Field not used</td></tr>

<tr class=1><td>Shipping Fee</td>
<td>Number</td>
<td>N</td>
<td>S</td>
<td>6.00</td>
<td>Shipping fee for this item - Used when shipping class is "no shipping"</td></tr>

<tr class=0><td>Downloadable Filename</td>
<td>TEXT/250</td>
<td>N</td>
<td>T</td>
<td>myfile.exe</td>
<td>If this item is a downloadable file, note that files need to be uploaded first using the Upload Files Link under Inventory</td></tr>

<tr class=1><td>Item Remarks</td>
<td>MEMO</td>
<td>N</td>
<td>U</td>
<td>Hot item, re order soon!</td>
<td>Remarks for item, to be seen by store admin when viewing invoice only!</td></tr>
</tr>

<tr class=0><td>Fractional Selling</td>
<td>BIT</td>
<td>N</td>
<td>V</td>
<td>N</td>
<td>Allow customers to purchase a portion of an item</td></tr>
</tr>

<tr class=1><td>Show</td>
<td>BIT</td>
<td>N</td>
<td>W</td>
<td>Y</td>
<td>Show this item to customers</td></tr>
</tr>

<tr class=0><td>Item Filename</td>
<td>TEXT/100</td>
<td>N</td>
<td>X</td>
<td>three-piece-knife-set</td>
<td>Filename for this item should contain only letters and numbers, no special chars.  </td></tr>


<tr class=1><td>Attributes</td>
<td>String</td>
<td>N</td>
<td>Y</td>
<td>
Size:Small:1:1:0:0:<BR>
0::|Size:Medium:2:<BR>
0:0:0:0::</td>
<td>List of attributes separated by | each attribute separated by :<BR>Format as follows:
<BR>
When using inventory item<BR>
Class:Value:Order:Default:-1: ItemId:Quantity
<BR><BR>
When not using inventory item<BR>
Class:Value:Order:Default:0: PriceDiff:WeightDiff:Sku:Hidden
</td></tr>



<tr class=0><td>View Order</td>
<td>Number</td>
<td>N</td>
<td>Z</td>
<td>0</td>
<td>The display order for the inventory item</td></tr>
</tr>

<tr class=1><td>Meta Keywords</td>
<td>TEXT/250</td>
<td>N</td>
<td>AA</td>
<td>battery,camera battery</td>
<td>The meta keywords to use for this item</td></tr>
</tr>

<tr class=0><td>Meta Description</td>
<td>TEXT/500</td>
<td>N</td>
<td>AB</td>
<td>A description for this item</td>
<td>The meta description to use for this item</td></tr>
</tr>

<tr class=1><td>Item Accessories</td>
<td>TEXT/500</td>
<td>N</td>
<td>AC</td>
<td>Sku1:Sku2</td>
<td>Listing of item skus, separate multiple items with a colon</td>
</tr>

<tr class=0><td>Meta Title</td>
<td>TEXT/100</td>
<td>N</td>
<td>AD</td>
<td>Item Title</td>
<td>The meta title to use for this item</td>
</tr>

<tr class=1><td>Ship Location</td>
<td>TEXT/100</td>
<td>N</td>
<td>AE</td>
<td>California</td>
<td>The name of the shipping location, use 0 for the default location</td>
</tr>
<tr class=0><td>Extended Field 1</td>
<td>MEMO</td>
<td>N</td>
<td>AF</td>
<td>Manufacturer Name</td>
<td>Extended Field 1 to be displayed in addition to the item details</td></tr>
</tr>

<tr class=1><td>Extended Field 2</td>
<td>MEMO</td>
<td>N</td>
<td>AG</td>
<td>Width</td>
<td>Extended Field 2 to be displayed in addition to the item details</td></tr>
</tr>


<tr class=0><td>Extended Field 3</td>
<td>MEMO</td>
<td>N</td>
<td>AH</td>
<td>Height</td>
<td>Extended Field 3 to be displayed in addition to the item details</td></tr>
</tr>


<tr class=1><td>Extended Field 4</td>
<td>MEMO</td>
<td>N</td>
<td>AI</td>
<td>Weight</td>
<td>Extended Field 4 to be displayed in addition to the item details</td></tr>
</tr>

<tr class=0><td>Extended Field 5</td>
<td>MEMO</td>
<td>N</td>
<td>AJ</td>
<td>Other</td>
<td>Extended Field 5 to be displayed in addition to the item details</td></tr>
</tr>

<tr class=1><td>Custom Link</td>
<td>TEXT/250</td>
<td>N</td>
<td>AK</td>
<td>URL</td>
<td>New Url for Item to redirect to from the Cart </td></tr>
</tr>

<tr class=0><td>User Defined Field 1</td>
<td>TEXT/250</td>
<td>N</td>
<td>AL</td>
<td>Field1</td>
<td>Custom named field, provides a textbox for the user to fill in when purchasing item.</td></tr>
</tr>

<tr class=1><td>Use User Defined Field 1</td>
<td>BIT</td>
<td>N</td>
<td>AM</td>
<td>Y</td>
<td>Show this field to customers</td></tr>
</tr>


<tr class=0><td>User Defined Field 2</td>
<td>TEXT/250</td>
<td>N</td>
<td>AN</td>
<td>Field2</td>
<td>Custom named field, provides a textbox for the user to fill in when purchasing item.</td></tr>
</tr>

<tr class=1><td>Use User Defined Field 2</td>
<td>BIT</td>
<td>N</td>
<td>AO</td>
<td>Y</td>
<td>Show this field to customers</td></tr>
</tr>

<tr class=0><td>User Defined Field 3</td>
<td>TEXT/250</td>
<td>N</td>
<td>AP</td>
<td>Field3</td>
<td>Custom named field, provides a textbox for the user to fill in when purchasing item.</td></tr>
</tr>


<tr class=1><td>Use User Defined Field 3</td>
<td>BIT</td>
<td>N</td>
<td>AQ</td>
<td>Y</td>
<td>Show this field to customers</td></tr>
</tr>

<tr class=0><td>User Defined Field 4</td>
<td>TEXT/250</td>
<td>N</td>
<td>AR</td>
<td>Field4</td>
<td>Custom named field, provides a textbox for the user to fill in when purchasing item.</td></tr>
</tr>


<tr class=1><td>Use User Defined Field 4</td>
<td>BIT</td>
<td>N</td>
<td>AS</td>
<td>Y</td>
<td>Show this field to customers</td></tr>
</tr>

<tr class=0><td>User Defined Field 5</td>
<td>TEXT/250</td>
<td>N</td>
<td>AT</td>
<td>Field5</td>
<td>Custom named field, provides a textbox for the user to fill in when purchasing item.</td></tr>
</tr>

<tr class=1><td>Use User Defined Field 5</td>
<td>BIT</td>
<td>N</td>
<td>AU</td>
<td>Y</td>
<td>Show this field to customers</td></tr>
</tr>
 <tr class=0><td>Handling Fee</td>
<td>Number</td>
<td>N</td>
<td>AV</td>
<td>6.00</td>
<td>Handling fee for this item - Added onto any shipping costs</td></tr>
 <tr class=0><td>Brand</td>
<td>TEXT/100</td>
<td>N</td>
<td>AW</td>
<td>Sony Ericcson</td>
<td>Brand name for this product, used if you are using google base</td></tr>
<tr class=0><td>Condition</td>
<td>TEXT/20</td>
<td>N</td>
<td>AX</td>
<td>New</td>
<td>Product condition, either New, Used or Refurbished, used if you are using google base</td></tr>
<tr class=0><td>Product Type</td>
<td>TEXT/20</td>
<td>N</td>
<td>AY</td>
<td>Camera</td>
<td>Product type, please see google base categories for sample ideas, used if you are using google base</td></tr>

<tr class=1><td>Price Matrix</td>
<td>String</td>
<td>N</td>
<td>AZ</td>
<td>
10:20:9.99|21:-1:7.99</td>
<td>List of matrixes separated by | each value separated by :<BR>Format as follows:
<BR>
Low Quantity:High Quantity:Price
</td></tr>

<tr class=0><td>Price Group</td>
<td>String</td>
<td>N</td>
<td>BA</td>
<td>
Wholesale:9.99|Distributor:7.99</td>
<td>List of customer groups separated by | each value separated by :<BR>Format as follows:
<BR>
Customer group name:Price
</td></tr>

</table>
