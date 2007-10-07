			Please note:&nbsp;<br>
			Item_Sku field (i.e. SKU=Stock Keeping Unit Number) Is the unique item identifier:<br>
			If Item_Sku and Pin does not exist in the DB a new item \ record will be inserted into store_pin table.<br>
			otherwise if Item_Sku and Pin already exists in the DB the item \ record will be updated.
			
			<BR><BR>Even if a particular column is not required you must leave the empty column there as a placeholder.
		<BR><BR>
<table width="100%" border="1" cellspacing="0" cellpadding=2>
							<tr bgcolor=#DDDDDD>
								<td class=0><b>Field name</b></td>
								<td><b>Type/Size</b></td>
								<td><b>Required?</b></td>
								<td><b>Example value</b></td>
								<td><b>Description</b></td>
							</tr>
							<tr class=0>
								<td>Item Sku</td>
								<td>TEXT/50</td>
								<td>Y</td>
								<td>UST1001</td>
								<td>The item Sku for which the pin belongs</td>
							</tr>
							<tr class=1>
								<td>Pin</td>
								<td>TEXT/20</td>
								<td>Y</td>
								<td>A1001</td>
								<td>the unique pin number</td>
							</tr>
							<tr class=0>
								<td>Other Info</td>
								<td>TEXT/100</td>
								<td>N</td>
								<td>pin info</td>
								<td>Any additional information to be sent with pin.</td>
							</tr>
						</table>
