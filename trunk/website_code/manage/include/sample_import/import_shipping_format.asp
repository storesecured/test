Please note:&nbsp;<br>
		Name is the unique shipping identifier:<br>
			If the name does not exist in the DB a new item \ record will be inserted.<br>
			otherwise if the name already exists in the DB the item \ record will be updated.
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
						<tr class=1>
							<td>ship class</td>
							<td>Number</td>
							<td>Y</td>
							<td>4</td>
							<td>calculate shipping type: use<BR> 1 for flat fee,<BR> 2 for Flat fee + weight,<BR> 3 for Per Item,<BR> 4 for Total Order Matrix,<BR> 5 for % of total order,<BR> 6 for Total Weight Matrix.</td>
						</tr>
						<tr class=0>
							<td>Name</td>
							<td>TEXT/200</td>
							<td>Y</td>
							<td>UPS</td>
							<td>Name of Shipping method</td>
						</tr>
						<tr class=1>
							<td>Matrix low</td>
							<td>Number</td>
							<td>N</td>
							<td>3.00</td>
							<td>Real values </td>
						</tr>
						<tr class=0>
							<td>Matrix high</td>
							<td>Number</td>
							<td>N</td>
							<td>9.99</td>
							<td>Real Values </td>
						</tr>
						<tr class=1>
							<td>Base fee </td>
							<td>Number</td>
							<td>N</td>
							<td>8.99</td>
							<td>base fee in money </td>
						</tr>
						<tr class=0>
							<td>Weight fee</td>
							<td>Number</td>
							<td>N</td>
							<td>10.00</td>
							<td>Weight fee in money </td>
						</tr>
						<tr class=1>
							<td>Countries</td>
							<td>Text/4000</td>
							<td>Y</td>
							<td>India:All Countries</td>
							<td>Name of countries,use colon to delimit multiple countries</td>
						</tr>
						<tr class=0>
							<td>Ship Location</td>
							<td>Text/100</td>
							<td>N</td>
							<td>test</td>
							<td>Location name, use Default if you have not created any shipping locations</td>
						</tr>
						<tr class=1><td>Zip Start</td>
							<td>Text/10</td>
							<td>N</td>
							<td>31524</td>
							<td>Zip Code from </td>
						</tr>
						<tr class=0>
							<td>Zip End</td>
							<td>Text/10</td>
							<td>N</td>
							<td>31525</td>
							<td>Zip code To </td>
						</tr>
						<tr class=1>
							<td>Always Insert</td>
							<td>Bit</td>
							<td>N</td>
							<td>N</td>
							<td>Overwrite imports normal handling to update shipping methods with the same name and insert this shipping method as new.</td>
						</tr>
					</table>
