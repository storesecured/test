Please note:&nbsp;<br>
		All 3 fields are used together to uniquely identify new records:<br>
			If the combination of the 3 fields does not exist in the DB a new budget \ record will be inserted.<br>
			otherwise if the 3 fields already exist in the DB the item \ record will be updated.
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

								<tr class=1><td>Budget_Amount</td>
									<td>Number</td>
									<td>Y</td>
									<td>45.00</td>
									<td>it stores the budget left </td>
								</tr>

								<tr class=0><td>Customer Id</td>
									<td>Number</td>
									<td>Y</td>
									<td>4991</td>
									<td>the unique id of the customer</td>
								</tr>

								<tr class=1><td>Notes</td>
									<td>Text/1000</td>
									<td>N</td>
									<td>oid 7 </td>
									<td>reason for buget change </td>
								</tr>
							</table>
