Please note:&nbsp;<br>
		Zip Name field is the unique tax identifier:<br>
			If the zip name does not exist in the DB a new record will be inserted into the tax table.<br>
			otherwise record will be updated.&nbsp;
			<BR><BR>Even if a particular column is not required you must leave the empty column there as a placeholder.
		<BR><BR>
      
			
			<table width="100%" border="1" cellspacing="0" cellpadding=2>
				
<tr bgcolor=#DDDDDD>
<td class=0><b>Field name</b></td>
<td><b>Type/Size</b></td>
<td><b>Required?</b></td>
<td><b>Example value</b></td>
<td><b>Description</b></td></tr>

<tr class=1><td>Zip Name</td>
<td>TEXT/50</td>
<td>Y</td>
<td>San Diego County</td>
<td>The unique name given to this tax rate</td></tr>

<tr class=0><td>Zip Start</td>
<td>Number</td>
<td>Y</td>
<td>91000</td>
<td>The starting number of the zip code range</td></tr>

<tr class=1><td>Zip End</td>
<td>Number</td>
<td>Y</td>
<td>91999</td>
<td>The ending number of the zip code range</td></tr>

<tr class=0><td>Tax Rate</td>
<td>Number</td>
<td>Y</td>
<td>6.5</td>
<td>The tax rate to charge, enter 6.5% as 6.5 not as .065</td></tr>

<tr class=1><td>Department Ids</td>
<td>TEXT/1000</td>
<td>N</td>
<td>1,2,3</td>
<td>The ids of the departments to which this tax rate applies.  Leave blank to indicate all departments</td></tr>

<tr class=0><td>Tax Shipping</td>
<td>BIT</td>
<td>Y</td>
<td>Y</td>
<td>Is shipping taxable for this tax rate.</td></tr>

</table>
