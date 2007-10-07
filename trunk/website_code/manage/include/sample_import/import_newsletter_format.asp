Please note:&nbsp;<br>
		Email Address field is the unique subscription identifier:<br>
			If email address does not exist in the DB a new record will be inserted into store_newsletter table.<br>
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

<tr class=1><td>Email Address</td>
<td>TEXT/50</td>
<td>Y</td>
<td>Barney@email.com</td>
<td>Email Address of subscriber  </td></tr>

<tr class=0><td>First Name</td>
<td>TEXT/100</td>
<td>Y</td>
<td>Barney</td>
<td>First Name of subscriber </td></tr>

<tr class=1><td>Last Name</td>
<td>TEXT/100</td>
<td>Y</td>
<td>Rubble</td>
<td>Last Name of subscriber </td></tr>

</table>
