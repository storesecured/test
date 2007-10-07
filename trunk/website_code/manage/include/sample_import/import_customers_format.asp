Please note:&nbsp;<br>
		Username Is the unique customer identifier:<br>
			If Username does not exist in the DB a new customer \ record will be inserted into store_customers table.<br>
			otherwise if Username already exists in the DB the customer \ record will be updated.&nbsp;
         <BR><BR>Even if a particular column is not required you must leave the empty column there as a placeholder.
		<BR><BR>
         <table width="100%" border="1" cellspacing="0" cellpadding=2>
				
								<tr bgcolor=#DDDDDD>
<td class=0><b>Field name</b></td>
<td><b>Type/Size</b></td>
<td><b>Required?</b></td>
<td><b>Example value</b></td>
<td><b>Description</b></td></tr>

<tr class=1><td>Username</td>
<td>TEXT/50</td>
<td>Y</td>
<td>barnie</td>
<td>Username for login purposes</td></tr>

<tr class=0><td>Password</td>
<td>TEXT/50</td>
<td>Y</td>
<td>pw</td>
<td>Password for login purposes</td></tr>

<tr class=1><td>Billing Last Name</td>
<td>TEXT/100</td>
<td>Y</td>
<td>Rubble</td>
<td>The last name on the bill to address</td></tr>

<tr class=0><td>Billing First Name</td>
<td>TEXT/100</td>
<td>Y</td>
<td>Barney</td>
<td>The first name on the bill to address</td></tr>

<tr class=1><td>Billing Company</td>
<td>TEXT/100</td>
<td>N</td>
<td>The Quarry</td>
<td>Name of company on the bill to address</td></tr>

<tr class=0><td>Billing Address1</td>
<td>TEXT/200</td>
<td>Y</td>
<td>123 Bedrock Lane</td>
<td>First line of billing address information </td>

<tr class=1><td>Billing Address2</td>
<td>TEXT/200</td>
<td>N</td>
<td>Apt 1</td>
<td>Second line of billing address information</td>

<tr class=0><td>Billing City</td>
<td>TEXT/200</td>
<td>Y</td>
<td>Bedrock</td>
<td>City name for billing purposes</td></tr>

<tr class=1><td>Billing State</td>
<td>TEXT/2</td>
<td>Y</td>
<td>CA</td>
<td>The 2 letter state abbrev, use AA for outside US and Canada for billing information</td></tr>

<tr class=0><td>Billing Zip Code</td>
<td>TEXT/50</td>
<td>Y</td>
<td>92026</td>
<td>Zip code for billing purposes</td></tr>

<tr class=1><td>Billing Country</td>
<td>TEXT/200</td>
<td>Y</td>
<td>United States</td>
<td>Full Country Name for billing purposes</td></tr>

<tr class=0><td>Billing Phone</td>
<td>TEXT/50</td>
<td>Y</td>
<td>(866)324-2764</td>
<td>Phone number formatted the way you want it displayed for billing purposes</td></tr>

<tr class=1><td>Billing Fax</td>
<td>TEXT/50</td>
<td>N</td>
<td>(866)324-2764</td>
<td>Fax number formatted the way you want it displayed for billing purposes</td></tr>

<tr class=0><td>Billing Email Address</td>
<td>TEXT/100</td>
<td>Y</td>
<td>brubble@quarry.com</td>
<td>The email address to send billing information to</td></tr>

<tr class=1><td>Shipping Last Name</td>
<td>TEXT/100</td>
<td>Y</td>
<td>Rubble</td>
<td>The last name on the ship to address</td></tr>

<tr class=0><td>Shipping First Name</td>
<td>TEXT/100</td>
<td>Y</td>
<td>Barney</td>
<td>The first name on the ship to address</td></tr>

<tr class=1><td>Shipping Company</td>
<td>TEXT/100</td>
<td>N</td>
<td>The Quarry</td>
<td>Name of company on the ship to address</td></tr>

<tr class=0><td>Shipping Address1</td>
<td>TEXT/200</td>
<td>Y</td>
<td>123 Bedrock Lane</td>
<td>First line of shipping address information </td>

<tr class=1><td>Shipping Address2</td>
<td>TEXT/200</td>
<td>N</td>
<td>Apt 1</td>
<td>Second line of shipping address information</td>

<tr class=0><td>Shipping City</td>
<td>TEXT/200</td>
<td>Y</td>
<td>Bedrock</td>
<td>City name for shipping purposes</td></tr>

<tr class=1><td>Shipping State</td>
<td>TEXT/2</td>
<td>Y</td>
<td>CA</td>
<td>The 2 letter state abbrev, use AA for outside US and Canada for shipping information</td></tr>

<tr class=0><td>Shipping Zip Code</td>
<td>TEXT/50</td>
<td>Y</td>
<td>92026</td>
<td>Zip code for shipping purposes</td></tr>

<tr class=1><td>Shipping Country</td>
<td>TEXT/200</td>
<td>Y</td>
<td>United States</td>
<td>Full Country Name for shipping purposes</td></tr>

<tr class=0><td>Shipping Phone</td>
<td>TEXT/50</td>
<td>Y</td>
<td>(866)324-2764</td>
<td>Phone number formatted the way you want it displayed for shipping purposes</td></tr>

<tr class=1><td>Shipping Fax</td>
<td>TEXT/50</td>
<td>N</td>
<td>(866)324-2764</td>
<td>Fax number formatted the way you want it displayed for shipping purposes</td></tr>

<tr class=0><td>Shipping Email Address</td>
<td>TEXT/100</td>
<td>Y</td>
<td>brubble@quarry.com</td>
<td>The email address to send shipping information to</td></tr>

<tr class=1><td>Send Newsletter</td>
<td>BIT</td>
<td>Y</td>
<td>Y</td>
<td>Whether or not to send emails to this customer when using mail merge feature</td></tr>

<tr class=0><td>Tax Exempt</td>
<td>BIT</td>
<td>Y</td>
<td>N</td>
<td>Whether this customer is tax exempt.</td></tr>

<tr class=0><td>Protected Acess</td>
<td>BIT</td>
<td>Y</td>
<td>N</td>
<td>Whether this customer gets special access to protected pages.</td></tr>

</table>
