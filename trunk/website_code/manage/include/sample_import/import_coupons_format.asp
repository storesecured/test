Please note:&nbsp;<br>
			Coupon Id field Is the unique coupon identifier:<br>
			If Coupon_Id does not exist in the DB a new coupon \ record will be inserted into store coupons table.<br>
			otherwise if Coupon_Id already exists in the DB the coupon \ record will be updated.&nbsp;
<BR><BR>Even if a particular column is not required you must leave the empty column there as a placeholder.
		<BR><BR>

			<table width="100%" border="1" cellspacing="0" cellpadding=2>
				
								<tr bgcolor=#DDDDDD>
<td class=0><b>Field name</b></td>
<td><b>Type/Size</b></td>
<td><b>Required?</b></td>
<td><b>Example value</b></td>
<td><b>Description</b></td></tr>

<tr class=1><td>Coupon Id</td>
<td>TEXT/255</td>
<td>Y</td>
<td>10OFF</td>
<td>The unique name to identify this coupon and what the customer will type in at checkout to receive discount</td></tr>

<tr class=0><td>Name</td>
<td>TEXT/150</td>
<td>Y</td>
<td>10% off anything</td>
<td>The internal name to remember why this coupon was created</td></tr>

<tr class=1><td>Start Date</td>
<td>Date/Time</td>
<td>Y</td>
<td>5/8/2003</td>
<td>The date this coupon starts to be valid</td></tr>

<tr class=0><td>End Date</td>
<td>Date/Time</td>
<td>Y</td>
<td>6/8/2003</td>
<td>The date that this coupon is no longer valid</td></tr>

<tr class=1><td>Type</td>
<td>BIT</td>
<td>Y</td>
<td>0</td>
<td>1=Flat Amount Off, 0=Percentage Off</td></tr>

<tr class=0><td>Coupon Amount</td>
<td>Number</td>
<td>Y</td>
<td>10</td>
<td>Percentage discount or dollar amount off depending on coupon type</td></tr>

<tr class=1><td>Total Usage</td>
<td>Number</td>
<td>Y</td>
<td>50</td>
<td>Total number of times this coupon can be used in the store by all customers</td></tr>

<tr class=0><td>Customer Usage</td>
<td>Number</td>
<td>Y</td>
<td>2</td>
<td>Total number of times this coupon can be used in the store by a single customer</td></tr>

<tr class=0><td>Purchase Min</td>
<td>Number</td>
<td>Y</td>
<td>10</td>
<td>Minimum purchase amount for coupon to apply to</td></tr>

<tr class=0><td>Purchase Max</td>
<td>Number</td>
<td>Y</td>
<td>50</td>
<td>Maximum purchase amount for coupon to apply to</td></tr>

<tr class=0><td>Discount Available on</td>
<td>Text</td>
<td>Y</td>
<td>All</td>
<td>Enter All or leave blank to indicate all skus or a comma delimited list of skus to indicate those items that can be discounted.</td></tr>


</table>
