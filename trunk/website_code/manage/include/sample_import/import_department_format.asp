			Please note:&nbsp;<br>
			Department_ID field  Is the unique department identifier:<br>
			If Department_ID does not exist in the DB a new department \ record will be inserted into store_dept table.<br>
			otherwise if Department_ID already exists in the DB the department \ record will be updated.
			
			<BR><BR>Even if a particular column is not required you must leave the empty column there as a placeholder.
		<BR><BR>

         <table width="100%" border="1" cellspacing="0" cellpadding=2>
					<tr bgcolor=#DDDDDD>
						<td class=0><b>Field name</b></td>
						<td><b>Type/Size</b></td>
						<td><b>Required</b></td>
						<td><b>Example value</b></td>
						<td><b>Description</b></td>
					</tr>
					<tr class=1>
						<td>Department ID</td>
						<td>NUMBER</td>
						<td>N</td>
						<td>4</td>
						<td>The unique Department Id (must be left blank for new)</td>
					</tr>
					<tr class=0>
						<td>Department Name</td>
						<td>TEXT/50</td>
						<td>Y</td>
						<td>Mens</td>
						<td>The Department name</td>
					</tr>
					<tr class=1>
						<td>Department Description</td>
						<td>TEXT/2500</td>
						<td>N</td>
						<td>Various items for men</td>
						<td>A brief description of the department</td>
					</tr>
					<tr class=0>
						<td>Parent</td>
						<td>TEXT/100</td>
						<td>N</td>
						<td>Clothing</td>
						<td>The parent departments or to make this a top level department use "<B>Main Department</B>"</td>
					</tr>
					<tr class=1>
						<td>Image Path</td>
						<td>TEXT/100</td>
						<td>N</td>
						<td>redshirt.gif</td>
						<td>Name of gif or jpeg image file</td>
					</tr>

					<tr class=0>
						<td>Department Html</td>
						<td>TEXT/3000</td>
						<td>N</td>
						<td>any html</td>
						<td>Html to be displayed at the top of the page</td>

					<tr class=1>
						<td>Visible</td>
						<td>BIT</td>
						<td>N</td>
						<td>Y</td>
						<td>Should this department be shown?</td>
						
					<tr class=0>
						<td>View Order</td>
						<td>NUMBER</td>
						<td>N</td>
						<td>1</td>
						<td>The display order for the department</td>
					</tr>
					<tr class=1>
						<td>Meta Keywords</td>
						<td>TEXT/1500</td>
						<td>N</td>
						<td>collared shirt, mens shirt</td>
						<td>The meta keywords to use for this department</td>
					</tr>
					<tr class=0>
						<td>Meta Description</td>
						<td>TEXT/1500</td>
						<td>N</td>
						<td>Lots of mens shirts</td>
						<td>The meta description to use for this department</td>
					</tr>
					<tr class=1><td>Protect Page</td>
						<td>BIT</td>
						<td>N</td>
						<td>Y</td>
						<td>Only show to certain users with protected access</td>
					</tr>
					<tr class=0>
						<td>Show Name</td>
						<td>BIT</td>
						<td>N</td>
						<td>Y</td>
						<td>Show the name of the department at the top of the page</td>
					</tr>
					<tr class=1>
						<td>Meta Title</td>
						<td>TEXT/100</td>
						<td>N</td>
						<td>M Title</td>
						<td>Fill with Meta_Title for department page</td>
					</tr>
					
					<tr class=0>
						<td>Department Html Bottom</td>
						<td>TEXT/8000</td>
						<td>N</td>
						<td>any html</td>
						<td>Html to be displayed at the bottom of the page</td>
					</tr>
				</table>
