
<%
if fn_get_querystring("Error_Log") <> "" then
	Error_Log = fn_get_querystring("Error_Log")
end if
error_array = split(Error_Log,"|")	%>

<table border="1" width="300" cellspacing="0">
	<tr> 
		<td width="100%" bgcolor="Red"><font face="Arial" size="2" color="White"><b>Error </b></font></td>
	</tr>

	<tr>
		<td width="100%"><font face="Arial" size="2">		
			<ul>
				<% for each one_error_array in error_array %>
					<% if len(one_error_array) > 1 then %>
						<li><%= one_error_array %></li>
					<% End If %>
				<% Next %>
			</ul>
			Use your browsers back button to return.</font> 
		</td>
	</tr>

</table>

