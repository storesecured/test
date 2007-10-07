<!--#include file="include/header_noview.asp"-->
<%
if request.querystring("Error_Log") <> "" then
	Error_Log = fn_get_querystring("Error_Log")
end if
error_array = split(Error_Log,"|")	%>

<table border="0" width="100%" cellspacing="0">


	<tr>
		<td width="100%" align=center><font face="Arial" size="2">
			<B>Error</b>
			<ul>
				<% for each one_error_array in error_array %>
					<% if len(one_error_array) > 1 then %>
						<li><%= one_error_array %></li>
					<% End If %>
				<% Next %>
			</ul>
			Use your browser's back button to return to the previous page.</font>
		</td>
	</tr>

</table>

<!--#include file="include/footer.asp"-->
