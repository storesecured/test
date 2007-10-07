<!--#include file="include/header.asp"-->

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td width="100%">
         <% if Overdue_Payment > 7 then %>
         This store has been administratively closed.<BR><BR>
         If you are the store owner please login to your store management area for more information.
         <% end if %>
         </td>
		</tr>
	</table>
<!--#include file="include/footer.asp"-->
