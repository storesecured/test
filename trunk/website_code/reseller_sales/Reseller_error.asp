<%
iError = trim(Request.QueryString("error"))
select case Request.QueryString("error")
		case 1 Message_text = "This domain name cannot be registered because it is already registered.<BR><BR>If you own this domain please request a transfer instead of registration.<BR><BR>If you do not own this domain please choose another name that is not already taken."
		case 2 Message_text = "This domain name cannot be transferred because it is not yet registered."
		case 3 Message_text = "Not a valid login, please try again.<br>Use your browsers back button to return."
		case 4 Message_text = "Please start again."
end select
%>	

<table border="1" width="100%" cellspacing="0" height="100%">
	
	<tr> 
		<td width="100%" bgcolor="Red"><font face="Arial" size="2" color="White"><b><%= Error_T %>&nbsp;</b></font></td>
	</tr>
    
	<tr>
		<td width="100%" height="100%" valign="center" align="center"><font face="Arial" size="2">    
			<p><b><%= Message_text %></b></p></font></td>
    </tr>

    <tr bgcolor="Red">
	<th>
		<font face='arial' size='2' color='#ffffff'><%= Click_anywhere_in_this_window_to_close_it_T %></font>
	</th>
	</tr>
 
</table>

