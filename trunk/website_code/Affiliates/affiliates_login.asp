<!--#include virtual="common/connection.asp"-->

<%

If Request.QueryString("Store_id") = "" then
	Store_id = ""
else
	Store_id = Request.QueryString("Store_id")
end if

Store_Language = session("Store_Language")
%>

<html>

<head>
	<meta name="Pragma" content="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>ESalesWizard E-Commerce Manager</title>
</head>

<SCRIPT LANGUAGE="JavaScript">

	function surfto(form) {
		var myindex=form.Language_name.selectedIndex
		if (form.Language_name.options[myindex].value != "0") {
			location=form.Language_name.options[myindex].value;}}

</SCRIPT>

<body text="#000066" >

<table border="1" cellpadding="0" cellspacing="0" width="50%" bgcolor="#FFFFFF" bordercolor="#000066">

	<tr>
		<td width="100%">
			<table border="0" width="100%" cellspacing="0">
				
				<tr>
					<td width="100%" colspan="2">
						<table border="0" width="100%" bordercolor="#000066">
								
							<tr>
									<td width="99%" colspan="3" valign="top">
									<font face="Arial" size="2">
										<b>Affiliates Login</b><br>&nbsp;</font></td>
								
								</tr>

						<form method="POST" action="affiliates_login_action.asp" target="_top" >						
						<% if Store_id = "" then %>
						<tr>
								<td width="1%" height="1" nowrap>&nbsp;</td>
							<td width="1%" height="1" nowrap><font face="Arial" size="2">Select Store</font></td>
								<td width="98%" height="1" valign="top" colspan="2">
								
									<select name="AFFILIATE_STORE">
										<% sql_sels = "select store_id, store_name from store_settings where store_id<>101 order by store_name" %>
										<% rs_store.open sql_sels, conn_store, 1, 1 %>
										<% do while not rs_store.eof %>
											<option value="<%= rs_store("Store_id") %>"><%= rs_store("Store_Name") %></option>
											<% rs_store.movenext %>
										<% loop %>
										<% rs_store.close %>
									</select>
								
							</td>
							</tr>
							<% else %>
								<input type="hidden" name="AFFILIATE_STORE" value="<%= Store_id %>">
							<% end if %>
							
						<tr>
							<td width="1%" height="1" nowrap>&nbsp;</td>
							<td width="1%" height="1" nowrap><font face="Arial" size="2">Login</font></td>
								<td width="98%" height="1" valign="top" colspan="2">
								<input type="text" name="AFFILIATE_CODE" size="20" value="<%= request.form("LOGIN_ID") %>"></td>
						</tr>
							
						<tr>
								<td width="1%" height="1" nowrap>&nbsp;</td>
								<td width="1%" height="1" nowrap><font face="Arial" size="2">Password</font></td>
								<td width="98%" height="1" valign="top" colspan="2">
								<input type="password" name="AFFILIATE_PASSWORD" size="20" value="<%= request.form("LOGIN_PASSWORD") %>"></td>
							</tr>
							
						<tr>
								<td width="1%" height="1" nowrap>&nbsp;</td>
								<td width="1%" height="1" nowrap>&nbsp;</td>
								<td width="84%" height="1" align="left" valign="top"><br>
								<input type="submit" class="Buttons" value="Login" name="I1"></td>
								<td width="84%" height="1" align="left" valign="top">&nbsp;</td>
						</tr>
						</form>
						</table>
					</td>
				</tr>
			
				
			</table>
		</td>
	</tr>
</table>


</body>
</html>
