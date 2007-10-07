
<html>

<body topmargin='0' leftmargin='0' rightmargin='0' bottommargin='0' marginheight='0' marginwidth='0'>
	<% if request.queryString("returnArg")<>"" then %>
		<img src="images/themes/<%= request.queryString("returnArg") %>_large.gif">
	<% End If %>
</body>
</html>
