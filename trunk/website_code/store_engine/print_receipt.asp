<!--#include file="include/header_noview.asp"-->
<base href="<%= Switch_Name %>">
<style type='text/css'>body,p,table,td,ul,ol
{font-family:verdana,helvetica,arial,sans-serif;
font-size:10pt}
</style>

<!--#include file="include/receipt_include.asp"-->
</head>
<table width=600><tr><td>
<% 
oid = fn_get_querystring("Soid")
create_receipt 1
%>
</td></tr></table>