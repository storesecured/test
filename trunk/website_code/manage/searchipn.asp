<!--#include virtual="common/connection.asp"-->

<%
  store_id=request.form("store_id")
  order_id=request.form("order_id")


  
  



'***************************************




%>
<html>
<head>
<script language="javascript">
<!--
function validate()
{
	
	if(document.myform.store_id.value=="")
	{
		alert("please enter Store Id");
		document.myform.store_id.focus();
		return false;
	}

	return true;
	
}
//-->
</script>

<title>
Paypal IPN search
</title>
</head>
<body>
<form name="myform" method="post" action="searchipn.asp" onsubmit="return validate()">
<table width="679" height="154">
            <tr bgcolor='#FFFFFF'>
              <TD bgcolor=#6495ED width=673 colspan=2 height="9">
              <font color="#FFFFFF"><b>&nbsp;Searching for PayPal IPN messages:</b></font></TD></TR>
            <tr bgcolor='#FFFFFF'>
				<td colspan=2 height="1" width="673"></td></tr>
 
				<tr bgcolor='#FFFFFF'>
				<td class="inputname" height="1" width="119">Store Id</td>
				<td class="inputvalue" height="1" width="550"><input type="text" name="store_id" size="20">
			</td>
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td class="inputname" height="26" width="119">Order Id</td>
					<td class="inputvalue" height="26" width="550"><input type="text" name="order_id" size="20">
				</tr>
				<tr bgcolor='#FFFFFF'>
					<td colspan=2 align=center height="1" width="673">
						<p>
						<input type="submit" class="Buttons" value="Search" name="search IPN"></td>
				</tr>

				
				
				
			</table>
			</form>
				<hr width="678" align="left">

 <%
 if request.form("store_id")<>"" then
 set rs_Store =  server.createobject("ADODB.Recordset")
  sql_select="select * from Store_Ipn where store_id=" & store_id 
  if order_id<>"" then
  sql_select=sql_select & " and OID=" & order_id
  end if
  rs_Store.open sql_select,conn_store,1,1
 %>

  <table align="left" border="1" borderColor="blue" cellPadding="0" cellSpacing="0" width="678" height="29">
    <tr bgcolor="#6495ed">
     <th width="65" height="27">
      <p align="left"><font color="#FFFFFF" size="2">Order Id</font></th>
    <th width="120" height="27"><font color="#FFFFFF" size="2">Date / Time</font></th>
    <th width="201" height="27"><font color="#FFFFFF" size="2">IPN message</font></th>
     <th width="87" height="27"><font color="#FFFFFF" size="2">Payment Type</font></th>
     <th width="92" height="27"><font color="#FFFFFF" size="2">Payment Status</font></th>
      <th width="93" height="27"><font color="#FFFFFF" size="2">Pending Reason</font></th>
        </tr>
   <%
   
    flag=1

 do while not rs_Store.eof 

if flag=1 then
coloro="white"
else
coloro="#deedfc"
end if

Response.Write("<font size='2' face='verdana'><tr align='left'><TD height='20' bgcolor=" & coloro & ">&nbsp;" & rs_Store("OID") & "</TD><b><font size='2' face='verdana'><TD height='20' bgcolor=" & coloro & ">&nbsp;" & rs_Store("sys_created")& "</TD><TD height='20' bgcolor=" & coloro & ">" & rs_Store("ipn_post")& "</TD><TD height='20' bgcolor=" & coloro & ">" & rs_Store("payment_type")& "</TD><TD height='20' bgcolor=" & coloro & ">" & rs_Store("payment_status") & "</TD><TD height='20' bgcolor=" & coloro & ">" & rs_Store("pending_reason")& "</TD></tr>")

flag=-flag
rs_Store.MoveNext
loop
rs_store.Close
  %>
  </table>
</div>
    <% end if %>

</body>
</html>
