<!--#include file="include/header.asp"-->

<script language="JavaScript">

	function checkAddress()
	{
		selVal = "";
		for (i=0;i<document.forms['adrsel'].addrs.length;i++)
			if (document.forms['adrsel'].addrs[i].selected)
				selVal=document.forms['adrsel'].addrs[i].value;
		document.location = "<%=Switch_Name %>Modify_my_Shipping.asp?ssadr="+selVal;}

</script>
<% 
ssAdr = fn_get_querystring("ssadr")
if ssAdr="" then
    ssAdr=1
elseif ssAdr="New" then
    ssAdr=-1
end if

sShipText = sShipText & ("<table border='0' width='400'>"&_	
    "<tr><td width='400'><table border='0' width='100%'>"&_
    "<tr><td colSpan='2' width='400' class='big'><b>Modify "&Shipping&" Info</b></td></tr>"&_
    "<tr><td colSpan='2' width='400'></td></tr>"&_
    "<tr><td colSpan='2' width='400' class='normal'><b>Address Book") 
sql_select_addrs = "exec wsp_customer_lookup_no "&store_id&","&cid&",0;"
fn_print_debug sql_select_addrs
set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select_addrs,mydata,myfields,noRecords)

sShipText = sShipText & "<form name='adrsel'>"
maxrt = -1
tota = 0 

sShipText = sShipText & ("<select name='addrs' onChange=""JavaScript:checkAddress();"">")
if noRecords = 0 then
    FOR rowcounter= 0 TO myfields("rowcount")
        tota = tota + 1 
        sRecord_type = mydata(myfields("record_type"),rowcounter)
        sShipText = sShipText & ("<option value='"&sRecord_type&"'")
        fn_print_debug sRecord_type&"="&ssadr
        if isNumeric(ssadr) then
	   If cint(sRecord_type)=cint(ssadr) then
            sShipText = sShipText & ("selected")
        End If
        end if
	    sShipText = sShipText & (">"&mydata(myfields("first_name"),rowcounter)&"&nbsp;"&_
	        mydata(myfields("last_name"),rowcounter)&"&nbsp;-&nbsp;"&_
	        mydata(myfields("address1"),rowcounter)&"</option>")
    Next
End If
          
sShipText = sShipText & ("<option value='-1'")
if isNumeric(ssadr) then
	if cint(ssadr)=-1 then
	    sShipText = sShipText & "selected"
	end if
end if
sShipText = sShipText & (">Add New</option></select></form>"&_
    "</font></b></td></tr><tr>"&_
    "<td colSpan='2' width='400'></td></tr>")
if tota>1 and not (maxrt+1)=ssadr then
	allow_delete = true
else
	allow_delete = false 
end if
if ssadr<>"" then
	Record_Type = ssadr
else
	Record_Type = 1
end if
response.Write sShipText
sShipText=""
%>
<!--#include file="include/Display_Cust_Form.asp"--> 		
<%
sShipText = sShipText & ("</table></td></tr>"&_
    "<form name='Modify_Shipping' action='before_payment.asp' method=post>"&_
	"<input type='hidden' name='Return_To' value='"&fn_get_querystring("Return_To")&"'>"&_
	"<tr><td align=center>"&fn_create_action_button ("Button_image_Checkout", "Check_Out", "Check Out")&_
	"</td></tr></form></table>")
	
response.Write sShipText
%>
<!--#include file="include/footer.asp"-->
