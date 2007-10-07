<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<script language="javascript">
		function pickCustomers()
		{
			var str1,frmLen;
			str1='';
			frmLen = document.frmSelectCID.length
			
			if (frmLen > 4) 
			{
				for (i=0;i<document.frmSelectCID.chk1.length;i++)
				{
					if (document.frmSelectCID.chk1[i].checked==true)
					{
						str1 = str1 +  document.frmSelectCID.chk1[i].value + ',';
					
					}
				}			
			}
			else
			{
				if (document.frmSelectCID.chk1.checked == true)
				{
				str1 = str1 +  document.frmSelectCID.chk1.value + ',';
				}
			}
			
			str1=str1.substring(0, str1.length-1);
			self.opener.document.forms[2].Group_id.value = str1;
			window.close();
		}
		

	function changeState()
	{
		if (document.frmSelectCID.chkAll.checked == true)
		{
			for (i=0; i < document.frmSelectCID.chk1.length; i++)
			{
				document.frmSelectCID.chk1[i].checked = true;
			}
		}
		else
		{
			for (i=0; i < document.frmSelectCID.chk1.length; i++)
			{
				document.frmSelectCID.chk1[i].checked = false;
			}
		}
	}


	function doCheck()
	{
				 str1 = self.opener.document.forms[2].Group_id.value;
				 frmName = self.opener.document.forms[2].name
				 CidArray = str1.split(",");
				 var frmLen
				 frmLen=document.frmSelectCID.length
				 for (j=0; j<CidArray.length; j++) //interating through the array
				 {
					if (frmLen > 4)
					{
						for (i=0; i < document.frmSelectCID.chk1.length; i++) //iterating through the checkboxes
						{
							if (document.frmSelectCID.chk1[i].value == CidArray[j])
							{
								document.frmSelectCID.chk1[i].checked = true;
							}
						}
					}
					else
					{
						if (document.frmSelectCID.chk1.value == CidArray[j])
							{
								document.frmSelectCID.chk1.checked = true;
							}
					}
				}
	}


</script>

<form name="frmSelectCID" >
<Style>
	.searchList
	{
		font-family:verdana;
		font-size:12;
	}
</style>

<table border=1 class="searchList" width=50%>
<tr bgcolor="#eaeaea">
	<th nowrap> Group Name </th>
	<th nowrap> Budget Min. </th>
	<th nowrap> Purchase Min. </th>
	<th nowrap>Select<br><input type="checkBox" name="chkAll" onClick="changeState()"> </th>
</tr>
<%

	
	set rs_cust = server.createobject("ADODB.Recordset")
	rs_cust.open "select group_Id, group_name,group_budget_min,group_purchase_history,group_company from Store_Customers_Groups where Store_id = "&Store_id&" order by group_name", conn_store,1,1
	
	while not rs_cust.eof
	%>
		<tr>
			<td class="inputname"><%=rs_cust.fields("group_name")%></td>
			<td class="inputname"><%=rs_cust.fields("group_budget_min")%></td>
			<td class="inputname"><%=rs_cust.fields("group_purchase_history")%></td>
			<td class="inputname" align="center">
			<input type="checkbox" name="chk1" value="<%=rs_cust.fields("group_Id")%>">
			</td>
		</tr>

	<%
		rs_cust.movenext
	wend
	rs_cust.close
	set rs_cust = nothing
%>
<tr>

<td colspan=4 align="right">
	<input type="button" onclick="pickCustomers()" Value="Select">&nbsp;&nbsp;&nbsp;
	<input type="button" onClick="window.close()" value="Close">
</td>
</form>
<script language="javascript">doCheck();</script>