<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<script language="javascript">
		function pickCustomers()
		{
			var str1;
			str1='';
			for (i=0;i<document.frmSelectCID.chk1.length;i++)
			{
			
				if (document.frmSelectCID.chk1[i].checked==true)
				{
					str1 = str1 +  document.frmSelectCID.chk1[i].value + ',';
				
				}


			}			
			str1=str1.substring(0, str1.length-1);
		//	document.frmSelectCID.lstCids.value =  str1;
		//    self.opener.document.form1.txt1.value = str1;
		self.opener.document.forms[0].Group_Cid.value = str1;
		
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
				 str1 = self.opener.document.forms[0].Group_Cid.value;
				// alert(str1);
				 CidArray = str1.split(",");
				 for (j=0; j<CidArray.length; j++) //interating through the array
				 {
					for (i=0; i < document.frmSelectCID.chk1.length; i++) //iterating through the checkboxes
					{
						if (document.frmSelectCID.chk1[i].value == CidArray[j])
						{
							document.frmSelectCID.chk1[i].checked = true;
						}
					}
				}

	}

</script>

<body onLoad="doCheck()">
<form name="frmSelectCID" >
<Style>
	.searchList
	{
		font-family:verdana;
		font-size:12;
	}
</style>
<table border=1 class="searchList">
<tr bgcolor="#eaeaea">
	<th> First Name </th>
	<th>Last Name</th>
	<th>Email</th>

	<th>Select<br><input type="checkBox" name="chkAll" onClick="changeState()"> </th>
</tr>
<%
	set rs_cust = server.createobject("ADODB.Recordset")
	rs_cust.open "select ccId, last_name, first_name, email from store_customers where record_type=0 and store_id=" &store_id&" order by last_name", conn_store,1,1

	while not rs_cust.eof
	%>
		<tr>
			<td class="inputname"><%=rs_cust.fields("first_name")%></td>
			<td class="inputname"><%=rs_cust.fields("last_name")%></td>
			<td class="inputname"><%=rs_cust.fields("email")%></td>
			<td class="inputname" align="center"><input type="checkbox" name="chk1" value="<%=rs_cust.fields("ccid")%>"></td>
		</tr>

	<%
		rs_cust.movenext
	wend
	rs_cust.close
	set rs_cust = nothing
%>
<tr>
<!--<td colspan=2>
	<input name="lstCids">
</td>
-->
<td colspan=4 align="right">
	<input type="button" onclick="pickCustomers()" Value="Select">&nbsp;&nbsp;&nbsp;
	<input type="button" onClick="window.close()" value="Close">
</td>
</form>
</body>
