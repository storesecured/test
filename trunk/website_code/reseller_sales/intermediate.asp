
<HTML>
<HEAD>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form name="frm" action="" method="post" >
	<input type="hidden" value="<%=Session("resellerid")%>" name="hidsession">
</form>

<script language="javascript">
document.frm.action = "http://managereseller.storesecured.com/reseller_home.asp"
//	document.frm.action = "../code/reselleradmin/reseller_home.asp"
	document.frm.submit()
</script>
</BODY>
</HTML>
