<!--#include file = "include/ESCheader.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441

msel13=3
msel14=2
%>
<HTML><HEAD><TITLE>Easystorecreator</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<script language=javascript src="../include/commonfunctions.js"></script>
<script language="javascript">
function fnSave()
{
	
	var ErrMsg
	ErrMsg = ""
	
	//check for only characters and numeric
	if (isWhitespace(document.frm.txtstoreid.value) == true)
	{
		ErrMsg = ErrMsg + "Store ID is mandatory.\n"
	}
	
	if (isWhitespace(document.frm.txtstoreid.value) == false)
	{
		if( isAllNumeric(document.frm.txtstoreid.value) == false)
			{	
				ErrMsg=ErrMsg + "Enter valid Store ID.\n"
			}
	}		
	
	if (isWhitespace(document.frm.txtamount.value) == true)
	{
		ErrMsg = ErrMsg + "Amount is mandatory.\n"	
	}
	if (isWhitespace(document.frm.txtamount.value) == false)
	{
		if( isAllNumeric(document.frm.txtamount.value) == false)
			{	
				ErrMsg=ErrMsg + "Enter valid amount.\n"
			}
	}
	if (ErrMsg!="")
	{
		alert(ErrMsg)
	}
	else
	{	
		document.frm.action = "refund_request.asp?save=yes"
		document.frm.submit()
	}
}
function sel()
{

		strAction = "selectstore.asp"
		
		var lWindow;
		var iid;
							
		var l = screen.width - 700
		//window.open(strAction)
		
		lWindow=window.open(strAction,"","height=600,width=700,toolbar=yes,menubar=yes,scrollbars=yes");
		//lWindow.focus()
}

</script>
	

<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
'code here gets execute when the admin choose a particular store id 
'****************************************************************************
if trim(Request.QueryString("save")) = "yes" then
	
	' code here to retrive the amount and storeid 
	'****************************************************************************								
	store_id = trim(Request.Form("txtstoreid"))
	amount = trim(Request.Form("txtamount"))
	'****************************************************************************
	intflag = "0"
		
	'code here to check whether the request is alrad placed or not 
	'****************************************************************************
	 sqlcheck = " select fld_customer_id from tbl_request where  fld_customer_id="&store_id 
	 set rscheck = conn.execute(sqlcheck)
	 if not rscheck.eof then
		intflag="4"
	 end if
	 
	 '****************************************************************************
	
	'code here to check whether the admin had entered the valid amount or not
	'****************************************************************************
	sql_select = "select sys_billing.amount,sys_payments.transdate,sys_payments.payment_term from sys_payments inner join sys_billing on sys_payments.store_id = sys_billing.store_id where sys_payments.store_id='"&store_id&"' order by sys_payments.transdate desc"
	
	set rs_store = conn.execute(sql_select)
	
		
	if not rs_store.eof then 
		checkamount = trim(rs_store("amount"))
		if csng(checkamount)<>csng(amount) then
			intflag="2"
		end if 
		
	else
		intflag="3"
	end if
	'****************************************************************************
	'Response.End
	
	
	if 	intflag = "0" then 
		'code here to put the reuest into the table
		'****************************************************************************
		sql = "insert into tbl_request (fld_customer_id,fld_amount,fld_refund_status) values('"&store_id&"','"&amount&"',0)"
		conn.execute(sql)
		intflag = "1"
		'****************************************************************************
	end if	
end if
'****************************************************************************

%>
		<DIV id="overDiv" style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
		<TABLE border="0" cellPadding="0" cellSpacing="0" width="750" align="center">
			<TBODY>
				<TR>
					<TD class="title">
						<TABLE border="0" cellPadding="0" cellSpacing="0">
							<TBODY>
								<TR>
									<TD class="title"><B>Easystorecreator</B></TD>
									<TD class="special" width="200">&nbsp;</TD>
									<TD align="left" class="special" width="60%">
										<UL>
											<BR>
											<BR>
										</UL>
									</TD>
								</TR>
							</TBODY></TABLE>
					</TD>
				</TR>
				<TR>
					<TD>
				<!--#include file="incESCmenu.asp"-->
				<TR>
					<TD>
						<TABLE border="0" cellPadding="0" cellSpacing="0">
							<TBODY></TBODY></TABLE>
					</TD>
				</TR>
				<TR>
					<TD>
						<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%">
							<TBODY>
								<TR vAlign="top">
									<TD rowSpan="2" width="180">
										<form name="frm" method="post" action="">
											<TABLE border="0" cellPadding="0" cellSpacing="0" width="150">
												<TBODY></TBODY></TABLE>
									</TD>
									<TD height="15" vAlign="top" width="570"></TD>
								</TR>
								<TR>
									<TD class="pagetitle" height="400" vAlign="top">
										<table border="1" id="TABLE1" width="100%">
											<tr>
												<td></td>
												<td>
												</td>
											</tr>
											<tr>
												Accept Requests
												<hr noshade>
											</tr>
											<%if intflag="4" then%>
											<tr>
												<td></td>
												<td><font color="red" size="1">Request is already placed.</font></td>
											</tr>
											<%elseif intflag="2" then%>
											<tr>
												<td></td>
												<td><font color="red" size="1">Invalid Amount</font></td>
											</tr>
											<%elseif intflag="3" then%>
											<tr>
												<td></td>
												<td><font color="red" size="1">Invalid Store ID</font></td>
											</tr>
											<%end if%>
											<tr>
												<td>Store Id</td>
												<td><input type="text" name="txtstoreid" maxlength="10">&nbsp;<a href="javascript:sel()">Select 
														store</a></td>
											</tr>
											<tr>
												<td width="30%">Amount</td>
												<td><input type="text" name="txtamount" value="" maxlength="10"></td>
											</tr>
											<TBODY>
												<tr>
													<td>
														<INPUT name="Save" type="button" value="Submit" onclick="javascript:fnSave()"> <INPUT name="Reset" type="Reset" value="Reset"></td>
												</tr>
												<tr>
													<td></td>
													<td>
													</td>
													<td>
													</td>
													<td align="Right" valign='Right'><font face="arial" size="2"></font></td>
												</tr>
										</table>
										</FORM>
										<TABLE align="left" border="0" cellPadding="0" cellSpacing="0" width="90%">
											<tr>
												<td align="middle"><font size="1" face="arial"></font></td>
											</tr>
											<TR>
												<TD height="20"></TD>
											</TR>
							</TBODY></TABLE>
					</TD>
				</TR>
			</TBODY></TABLE>
		</TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM> </table>
	</BODY>
</HTML>
<%if intflag="1" then%>
<script language="javascript">
	alert("The request has been saved ")
	document.location.href = "refund_request.asp"
</script>
<%end if%>
