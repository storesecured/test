<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page allows reseller to set the minimum balance.
'	Page Name:		    minimum_balance.asp
'	Version Information:EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    minimum_balance.asp
'	Date & Time:			19 Jan 2005 		
'	Created By:			Suman Sharma
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
'mSel7=3
'mSel8=2
%>
<HTML>
	<HEAD>
		<TITLE>Reseller Admin</TITLE>
		<META content=no-cache name=Pragma>
		<META content=no-cache http-equiv=pragma>
		<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
		<LINK href="images/style.css" rel=stylesheet type=text/css>
		<script src="include/script.js" language="JavaScript" type="text/javascript"></script>
		<SCRIPT LANGUAGE="JavaScript" SRC="include/CalendarPopup.js"></SCRIPT>
		<SCRIPT language=JavaScript src="include/commonfunctions.js" ></script>
		<META content="MSHTML 5.00.3813.800" name=GENERATOR>
		<script language="javascript">
			function fnValidate(){		
				var amt;
				var ErrMsg = "";
				amt = document.frm.txtMinBal.value;				
				if (amt != ""){				
					if (!isAllNumeric(amt) || (amt<=0)){
							ErrMsg = ErrMsg + "Please enter valid amount.\n";			
						}
				}else{
							ErrMsg = ErrMsg + "Please enter min amount.\n";
				}
				if(ErrMsg != ""){
						alert(ErrMsg);
				}else{
					document.frm.action = "minimum_balance.asp?flag=1";
					document.frm.submit();
				}
			}
		</script>
		</script>
	</HEAD>

<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<form name="frm" method="post">
<DIV id=overDiv style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  <TBODY>
		<TR>
			<TD class=title>
				<TABLE border=0 cellPadding=0 cellSpacing=0>
					<TBODY>
						<TR>
							<TD class=title><B>Reseller</B></TD>
							<TD class=special width=200>&nbsp;</TD>
							<TD align=left class=special width="60%"><UL><BR><BR></UL></TD>
						</TR>
					</TBODY>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD>
				<!--#include file="incmenu.asp"-->   
       </TD>
		</TR>
		<TR>
			<TD>
	      <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
		      <TBODY>
				    <TR vAlign=top>
						  <TD rowSpan=2 width=180>								
								<TABLE border=0 cellPadding=0 cellSpacing=0 width=150>
									<TBODY>       
										<TR>
											<TD class=meniu height=20 
											onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
											onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
											<A class=b href="Reseller_sales_history.asp">View Sales History</A>
											</TD>
										</TR>
										<TR>
											<TD class=meniu height=20 
											onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
											onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
											<A class=b href="Reseller_payment_Mode.asp">Choose Payment Mode</A>
											</TD>
										</TR>         
										<TR>
											<TD class=meniu height=20 
											onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
											onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
											<A class=b href="edit_reseller.asp">Edit Profile</A>
											</TD>
										</TR>  
										<TR>
											<TD class=meniu height=20 onmouseout="style.backgroundColor='#EBF9D8', 			style.color='#8FB25E'" onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
											<A class=b href="minimum_balance.asp">Set Minimum Balance</A>
											</TD>
										</TR>          
									</TBODY>
								</TABLE>
<!--								</form>-->
							</TD>
							<TD height=15 vAlign=top width=570></TD>
						</TR>
	<%
		Dim fMinAmt
		Dim rsMinAmt, rsSaveAmt

		intResellerId = Trim(session("ResellerID"))
		Set rsMinAmt = conn.execute("select fld_min_amt from Tbl_Reseller_Master where fld_reseller_id="&intResellerId)
		
		If Not rsMinAmt.Eof Then
			fMinAmt = Trim(rsMinAmt("fld_min_amt"))
		End If

		If Trim(Request("flag")) = "1" Then		
			fMinAmt = ""
			fMinAmt = Trim(Request("txtMinBal"))
			Set rsSaveAmt = Server.CreateObject("adodb.recordset")
			With rsSaveAmt
				.Source           = "sp_setminamt " & intResellerId & "," &fMinAmt
				.ActiveConnection = conn
				.CursorType       = adOpenForwardOnly
				.LockType         = adLockReadOnly				
				.Open
			End With
			response.redirect"minimum_balance.asp"
		End If
	
	
	%>
						<TR>
							<TD class=pagetitle height=400 vAlign=top>
			          <table border="0" cellspacing="0" cellpadding="1" align="center">
									<tr><br>Set Minimum Balance<hr noshade></tr>          
									<tr>
										<td></td>		         
										<td>		
											<tr bgcolor="#ffffff"> 
												<td width="20%" height="28">
													<font face="Arial, Helvetica, sans-serif" size="2">
														&nbsp;Minimum Balance&nbsp;
													</font>
												</td>
										</td>
										<td width="83%" height="28">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											$<input type="text" name="txtMinBal" maxlength=6 size=6 value="<%=fMinAmt%>">
										</td>	
									</tr>						
									<tr>			
										<td colspan="2"><br><input name="btnSave" type="button" value="Save" onclick="javascript:fnValidate();"><input type="reset" value="Reset" name=reset>
										</td>											
								  </tr>          
			          </table>
							</td>
						</tr>						  
						<DIV ID="testdiv1" 		STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;">
						</DIV>        
					</TBODY>
				</TABLE>
			</TD>
		</TR>
	</TBODY>
</TABLE>
</form>
</BODY>
</HTML>
