<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page provides facility for reseller to enter his own website's name. 
'	Page Name:		   	reseller_display_name.asp
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_display_name.asp
'	Date & Time:		14 Jan 2005 		
'	Created By:			Suman Sharma
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 


%>


<HTML>
	<HEAD>
		<TITLE>Reseller Admin</TITLE>
			<META content=no-cache name=Pragma>
			<META content=no-cache http-equiv=pragma>
			<META content="text/html; charset=windows-1252" http-equiv=Content-Type>
			<LINK href="images/style.css" rel=stylesheet type=text/css>
			<script language="javascript" src="include/commonfunctions.js"></script>
			<META content="MSHTML 5.00.3813.800" name=GENERATOR>
			<script>
				function fnSave(){					
					document.frmName.action = "Reseller_display_name.asp?flag=1"			
					document.frmName.submit();											
				}
			</script>
	</HEAD>
<%
	Dim strName, intResellerId, strOrgName
	Dim intFlag
	intFlag = ""
	
	intResellerId = Trim(session("ResellerID"))
	Set rsRetriveName = conn.execute("select fld_Display_Name from Tbl_Reseller_Master where fld_reseller_id="&intResellerId)
	If Not rsRetriveName.Eof Then
		strName = trim(checkencode(rsRetriveName("fld_Display_Name")))
		'response.end
	End If
	
	'Storing the name into the table:TBL_ResellerMaster
	If Trim(Request.querystring("flag")) = "1" Then	
		strName=""
		strName = fixQuotes(Trim(Request("txtName")))
		Set rsSaveName = Server.CreateObject("adodb.recordset")
			With rsSaveName
				.Source           = "sp_saveinfo " & intResellerId & ",' "& strName &" '"
				.ActiveConnection = conn
				.CursorType       = adOpenForwardOnly
				.LockType         = adLockReadOnly						
				.Open
			End With

		'	Set rsSaveName = nothing
		'	rsSaveName.close
		response.redirect "reseller_display_name.asp"
	End If

%>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<DIV id=overDiv 
	style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000">
</DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
	<FORM action="" method=post name="frmName">
  <TBODY>
		<TR>
			<TD class=title>
				<TABLE border=0 cellPadding=0 cellSpacing=0>
					<TBODY>
						<TR>
							<TD class=title><B>Reseller</B></TD>
							<TD class=special width=200>&nbsp;</TD>
							<TD align=left class=special width="60%">
								<UL><BR><BR></UL>
							</TD>
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
												<A class=b href="Reseller_change_Logo.asp">Change Logo</A>
											</TD>
										</TR>
										<TR>
												<TD class=meniu height=20 
													onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
													onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
													<A class=b href="Reseller_change_HKeywords.asp">Manage Keywords</A>
												</TD>
										</TR>              
										<TR>
												<TD class=meniu height=20 
													onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
													onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
													<A class=b href="Reseller_change_Plan.asp">Change Plan Pricing</A>
												</TD>
										</TR>             
										<TR>
												<TD class=meniu height=20 
													onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
													onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
													<A class=b href="Reseller_change_Desc.asp">Define Description</A>
												</TD>
										</TR>          
										<TR>
												<TD class=meniu height=20 
													onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
													onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
													<A class=b href="Reseller_display_name.asp">Choose Name</A>
												</TD>
										</TR>
										<TR>
											<TD class=meniu height=20 
												onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
												onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'">
												<A class=b href="Reseller_contact_us.asp">Contact us</A>
											</TD>
									 </TR>
									</TBODY>
								</TABLE>
							</TD>
							<TD height=15 vAlign=top width=570></TD>
						</TR>
						<TR>
							<TD class=pagetitle height=400 vAlign=top>Choose Display Name
								<TABLE align=left border=0 cellPadding=0 cellSpacing=0 
									width="90%"> 
									
										<TBODY>
											<tr><br>
											 <td>Name:  <INPUT name="txtName" maxlength="100" value="<%=strName%>"></td>
											</tr>	
											<tr>
												<td><br>
													 <INPUT name="save" type="button" value="Save" onclick="javascript:fnSave()"> 
													 <input type="reset" value="Reset" >
												</td>
											</tr>		            
				             <!-- <TBODY>-->
												<TR>				
													<TD height=20></TD>
												</TR>
											</TBODY>
										</TABLE>
									</TD>
								</TR>
					  </TBODY>
					</TABLE>
				</TD>
			</TR>
		</TBODY>
	</TABLE>
</FORM>
</BODY>
</HTML>
