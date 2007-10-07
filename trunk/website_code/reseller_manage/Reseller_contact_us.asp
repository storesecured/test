<!--#include file = "include/header.asp"-->

<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page provides facility for reseller to enter his own address. 
'	Page Name:		   	reseller_contact_us.asp
'	Version Information: EasystoreCreator
'	Input Page:		    reseller_home.asp
'	Output Page:	    reseller_display_name.asp
'	Date & Time:		  17 Jan 2005 		
'	Created By:			  Suman Sharma
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
			<script language="javascript" src="include/CommonFunctions.js"></script>
			<META content="MSHTML 5.00.3813.800" name=GENERATOR>
			<script>
				function fnSave(){
					var ErrMsg
					ErrMsg=""

					//Validations for blank fields					
/*					if (isWhitespace(document.frmName.txtcompany.value)== false)
						{
							if (isAlphaNumeric(document.frmName.txtcompany.value)==false)
							{
								ErrMsg = ErrMsg + "Company should not contain special characters.\n" ;		
							}	
						}
					if (isWhitespace(document.frmName.txtAddress.value)== true)
						{
							ErrMsg = ErrMsg + "Address is mandatory. \n" ;		
						}
					if (isWhitespace(document.frmName.txtCity.value)== true)
						{
							ErrMsg = ErrMsg + "City is mandatory. \n" ;		
						}
					if (isWhitespace(document.frmName.txtCity.value)== false)
						{
							if(isAllCharacters(document.frmName.txtCity.value)==false)
							{
								ErrMsg = ErrMsg + "City should contain characters only.\n" ;		
							}	
						}
					if (isWhitespace(document.frmName.txtEmail.value)== true)
						{
							ErrMsg = ErrMsg + "Email is mandatory. \n" ;		
						}
					if (isWhitespace(document.frmName.txtEmail.value)== false)
						{
							if(IsEmail(document.frmName.txtEmail.value)==false)
							{
								ErrMsg = ErrMsg + "Invalid Email id. \n" ;		
							}	
						}*/
if (isPhone(document.frmName.txtZipCode.value)== false)
						{
							ErrMsg = ErrMsg + "Zip should be valid. \n" ;		
						}					

						
						if(isPhone(document.frmName.txtPhone.value)== false)
							{
								ErrMsg = ErrMsg + "Phone should be valid. \n" ;		
							}
					
						
						if(isPhone(document.frmName.txtAltNum.value)== false)
							{
								ErrMsg = ErrMsg + "Alternate Phone should be valid. \n" ;		
							}
if (isPhone(document.frmName.txtFax.value)== false)
						{
							ErrMsg = ErrMsg + "Fax should be valid. \n" ;		
						}
if (isPhone(document.frmName.txtCalltoll.value)== false)
						{
							ErrMsg = ErrMsg + "Toll Free Number should be valid. \n" ;		
						}
/*
/*


					

					if(document.all["txtZipCode"].value != "")
					{
						 
						if((IsPhone(document.all["txtZipCode"].value) || isAlphaNumeric(document.all["txtZipCode"].value)||isSpecialCharPhoneNumber(document.all["txtZipCode"].value))  == false)
						{
							ErrMsg=ErrMsg + "Invalid Zip Code. \n";
						}
						
					}

					if(document.frmName.txtPhone.value != "")
					{
						 
						if((IsPhone(document.frmName.txtPhone.value) || isAlphaNumeric(document.frmName.txtPhone.value)||isSpecialCharPhoneNumber(document.frmName.txtPhone.value))  == false)
						{
							ErrMsg=ErrMsg + "Invalid Phone Number. \n";
						}
						
					}//txtAltNum

					if(document.all["txtAltNum"].value != "")
					{
						 
						if((IsPhone(document.all["txtAltNum"].value) || isAlphaNumeric(document.all["txtAltNum"].value)||isSpecialCharPhoneNumber(document.all["txtAltNum"].value))  == false)
						{
							ErrMsg=ErrMsg + "Invalid Alternate Phone Number. \n";
						}
						
					}//txtAltNumtxtFax

					
					if(document.all["txtFax"].value != "")
					{
						 
						if((IsPhone(document.all["txtFax"].value) || isAlphaNumeric(document.all["txtFax"].value)||isSpecialCharPhoneNumber(document.all["txtFax"].value))  == false)
						{
							ErrMsg=ErrMsg + "Invalid Fax Number. \n";
						}
						
					}//txtAltNumtxtFax txtCalltoll

					if(document.all["txtCalltoll"].value != "")
					{
						 
						if((IsPhone(document.all["txtCalltoll"].value) || isAlphaNumeric(document.all["txtCalltoll"].value)||isSpecialCharPhoneNumber(document.all["txtCalltoll"].value))  == false)
						{
							ErrMsg=ErrMsg + "Invalid Toll Free Number. \n";
						}
						
					}//txtAltNumtxtFax txtCalltoll*/


							if (ErrMsg!="")
							{
								alert(ErrMsg);
							}
							else
							{
								document.frmName.action = "Reseller_contact_us.asp?flag=1"			
								document.frmName.submit();				
							}
				}
			</script>
	</HEAD>
<%
	on error goto 0
	Dim intResellerId, intZip, intPhone, intCountry, intAltNum
	Dim strCompName, strAdd, strCity, strState, strfax, strEmail, strCallToll

	intResellerId = Trim(session("ResellerID"))
	strSql="select fld_company_name, fld_address, fld_city, fld_state, fld_zip_code, fld_phone, fld_alt_phone, fld_fax, fld_country, fld_email, fld_calltoll from Tbl_Reseller_Address where fld_reseller_id="&intResellerId
	Set rsRetrieveAdd = conn.execute(strSql)

	
	If Not rsRetrieveAdd.Eof Then
		strCompName = checkencode(Trim(rsRetrieveAdd("fld_company_name")))
		strAdd = checkencode(Trim(rsRetrieveAdd("fld_address")))
		strCity = checkencode(Trim(rsRetrieveAdd("fld_city")))
		strState = checkencode(Trim(rsRetrieveAdd("fld_state")))
		intZip = Trim(rsRetrieveAdd("fld_zip_code"))
		intPhone = Trim(rsRetrieveAdd("fld_phone"))
		intAltNum = Trim(rsRetrieveAdd("fld_alt_phone"))
		strfax = Trim(rsRetrieveAdd("fld_fax"))
		intCountry = Trim(rsRetrieveAdd("fld_country"))
		strEmail =  Trim(rsRetrieveAdd("fld_email"))
		strCallToll = Trim(rsRetrieveAdd("fld_calltoll"))

	End If
	
'	strCountryQuery = "select Country from sys_countries where Country_id ="&intCountry
'	set rscountry = conn.execute(strCountryQuery)
'	If Not rscountry.Eof Then
'		strCountry = Trim(rscountry("Country"))
'	End If
	'Storing the name into the table:TBL_ResellerMaster
	If Trim(Request("flag")) = "1" Then			
		
		strCompName = fixQuotes(Trim(Request("txtcompany")))
		strAdd = fixQuotes(Trim(Request("txtAddress")))
		strCity = fixQuotes(Trim(Request("txtCity")))
		strState = fixQuotes(Trim(Request("txtState")))
		intZip = Trim(Request("txtZipCode"))
		intPhone = Trim(Request("txtPhone"))
		intAltNum = Trim(Request("txtAltNum"))
		strfax = Trim(Request("txtFax"))
		intCountry = Trim(Request("selcountry"))
		strEmail =  Trim(Request("txtEmail"))
		strCallToll = Trim(Request("txtCalltoll"))
          strSql ="sp_savereseladd " & intResellerId & ",' "& strCompName &" ',"&_
														" '"& strAdd &" ',' " & strCity & " ',' "& strState &" ', '"&intZip &_
														"', '" & intPhone & "', '"& intAltNum &"' ,'"& strfax &"', '"& intCountry &_
														"' ,'" & strEmail &"','" & strCallToll & "'"
		conn.execute strSql

			
		response.redirect "Reseller_contact_us.asp"
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
							<TD class=pagetitle height=400 vAlign=top>Add Contact Address
								<TABLE align="center" border="0" cellPadding="2" cellSpacing="0" width="100%">			
										<TBODY>
											<br>
											<tr>		
												<td>Company Name</td>		
												<td height=1 width="98%">
													<INPUT maxLength=200 name="txtcompany" value="<%=strCompName%>">
												</td>
											</tr>
											<tr>		
												<td>Address</td>		
												<td height=1 width="98%">
													<INPUT maxLength=100 name="txtAddress" value="<%=strAdd%>">
												</td>
											</tr>
											<tr>		
												<td>City</td>		
												<td height=1 width="98%">
													<INPUT maxLength=100 name="txtCity" value="<%=strCity%>">
												</td>
											</tr>
											<tr>		
												<td>State</td>		
												<td height=1 width="98%">
													<INPUT maxLength=100 name="txtState" value="<%=strState%>">
												</td>
											</tr>
											<tr>		
												<td>ZipCode</td>		
												<td height=1 width="98%">
													<INPUT maxLength=20 name="txtZipCode" value="<%=intZip%>">
												</td>
											</tr>
											<tr>		
												<td>Phone</td>		
												<td height=1 width="98%">
													<INPUT maxLength=20 name="txtPhone" value="<%=intPhone%>">
												</td>
											</tr>
											<tr>		
												<td>Alternate number</td>		
												<td height=1 width="98%">
													<INPUT maxLength=20 name="txtAltNum" value="<%=intAltNum%>">
												</td>
											</tr>
											<tr>		
												<td>Fax</td>		
												<td height=1 width="98%">
													<INPUT maxLength=20 name="txtFax" value="<%=strfax%>">
												</td>
											</tr>
											<tr>		
												<td>Country</td>
												<td>
												<%
												'To retrive country name & country id from the database
												strretrievecountry="select Country,Country_id from sys_countries where country <> 'All Countries' ORDER BY Country"
												Set rscountry = conn.execute(strretrievecountry)
												If Not rscountry.Eof Then%>
													<select name="selcountry">
														<option  value="0" <%=sel%>>--Select Country--</option>
														<% While Not rscountry.Eof		
															countryname = Trim(rscountry("Country"))			
															countryid = Trim(rscountry("Country_id"))			
															If intcountry = Trim(rscountry("Country_id")) Then 
																 sel = "selected"
															Else
																sel = ""
															End If			
															%>
															<option  value="<%= countryid%>" <%=sel%>><%= countryname%></option>			
															<%			
															rscountry.movenext
															Wend
														End If 
														rscountry.close
														set rscountry=nothing
														%>
													</select>
												</td>
											</tr>	
											<tr>		
												<td>Email</td>		
												<td height=1 width="98%">
													<INPUT maxLength="50" name="txtEmail" value="<%=strEmail%>">
												</td>
											</tr>
											<tr>		
												<td>Call toll free</td>		
												<td height=1 width="98%">
													<INPUT maxLength="50" name="txtCalltoll" value="<%=strCallToll%>">
												</td>
											</tr>
											<tr>
												<td colspan="2"><br></td>												
											</tr>
											<tr>
												<td>
													 <INPUT name="save" type="button" value="Save" onclick="javascript:fnSave()"> 
												</td>
												<td><input type="reset" value="Reset"></td>											
											</tr>		            
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
