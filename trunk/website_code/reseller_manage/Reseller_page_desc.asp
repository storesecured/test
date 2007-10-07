 <!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page show the description for the selected page
'	Page Name:		   	reseller_page_Desc.asp
'	Version Information:EasystoreCreator
'	Input Page:		    Reseller_change_Desc.asp
'	Output Page:	    Reseller_change_Desc.asp
'	Date & Time:		11 th August
'	Created By:			Devki Anote
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel1=3
mSel2=2
%>

<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/.aspl; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="../include/commonfunctions.js"></SCRIPT>
<script language="javascript">

function fnSave()
{
	var ErrMsg = "" 
	var str,str1
	
	if (isWhitespace(document.frm.txtDesc.value)== true)
	{
		ErrMsg = ErrMsg + "Description is mandatory.\n" 
	}
	if (isWhitespace(document.frm.txtDesc.value)== false)
	{
			
			str=document.frm.txtDesc.value.length
			//str1=str-1
			//alert(str)
			if (str > 200) 
			{
			
			ErrMsg = ErrMsg + "Description should not be more than 200 characters.\n" 
			}
	}
	
	if (ErrMsg!="")
	{
		alert(ErrMsg)
	}
	else
	{
		document.frm.action = "Reseller_Page_Desc.asp?saveaction=save"
		document.frm.submit()
	}
}

</script>

<META content="MS.aspL 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightmargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
'code here to get the page id 
if trim(Request.QueryString("action"))<> "" then 
	intpageid = trim(Request.QueryString("action"))
	
	'code here to retreive the name of the page here
		'code here to retreive the name of the page here
	Set rsPageName = Server.CreateObject("ADODB.Recordset")
	With rsPageName
		.Source = "Get_Page_name " &intpageid
		.ActiveConnection = strConn
		.CursorType = adOpenForwardOnly
		.LockType = adLockReadOnly
		.Open
	End With
	if not rsPageName.eof then 
		pagename = trim(rsPageName("fld_page_name"))
	end if
	
	Set rsPageName = Nothing	
	'code here to retreive the description here
	Set rsDesc = Server.CreateObject("ADODB.Recordset")
	With rsDesc
		.Source = "Get_Reseller_Page_Desc " & session("ResellerID")&", "&intpageid
		.ActiveConnection = strConn
		.CursorType = adOpenForwardOnly
		.LockType = adLockReadOnly
		.Open
	End With
	
	if not rsDesc.EOF then 
			strDesc = trim(rsDesc("fld_desc_text"))
			
			
	end if
	Set rsDesc = Nothing
end if


'code here to update the description here
if trim(Request.QueryString("saveaction")) = "save" then 
	pageid = trim(Request.Form("HidPageId"))
	pageDesc = trim(checkencode(Request.Form("txtDesc")))
	
	pageDesc = fixquotes(pageDesc)
	
	
	'code here to check whether there is any description for the page
	
	'code here to retreive the description here
	Set rsDesc = Server.CreateObject("ADODB.Recordset")
	With rsDesc
		.Source = "Get_Reseller_Page_Desc1 " & session("ResellerID")&", "&pageid 
		.ActiveConnection = strConn
		.CursorType = adOpenForwardOnly
		.LockType = adLockReadOnly
		.Open
	End With
	
	
	if rsDesc.eof then 
	
	
		'procedure call here to insert the desc for that page if no value is there 
		Set rsDesc1 = Server.CreateObject("ADODB.Recordset")
		With rsDesc1
			.Source = "Put_Reseller_Page_Desc '"&pageDesc&"' ,"& session("ResellerID")&", "&pageid
			.ActiveConnection = strConn
			.CursorType = adOpenForwardOnly
			.LockType = adLockReadOnly
			.Open
		End With
		intFlag = "2"
		Set rsDesc1 = Nothing	
	
	else
	
		'procedure call here to update the desc for that page if some value is there 
		Set rsDesc2 = Server.CreateObject("ADODB.Recordset")
		With rsDesc2
			.Source = "Update_Reseller_Page_Desc '"&pageDesc&" ' ,"&session("ResellerID")&", "&pageid
			.ActiveConnection = strConn
			.CursorType = adOpenForwardOnly
			.LockType = adLockReadOnly
			.Open
		End With
		intFlag = "1"
		Set rsDesc2 = Nothing	
	end if
		Set rsDesc = Nothing		
end if
%>
<DIV id=overDiv 
style="POSITION: absolute; VISIBILITY: hidden; Z-INDEX: 1000"></DIV>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=750>
  <TBODY>
  <TR>
    <TD class=title>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        <TR>
          <TD class=title><B>Reseller</B></TD>
          <TD class=special width=200>&nbsp;</TD>
          <TD align=left class=special width="60%">
            <UL><BR><BR></UL></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD>
     
        <!--#include file="incmenu.asp"-->
     </TD></TR>
  <TR>
    <TD>
      <TABLE border=0 cellPadding=0 cellSpacing=0>
        <TBODY>
        </TBODY></TABLE></TD></TR>
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
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
               href="Reseller_change_Logo.asp">Change Logo
              </A></TD>
        </TR>
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
              href="Reseller_change_HKeywords.asp">Manage Keywords
              </A></TD>
        </TR>      
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
                href="Reseller_change_Plan.asp">Change Plan Pricing
              </A></TD>
        </TR>                         
        <TR>
            <TD class=meniu height=20 
            onmouseout="style.backgroundColor='#EBF9D8', style.color='#8FB25E'" 
            onmouseover="style.backgroundColor='#8FB35B', style.color='#ffffff'"><A 
              class=b 
                href="Reseller_change_Desc.asp">Define Description</A></TD>
        </TR>      
        
        
                          
                  </TBODY></TABLE></TD>
          <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Define Description
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              <FORM action="" method=post name="frm">
              <TR>
				<TD>	
				<b>	<%=pagename%> </b>
				</TD>
			 </TR>
              
              <tr>
               <br>
				 <td>
					 Add Description
				</td>
				<td>
					 <textarea name="txtDesc" rows=7 cols=45 ><%=strDesc%></textarea>
					 
				</td>	
				</tr>	
				<input type="hidden" value="<%=intpageid%>" name="hidpageid">
              <tr>
                <td>
					  
				</td>
				<td>	
				<br>
					 <INPUT name="save" type="button" value="Save" onclick="javascript:fnSave()"> 
					 <input type="reset" type="reset" value="Reset">
					
				</td>	
              
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></html>
<%if intFlag="1" then %>
	
	<script language="javascript">
		alert("Description for the page is updated.")
		document.location.href = "Reseller_change_Desc.asp"
		
		
	</script>
	
<% end if%>
<%if intFlag="2" then %>
	
	<script language="javascript">
		alert("Description for the page is added.")
		document.location.href = "Reseller_change_Desc.asp"
	</script>
	
<% end if%>