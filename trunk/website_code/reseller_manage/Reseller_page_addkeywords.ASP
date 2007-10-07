 <!--#include file = "include/header.asp"-->
<%
Response.Buffer = true
Response.Expires = -1441


'	#################	HEADER INFORMATION	############
'	-----------------------------------------------------------------------------------------
'	Purpose of page:	This page show the description for the selected page
'	Page Name:		   	reseller_page_addkeyword.asp
'	Version Information:EasystoreCreator
'	Input Page:		    Reseller_change_HKeywords.asp
'	Output Page:	    Reseller_change_HKeywords.asp
'	Date & Time:		12 th August
'	Created By:			Devki Anote
'	-----------------------------------------------------------------------------------------
'	#################	HEADER INFORMATION	############## 
mSel1=3
mSel2=2
%>

<HTML><HEAD><TITLE>Reseller Admin</TITLE>
<META content=no-cache name=Pragma>
<META content=no-cache http-equiv=pragma>
<META content="text/html; charset=windows-1252" http-equiv=Content-Type><LINK 
href="images/style.css" rel=stylesheet type=text/css>
<SCRIPT language=JavaScript src="images/script.js" 
type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript src="../include/commonfunctions.js"></SCRIPT>
<script language="javascript">

//function is called when the user clicks on save button
function fnSave()
{

	var ErrMsg = "" 
	
	if (isWhitespace(document.frm.txtkeyword1.value) == true)
	{
		ErrMsg = ErrMsg + "Keyword 1 is mandatory.\n"
	}
	
	if (isWhitespace(document.frm.txtkeyword2.value) ==  true)
	{
		ErrMsg = ErrMsg + "Keyword 2 is mandatory."
	}
	
	
	if (ErrMsg!="")
	{
		alert(ErrMsg)
	}
	
	else
	{
	
		document.frm.action = "reseller_page_addkeywords.asp?saveaction=save"
		document.frm.submit()
	}
	
	
}
</script>
<META content="MSHTML 5.00.3813.800" name=GENERATOR></HEAD>
<BODY bottomMargin=0 leftMargin=0 rightMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
'code here to get the page id 
if trim(Request.QueryString("action"))<> "" then 
	intpageid = trim(Request.QueryString("action"))
	
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
	'code here to check if the keywords are there or not
	Set rsKey = Server.CreateObject("ADODB.Recordset")
	With rsKey
		.Source = "Get_Reseller_Page_Key " & session("ResellerID")&", "&intpageid
		.ActiveConnection = strConn
		.CursorType = adOpenForwardOnly
		.LockType = adLockReadOnly
		.Open
	End With
	
	
	if not rsKey.EOF then 
		Keyword1 = trim(checkencode(rsKey("fld_Reseller_Page_Keyword1")))
		Keyword2 = trim(checkencode(rsKey("fld_Reseller_Page_Keyword2")))
		Keyword3 = trim(checkencode(rsKey("fld_Reseller_Page_Keyword3")))
		Keyword4 = trim(checkencode(rsKey("fld_Reseller_Page_Keyword4")))
		Keyword5 = trim(checkencode(rsKey("fld_Reseller_Page_Keyword5")))
	end if
	Set rsKey = Nothing
	
end if

'code here to perform the insert/update operation after the user clicks save button

if trim(Request.QueryString("saveaction")) = "save" then
	
	intpageid = trim(Request.Form("hidpageid"))
	strKeyword1 = trim(Request.Form("txtKeyword1"))
	strKeyword2 = trim(Request.Form("txtKeyword2"))
	strKeyword3 = trim(Request.Form("txtKeyword3"))
	strKeyword4 = trim(Request.Form("txtKeyword4"))
	strKeyword5 = trim(Request.Form("txtKeyword5"))
	
	strKeyword1 = fixquotes(strKeyword1)
	strKeyword2 = fixquotes(strKeyword2)
	strKeyword3 = fixquotes(strKeyword3)
	strKeyword4 = fixquotes(strKeyword4)
	strKeyword5 = fixquotes(strKeyword5)
	
	'code here to check if the keywords are there or not
	Set rsKey = Server.CreateObject("ADODB.Recordset")
	With rsKey
		.Source = "Get_Reseller_Page_Key " & session("ResellerID")&", "&intpageid
		.ActiveConnection = strConn
		.CursorType = adOpenForwardOnly
		.LockType = adLockReadOnly
		.Open
	End With
	
	if rsKey.eof then 
		'means keywords are not there and you have to insert
		 	Set rsPutKey = Server.CreateObject("ADODB.Recordset")
			With rsPutKey
				.Source = "Put_Reseller_Page_Key " & session("ResellerID")&", "&intpageid&",'"&strKeyword1&"','"&strKeyword2&"'"&_
						  ",'"&strKeyword3&"','"&strKeyword4&"','"&strKeyword5&"'"	
				.ActiveConnection = strConn
				.CursorType = adOpenForwardOnly
				.LockType = adLockReadOnly
				.Open
			End With
			Set rsPutKey = Nothing
			intflag ="2"	
	
	
		 
	else
		
		'means keywords are there you have to update
		 Set rsUpdateKey = Server.CreateObject("ADODB.Recordset")
			With rsUpdateKey
				.Source = "update_Reseller_Page_Key " & session("ResellerID")&", "&intpageid&",'"&strKeyword1&"','"&strKeyword2&"'"&_
						  ",'"&strKeyword3&"','"&strKeyword4&"','"&strKeyword5&"'"	
				.ActiveConnection = strConn
				.CursorType = adOpenForwardOnly
				.LockType = adLockReadOnly
				.Open
			End With
			Set rsUpdateKey = Nothing
			intflag ="1"
	end if 
	Set rsKey = Nothing
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
                href="Reseller_change_Desc.asp">Define Description               </A></TD>
        </TR>    
                          
                  </TBODY></TABLE>
        </TD>
         <TD height=15 vAlign=top width=570></TD></TR>
        <TR>
          <TD class=pagetitle height=400 vAlign=top>Add Keywords
        <FORM action="" method=post name="frm">
        
            <TABLE align=left border=0 cellPadding=0 cellSpacing=0 
              width="90%">
              
         <TR>
          <font color="red" size="1">Note : You can add a maximum of 5 Keywords.</font>
          </TR>
	          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><FONT face=Arial size=2><b><%=pagename%></b></FONT></TD>
          <TD width="1%"></TD>
          </TR>
	          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><br><FONT face=Arial size=2>Add Keyword 1</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=100 name=txtkeyword1 value="<%=keyword1%>"> 
            </TD></TR>
			          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><br><FONT face=Arial size=2>Add Keyword 2</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=100 name=txtkeyword2 value="<%=keyword2%>" > 
            </TD></TR>
			          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><br><FONT face=Arial size=2>Add Keyword 3</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=100 name=txtkeyword3 value="<%=keyword3%>" > 
            </TD></TR>
	          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><br><FONT face=Arial size=2>Add Keyword 4</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=100 name=txtkeyword4 value="<%=keyword4%>" > 
            </TD></TR>
	          
         <TR>
          <TD noWrap width="1%">&nbsp;</TD>
          <TD noWrap width="1%"><br><FONT face=Arial size=2>Add Keyword 5</FONT></TD>
          <TD width="1%"></TD>
          <TD height=1 width="98%"><INPUT maxLength=100 name=txtkeyword5 value="<%=keyword5%>" > 
            </TD></TR>
	        
        <tr>		
        
        <TD noWrap width="1%">&nbsp;</TD>
        
        <td noWrap width="1%">
        
        <input type=hidden value="<%=intpageid%>" name="hidpageid" >
        <br>
        <INPUT name="save" type="button" value="Save" onclick="javascript:fnSave()">&nbsp; 
        <input type="reset" value="Reset" name="Reset">        
				</td></tr>	
		  
              
              <TBODY>
         

                <TD height=20></TD></TR>
              </TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TABLE></FORM>
</BODY></HTML>
<%if intflag = "1" then %>
	<script language="javascript">
		alert("Keywords for the page are updated.")
		document.location.href = "Reseller_change_HKeywords.asp"
	</script>
<% end if%>
<%if intflag = "2" then %>
	<script language="javascript">
		alert("Keywords for the page are added.")
		document.location.href = "Reseller_change_HKeywords.asp"
	</script>
<% end if%>