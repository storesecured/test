<!--#include file="include/header.asp"-->
<!--#include file="include/scart_misc.asp"-->


<table><tr><td align='middle'>

<form action='search_items.asp' method='post' name=search>
<input type=hidden name=Form_Name value=Search_Store>
<table border='0' cellPadding='3' cellSpacing='2'><tr>
<td align='left' class='normal'>Search Keywords</td>
</tr><tr>
<td class='normal'><input name='Search_Text' size='40'></td>
</tr><tr><td align='left'>
<table cellspacing='0' cellpadding='0' border='0' width='100%'><tr>
<td colspan='4' class='normal'><I><b>Search Options</b></i></td>
</tr><tr>
<td nowrap>&nbsp;</td>
<td nowrap class='normal'>In Department</td>
<td nowrap class='normal'>&nbsp;</td>
<td nowrap class='normal'>
<select name='sub_department_id'>
<option selected value='0'>All</option>
<%
  sql_select = "select Department_ID,Full_name from Store_Dept WITH (NOLOCK) where Store_id="&Store_id&" and Department_ID <> 0 and visible=1 order by full_name"
  
  set myfields=server.createobject("scripting.dictionary")
  Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)

  if noRecords = 0 then
  FOR rowcounter= 0 TO myfields("rowcount")
        sId=mydata(myfields("department_id"),rowcounter)
        sName=mydata(myfields("full_name"),rowcounter)
        if len(sName)>70 then
            sNameShort=" . . . "&right(sName,67)
        else
            sNameShort=sName    
        end if
	 response.write "<option value='"&sId&"'>"&sNameShort&"</option>"
  Next
  End If
%>
</select></td></tr><tr>
<td nowrap class='normal'>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td nowrap class='normal'>With Price Between</td>
<td nowrap class='normal'>&nbsp;</td>
<td nowrap class='normal'><input type='text' size='8' name='SEARCH_MINPRICE'>&nbsp;And&nbsp;<input type='text' size='8' name='SEARCH_MAXPRICE'></td>
</tr><tr>
<td nowrap class='normal'>&nbsp;</td>
<td nowrap class='normal'>with SKU like</td>
<td nowrap class='normal'>&nbsp;</td>
<td nowrap class='normal'><input type='text' size='20' name='SEARCH_SKU'></td>
</tr></table></td></tr><tr>
<td align='left' colSpan='2' class='normal'>
<%= fn_create_action_button ("Button_image_Search","Search_Store","Search") %>
</td></tr></table></form>
<%
      sFormName = "search"
      sSubmitName = "Search_Store"

      %>
      <SCRIPT language="JavaScript">
       var frmvalidator  = new Validator("<%= sFormName %>");

      </script>



</td></tr></table>

<!--#include file="include/footer.asp"-->
