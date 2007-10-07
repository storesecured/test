<!--#include file="Global_Settings.asp"-->
<head>
<SCRIPT LANGUAGE="JavaScript">

var arrUserInfo1D = new Array();

function move1(fbox, tbox) {
var arrFbox = new Array();
var arrTbox = new Array();
var arrLookup = new Array();
var i;
for (i = 0; i < tbox.options.length; i++) {
arrLookup[tbox.options[i].text] = tbox.options[i].value;
arrTbox[i] = tbox.options[i].text;
}
var fLength = 0;
var tLength = arrTbox.length;
for(i = 0; i < fbox.options.length; i++) {
arrLookup[fbox.options[i].text] = fbox.options[i].value;
if (fbox.options[i].selected && fbox.options[i].value != "") {
arrTbox[tLength] = fbox.options[i].text;
tLength++;
}
else {
arrFbox[fLength] = fbox.options[i].text;
fLength++;
   }
}
arrFbox.sort();
arrTbox.sort();
fbox.length = 0;
tbox.length = 0;
var c;
for(c = 0; c < arrFbox.length; c++) {
var no = new Option();
no.value = arrLookup[arrFbox[c]];
no.text = arrFbox[c];
fbox[c] = no;
}
for(c = 0; c < arrTbox.length; c++) {
var no = new Option();
no.value = arrLookup[arrTbox[c]];
no.text = arrTbox[c];
tbox[c] = no;
 }

var list2;
list2="";

window.parent.opener.globlist = ""
for(i = 0; i < fbox.options.length; i++) {
window.parent.opener.globlist = window.parent.opener.globlist  + fbox.options[i].text + ",";
}
}


function move2(fbox, tbox) {
var arrFbox = new Array();
var arrTbox = new Array();
var arrLookup = new Array();
var i;
for (i = 0; i < tbox.options.length; i++) {
arrLookup[tbox.options[i].text] = tbox.options[i].value;
arrTbox[i] = tbox.options[i].text;
}
var fLength = 0;
var tLength = arrTbox.length;
for(i = 0; i < fbox.options.length; i++) {
arrLookup[fbox.options[i].text] = fbox.options[i].value;
if (fbox.options[i].selected && fbox.options[i].value != "") {
arrTbox[tLength] = fbox.options[i].text;
tLength++;
}
else {
arrFbox[fLength] = fbox.options[i].text;
fLength++;
   }
}
arrFbox.sort();
arrTbox.sort();
fbox.length = 0;
tbox.length = 0;
var c;
for(c = 0; c < arrFbox.length; c++) {
var no = new Option();
no.value = arrLookup[arrFbox[c]];
no.text = arrFbox[c];
fbox[c] = no;
}
for(c = 0; c < arrTbox.length; c++) {
var no = new Option();
no.value = arrLookup[arrTbox[c]];
no.text = arrTbox[c];
tbox[c] = no;
 }

window.parent.opener.globlist = ""
for(i = 0; i < tbox.options.length; i++) {
window.parent.opener.globlist = window.parent.opener.globlist  + tbox.options[i].text + ",";
}
}


function fillList() {  


var arg;
arg = "<%= request("returnArg") %>"


if (arg == "Discounted_Items_Skus") {
if (window.parent.opener.itemflag==1) {
	window.parent.opener.globlist=window.parent.opener.DiscVal;
	window.parent.opener.itemflag=0;
}
}

if (arg == "Free_items") {
if (window.parent.opener.itemflag==1) {
	window.parent.opener.globlist=window.parent.opener.FreeVal;
	window.parent.opener.itemflag=0;
}
}

var flgCheck;
var lstCnt = 0;
str = window.parent.opener.globlist.split(",")

for(c = 0; c < str.length-1; c++) {
var no = new Option();
no.value = str[c];
no.text = str[c];
document.items_select.list2[c]  = no;
}


if (window.parent.opener.itemflag != 1) {
	for (c=0;c<arrUserInfo1D.length;c++) {
		flgCheck= 1;
			for(n = 0; n < str.length-1; n++) {
					if(arrUserInfo1D[c] == str[n]) {
					 // alert(arrUserInfo1D[c]);
					 flgCheck = 0;
					 break;
					}
			}

			if (flgCheck == 1) {
				var no = new Option();
				no.value = arrUserInfo1D[c];
				no.text = arrUserInfo1D[c];
				document.items_select.list1[lstCnt]  = no;
				lstCnt ++;
			 }
	}
}
}

</script>

<%
'This function converts a VBScript Array to Javascript Array
Function ConvertToJSArray1D(VBArray , ArrayName)
	Dim vb2jsRow , vb2jsStr, vb2jsi
	vb2jsRow = Ubound(VBArray,1)
	%>
	<SCRIPT LANGUAGE = 'JAVASCRIPT' >
	var vb2jsi
	<%=ArrayName%> = new Array(<%=vb2jsRow+1%>);
	for (vb2jsi=0; vb2jsi < <%=vb2jsRow+1%>; vb2jsi++) 
	{ 
		<%=ArrayName%>[vb2jsi]= " "
	}
	</SCRIPT>
	<%
	Response.Write("<SCR"&"IPT LANGUAGE = 'JAVASCRIPT' >"&chr(13))
	for vb2jsi=0 to vb2jsRow
		vb2jsstr = "VBArray("&vb2jsi&")"
	%>     
		<%=ArrayName%>[<%=vb2jsi%>]= "<%=trim(eval(vb2jsstr))%>" 
	<%
	Next
	Response.Write("</SCR"&"IPT>")
End Function
%>


</head>



<% 
dim clevel 
dim fclicks

dim theSubDepts

sub getSubDepts(masterD)
	if (theSubDepts="") then
		theSubDepts = cstr(masterD)
	else
		theSubDepts = theSubDepts&", "&cstr(masterD)
	end if
	sql_sd = "select department_id from store_dept where Belong_to="&masterD&" and store_id="&store_id
	set myfieldslocal=server.createobject("scripting.dictionary")
   Call DataGetrows(conn_store,sql_sd,mydatalocal,myfieldslocal,noRecordslocal)

	if noRecordslocal then
		FOR rowcounterlocal= 0 TO myfieldslocal("rowcount")
			getSubDepts mydatalocal(myfieldslocal("department_id"),rowcounterlocal)
		Next
	end if
end sub

sub create_instance_dept_tree(link_color,link_face,link_size)
	
	small_link_size = link_size - 1
	strsql = "select D1.department_name, D1.department_id, (select count(d2.department_id) from store_dept d2 where d2.belong_to=d1.department_id and d2.store_id="&store_id&") as sdepts from store_Dept D1 where (D1.belong_to=-1) and D1.store_id="&store_id

	set myfields=server.createobject("scripting.dictionary")
    Call DataGetrows(conn_store,strsql,mydata,myfields,noRecords)


	if noRecords = 0 then %>

	<script src="include/ftiens4.js">
	</script>
	<script>
	USETEXTLINKS = 1
	
	sethrefTarget(' target=\"_self\" ');
	
	level_0 = gFld("<font face=<%= link_face %> size=<%= small_link_size %>><%= Replace(store_name,"'","''") %></font>", "items_select?dept_id=-1&returnArg=<%= request.queryString("returnArg") %>")<%

	clevel = 1
	fclicks = ";"&vbcrlf
	 FOR rowcounter= 0 TO myfields("rowcount")
		if mydata(myfields("sdepts"),rowcounter)>0 then%>
			level_<%= clevel %> = insFld(level_0, gFld("<font face=<%= link_face %> size=<%= small_link_size %>><%= Replace(mydata(myfields("department_name"),rowcounter),"'","''") %></font>", "items_select.asp?dept_id=<%= mydata(myfields("department_id"),rowcounter) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
			fclicks = fclicks&"clickOnFolder("&clevel&");"&vbcrlf
			clevel = clevel+1
			create_instance_sub_dept_tree link_color,link_face,link_size,"level_"&clevel-1 ,mydata(myfields("department_id"),rowcounter),"0,"&mydata(myfields("department_id"),rowcounter)
		else
			clevel = clevel+1%>
			insDoc(level_0, gLnk(0, "<font face=<%= link_face %> size=<%= small_link_size %>><%= mydata(myfields("department_name"),rowcounter) %></font>", "items_select.asp?dept_id=<%= mydata(myfields("department_id"),rowcounter) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
		end if

	Next
	End If %>
	</script>
	<font face="<%= link_face %>" size="<%= small_link_size %>">
	<script>
		initializeDocument()
		<%= fclicks %>
	</script>
	</font><%

end sub

sub create_instance_sub_dept_tree(link_color, link_face, link_size, parent_dept_level, parent_dept, parents_ids)
	small_link_size = link_size - 1
	strsql = "select D1.department_name, D1.department_id, (select count(d2.department_id) from store_dept d2 where d2.belong_to=d1.department_id and d2.store_id="&store_id&") as sdepts from store_Dept D1 where D1.belong_to="&parent_dept&" and D1.store_id="&store_id 

	set myfieldslocal=server.createobject("scripting.dictionary")
   Call DataGetrows(conn_store,strsql,mydatalocal,myfieldslocal,noRecordslocal)


	if noRecordslocal = 0 then

	 FOR rowcounterlocal= 0 TO myfieldslocal("rowcount")

		if mydatalocal(myfieldslocal("sdepts"),rowcounterlocal)>0 then%>
			level_<%= clevel %> = insFld(<%= parent_dept_level %>, gFld("<font  face=<%= link_face %> size=<%= small_link_size %>><%= Replace(mydatalocal(myfieldslocal("department_name"),rowcounterlocal),"'","''") %></font>", "items_select.asp?dept_id=-<%=mydatalocal(myfieldslocal("department_id"),rowcounterlocal) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
			fclicks = fclicks&"clickOnFolder("&clevel&");"&vbcrlf
			clevel = clevel+1
			create_instance_sub_dept_tree link_color,link_face,link_size,"level_"&clevel-1 ,mydatalocal(myfieldslocal("department_id"),rowcounterlocal),parents_ids&","&mydatalocal(myfieldslocal("department_id"),rowcounterlocal)
		else
			clevel = clevel+1%>	
			insDoc(<%= parent_dept_level %>, gLnk(0, "<font face=<%= link_face %> size=<%= small_link_size %>><%= mydatalocal(myfieldslocal("department_name"),rowcounterlocal) %></font>", "items_select.asp?dept_id=<%= mydatalocal(myfieldslocal("department_id"),rowcounterlocal) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
		end if
	Next
	end if
end sub

%>
 
<html>


<SCRIPT LANGUAGE = "JavaScript">
	<!--
	function setParentResults(resArg,resVal) {
		if (resArg=="Discounted_Items_Skus") {
			window.parent.opener.DiscVal = resVal;
		}
		if(window.parent.opener.DiscVal != "") {
			window.parent.opener.DiscVal = window.parent.opener.DiscVal + ",";
		}
		if (resArg=="Free_items") {
			window.parent.opener.FreeVal = resVal;
		}
		if(window.parent.opener.FreeVal != "") {
			window.parent.opener.FreeVal = window.parent.opener.FreeVal + ",";
		}
		window.parent.opener.setResults(resArg, resVal);
		window.parent.close();}
	//-->

</SCRIPT>

<body onLoad="javascript:fillList();">
<form name="items_select">
<input type="hidden" name="hidVal">
<table border="3" width="100%" height="100%" cellpadding="0" cellspacing="0" bordercolor="000066">
	<tr>
		<td colspan="3">
			<table border="1" width="100%" height="100%" cellpadding="3" cellspacing="3">			
				<tr>
					<td align="center"><font face="Arial" size="2"><b>Store Departments</b></font></td>
						<%
						if request.querystring("dept_id")<>"" then   
                            dept_id=request.querystring("dept_id")
							if cint(request.querystring("dept_id"))<0 then
							    dept_id=0-dept_id
							end if
							sql_dept = "select Department_Name from store_dept where department_id = "&dept_id&" and store_id="&store_id
							rs_store.open sql_dept, conn_store, 1, 1
							dep_name = rs_store("Department_Name")
							rs_store.close 
							response.Write "<td align='center'><font face='Arial' size='2'><b>Display "&dep_name&"</b></font></td>"
							sql_sel = "select show, item_id, item_name, item_sku from store_items where item_id in (select item_id from store_items_dept where store_id="&store_id&" and department_id="&dept_id&") and store_id="&store_id&" order by item_sku"				
                        Else
                            response.write "<td align='center'><font face='Arial' size='2'><b>Display No Dept Selected</b></font></td>"
                        End If
                        %>
				</tr>
		
				<tr>
					<td width="30%" valign="center">
						<div style=" width:100%; height:100%; overflow:auto;">
								<% call create_instance_dept_tree("","Arial","3") %>
						</div>
					</td>
					<td width="70%">
					<table border="0" width="100%" height="100%" cellspacing="0">
					<tr valign="top">
						<td width="42%" align="center"><b>Select items to move</b><br>
							<select multiple size="20" name="list1" style="width:150">
								<% if request.querystring("dept_id")<>"" then %>
								<% set myfields=server.createobject("scripting.dictionary")
								  Call DataGetrows(conn_store,sql_sel,mydata,myfields,noRecords) %>						

					<% if noRecords = 0 then


				Dim  arrUserInfo1D, iRcnt
				iRcnt = cint(myfields("rowcount"))
				Redim arrUserInfo1D(iRcnt)


							FOR rowcounter= 0 TO iRcnt 								
								arrUserInfo1D(rowcounter) = mydata(myfields("item_sku"),rowcounter)
							 Next %>
					<% Call ConvertToJSArray1D(arrUserInfo1D,"arrUserInfo1D") %>
					<% End If %>
				<% End If %>
						</select>
					</td>
					<td align="center" width="16%">
		  				<input type="button" onClick="move1(this.form.list2,this.form.list1)" value="<<">
						<input type="button" onClick="move2(this.form.list1,this.form.list2)" value=">>">
					</td>
		  
					<td width="42%" align="center"><b>Selected items</b><br>
						<select multiple size="20" name="list2" style="width:150">
					    </select>
					</td>
				</tr>
				<tr><td colspan="3">&nbsp;</td></tr>
				<tr>
					<td colspan=3 align=center>
						<input type="button" name="select" value="select" onClick="JavaScript:setParentResults('<%= request.queryString("returnArg") %>',window.parent.opener.globlist.substring(0,window.parent.opener.globlist.length-1));" >
					</td>
				</tr>
				<tr><td colspan="3">&nbsp;</td></tr>
		</table>
			</td>
				</tr>	
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="hidList">
</form>
</body> 

</html>
