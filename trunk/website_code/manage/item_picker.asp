<!--#include file="Global_Settings.asp"-->

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
	
	level_0 = gFld("<font face=<%= link_face %> size=<%= small_link_size %>><%= Replace(store_name,"'","''") %></font>", "item_picker.asp?dept_id=-1&returnArg=<%= request.queryString("returnArg") %>")<%

	clevel = 1
	fclicks = ";"&vbcrlf
	 FOR rowcounter= 0 TO myfields("rowcount")
		if mydata(myfields("sdepts"),rowcounter)>0 then%>
			level_<%= clevel %> = insFld(level_0, gFld("<font face=<%= link_face %> size=<%= small_link_size %>><%= Replace(mydata(myfields("department_name"),rowcounter),"'","''") %></font>", "item_picker.asp?dept_id=<%= mydata(myfields("department_id"),rowcounter) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
			fclicks = fclicks&"clickOnFolder("&clevel&");"&vbcrlf
			clevel = clevel+1
			create_instance_sub_dept_tree link_color,link_face,link_size,"level_"&clevel-1 ,mydata(myfields("department_id"),rowcounter),"0,"&mydata(myfields("department_id"),rowcounter)
		else
			clevel = clevel+1%>
			insDoc(level_0, gLnk(0, "<font face=<%= link_face %> size=<%= small_link_size %>><%= mydata(myfields("department_name"),rowcounter) %></font>", "item_picker.asp?dept_id=<%= mydata(myfields("department_id"),rowcounter) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
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
			level_<%= clevel %> = insFld(<%= parent_dept_level %>, gFld("<font  face=<%= link_face %> size=<%= small_link_size %>><%= Replace(mydatalocal(myfieldslocal("department_name"),rowcounterlocal),"'","''") %></font>", "item_picker.asp?dept_id=-<%=mydatalocal(myfieldslocal("department_id"),rowcounterlocal) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
			fclicks = fclicks&"clickOnFolder("&clevel&");"&vbcrlf
			clevel = clevel+1
			create_instance_sub_dept_tree link_color,link_face,link_size,"level_"&clevel-1 ,mydatalocal(myfieldslocal("department_id"),rowcounterlocal),parents_ids&","&mydatalocal(myfieldslocal("department_id"),rowcounterlocal)
		else
			clevel = clevel+1%>	
			insDoc(<%= parent_dept_level %>, gLnk(0, "<font face=<%= link_face %> size=<%= small_link_size %>><%= mydatalocal(myfieldslocal("department_name"),rowcounterlocal) %></font>", "item_picker.asp?dept_id=<%= mydatalocal(myfieldslocal("department_id"),rowcounterlocal) %>&returnArg=<%= request.queryString("returnArg") %>"))<%
		end if
	Next
	end if
end sub

%>
 
<html>


<SCRIPT LANGUAGE = "JavaScript">
	<!--
	function setParentResults(resArg,resVal) {
		window.parent.opener.setResults(resArg, resVal);
		window.parent.close();}
	//-->

</SCRIPT>

<body>

<table border="1" width="100%" height="100%" cellpadding="0" cellspacing="0" bordercolor="000066">

	
	<tr>
		<td colspan="3">
			<table border="1" width="100%" height="100%" cellpadding="3" cellspacing="3">
				
				<tr>
					<td align="center"><font face="Arial" size="2"><b>Store Departments</b></font></td>
						<% if request.querystring("dept_id")<>"" then

							if cint(request.querystring("dept_id"))=>0 then
								sql_dept = "select Department_Name from store_dept where department_id = "&request.querystring("dept_id")&" and store_id="&store_id
								rs_store.open sql_dept, conn_store, 1, 1
									if not rs_store.eof and not rs_store.bof then
									   dep_name = rs_store("Department_Name")
									end if
								rs_store.close %>
								<td align="center"><font face="Arial" size="2"><b>Display <%= dep_name %></b></font></td>
								<% sql_sel = "select show, i.item_id, item_name, item_sku from store_items_dept i_d WITH (NOLOCK) inner join store_items i WITH (NOLOCK) on i_d.store_id=i.store_id and i_d.item_id=i.item_id where i_d.department_id = ("&request.querystring("dept_id")&") and i_d.store_id="&store_id&" order by item_name"
							End If
							if cint(request.querystring("dept_id"))<0 then
								sql_dept = "select Department_Name from store_dept where department_id = "&(1-cint(request.querystring("dept_id")))&" and store_id="&store_id
								rs_store.open sql_dept, conn_store, 1, 1
                           if not rs_store.eof and not rs_store.bof then
									   dep_name = rs_store("Department_Name")
									end if
								rs_store.close
								theSubDepts = ""
								getSubDepts 0-cint(request.querystring("dept_id")) %>
								<td align="center"><font face="Arial" size="2"><b>Display <%= dep_name %> and SubDepartments</b></font></td>
								<% sql_sel = "select show, i.item_id, item_name, item_sku from store_items_dept i_d WITH (NOLOCK) inner join store_items i WITH (NOLOCK) on i_d.store_id=i.store_id and i_d.item_id=i.item_id where i_d.department_id in ("&theSubDepts&") and i_d.store_id="&store_id&" order by item_name"
							End If
						Else %>
							<td align="center"><font face="Arial" size="2"><b>Display No Dept Selected</b></font></td>
						<% End If %>
				</tr>
		
				<tr>
					<td width="30%" valign="center">
						<% call create_instance_dept_tree("","Arial","3") %>
					</td>
					<td width="70%">
						<table border="1" cellspacing="0" cellpadding="0" width="100%" height="100%">
							
							<tr height="1%">
								<td width="10%" align="center"><font face="Arial" size="2"><b>Item ID</b></font></td>
								<td width="30%" align="center"><font face="Arial" size="2"><b>Item SKU</b></font></td>
								<td width="40%" align="center"><font face="Arial" size="2"><b>Item Name</b></font></td>
								<td width="10%" align="center"><font face="Arial" size="2"><b>Live</b></font></td>
								<td width="10%" align="center"><font face="Arial" size="2">&nbsp;</font></td>
							</tr>
				
							<% if request.querystring("dept_id")<>"" then %>
								<% set myfields=server.createobject("scripting.dictionary")
                          Call DataGetrows(conn_store,sql_sel,mydata,myfields,noRecords)
                           %>
								<% cline = 1 %>
								<% if noRecords = 0 then
                            FOR rowcounter= 0 TO myfields("rowcount")
                             %>
									<tr height="25"
									<% if cline mod 2 = 0 then %>
										bgcolor="#AAAAAA"
									<% End If %>
									>
										<td width="10%" align="center"><font face="Arial" size="2"><%= mydata(myfields("item_id"),rowcounter) %></font></td>
										<td width="30%" align="center"><font face="Arial" size="2"><%= mydata(myfields("item_sku"),rowcounter) %></font></td>
										<td width="40%"><font face="Arial" size="2"><%= mydata(myfields("item_name"),rowcounter) %></font></td>
										<td width="10%" align="center"><font face="Arial" size="2">
										<% if cint(mydata(myfields("show"),rowcounter))=-1 then%>
											Yes
										<% Else %>
											No
										<% End If %>
										</font></td>
										<td width="10%" align="center"><font face="Arial" size="2">

											<a class=link href="JavaScript:setParentResults('<%= request.queryString("returnArg") %>','<%= mydata(myfields("item_id"),rowcounter) %>');">
											Select</a>
										</font></font></td>
									</tr>
									<% cline = cline + 1 %>
								<% Next %>
								<% End If %>
							<% End If %>
					
							<tr>
								<td colspan="5"><font face="Arial" size="2"><b>&nbsp;</b></font></td>
							</tr>
						</table>
					</td>
				</tr>	
			</table>
		</td>
	</tr>
</table>

</body> 
</html>
