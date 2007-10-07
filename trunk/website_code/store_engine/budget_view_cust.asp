
<!--#include file="include/header.asp"-->
<%
select_data="select a.budgetid,convert(char(12), a.budget_date,101) as budget_date,a.Budget_Amount,a.Store_id ,a.cid,a.Notes,b.budget_orig,b.budget_left  from Store_Budget_History a ,store_customers b where a.cid= b.cid and b.cid="&cid&"and b.record_type=0 and b.store_id="& store_id
	set rs_store = conn_store.execute (select_data)
	if not rs_store.eof then
	%>
	<table width="100%" border="1" cellspacing="0" cellpadding=2>
		<tr bgcolor=#DDDDDD>
			
			<TD class=0><B>Date</B></TD>
			<TD class=0><B>Amount</B></TD>
			<TD class=0><B>Note</B></TD>
		</TR>

	<%
	varclass= 1
	do while not rs_store.eof
		strlcl_budgetid=rs_store("budgetid") 
		strlcl_budget_date=rs_store("budget_date") 
		strlcl_Budget_Amount=rs_store("Budget_Amount") 
		intlcl_Cid=rs_store("cid") 
		intlcl_budget_orig=rs_store("budget_orig") 
		intlcl_budget_left =rs_store("budget_left") 
		intlcl_budget_note =rs_store("Notes") 
		if varclass=1 then
			varclass=0
		else
			varclass=1
		end if
	%>
	<TR class=<%=varclass%>>
		<TD><%=strlcl_budget_date%></td>
		<TD><%=strlcl_Budget_Amount%></td>
		<TD><%=intlcl_budget_note%></td>
	</tr>
<%
	rs_store.movenext
	loop
	
	
 %>
<TABLE>
<% end if %>
<!--#include file="include/footer.asp"-->
