<!--#include file="include/header.asp"-->
<%

If Enable_affiliates = 0 Then
    fn_redirect Switch_Name&"store.asp"
Else
%>
	 
		<table><tr><td class='normal'
		Welcome to the <%= Store_name %> Affiliate Program.<br><br>
		<% If Affiliate_type = 0 Then %>
			We pay affiliates <%= Currency_Format_Function(Affiliate_amount) %> per click.<br>
		<% Else %>
			We pay affiliates <%= Affiliate_amount %> % of each referred sale.<br>
		<% end if %>
		Affiliates are paid after accruing at least <%= Currency_Format_Function(Affiliate_payout) %> in payment. <br>

		<a href="<%= Site_Name %>affiliate_signup.asp" class='link'>Click here to signup now.</a>
		<br><br>
		<a href="http://affiliates.easystorecreator.com/affiliates_login.asp?Store_id=<%= Store_id %>" class='link'>If you are already signed up click here to login.</a>
		<% If Screen_affiliates = 0 Then %>
			All affiliate signups are automatically approved and may begin immediately referring people.
		<% Else %>
			All new signups must be screened by store management for approval.
		<% end if %>
		</td></tr></table>
	<% end if %>




	

<!--#include file="include/footer.asp"-->
