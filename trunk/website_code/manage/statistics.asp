<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sTitle = "Store Statistics"
thisRedirect = "statistics.asp"
sMenu="statistics"
createHead thisRedirect

response.redirect Site_Name2&"logs/stats.pl"

createFoot thisRedirect, 0 %>


