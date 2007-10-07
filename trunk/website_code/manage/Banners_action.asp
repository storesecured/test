<!--#include file="Global_Settings.asp"-->

<%
If not CheckReferer then
   Response.Redirect "admin_error.asp?message_id=2"
end if

'ERROR CHECKING
If Form_Error_Handler(Request.Form) <> "" then 
	Error_Log = Form_Error_Handler(Request.Form)
	%><!--#include file="Include/Error_Template.asp"--><% 
else

	BID = request.form("BID")
	Banner_Name = checkStringForQ(request.form("Banner_Name"))
	Image_Path = replace(replace(request.form("Image_Path"),"""",""),"'","")
	URL = replace(replace(request.form("URL"),"""",""),"'","")
	Start_Date = request.form("Start_Date")
	End_Date = request.form("End_Date")
	RunDays = request.form("RunDays")
	StartHour = request.form("StartHour")
	EndHour = request.form("EndHour")
	Enabled = request.form("Enabled")
	if not isdate(Start_Date) then
		Start_Date = dateserial(year(now()), month(now()), day(now()))
	end if
	if not isdate(End_Date) then
		End_Date = dateadd("m",1,Start_Date)
	end if
	response.write Enabled
	if Enabled = "on" then
		Enabled = -1
	else
		Enabled = 0
	end if

     if request.form("Action")="Add" then
		sql_add = "insert into store_banners (enabled, store_id, Banner_Name, Image_Path, URL, Start_Date, End_Date, RunDays, StartHour, EndHour) values ("&Enabled&", "&store_id&", '"&Banner_Name&"', '"&Image_Path&"', '"&URL&"', '"&Start_Date&"', '"&End_Date&"', '"&RunDays&"', "&StartHour&", "&EndHour&")"
	Else
		sql_add = "update store_banners set Enabled = "&Enabled&", Banner_Name = '"&Banner_Name&"', Image_Path = '"&Image_Path&"', URL = '"&URL&"', Start_Date = '"&Start_Date&"', End_Date = '"&End_Date&"', RunDays = '"&RunDays&"', StartHour = "&StartHour&", EndHour = "&EndHour&" where store_id="&store_id&" and Bann_ID="&BID
   end if
  session("sql") = sql_add

	conn_store.execute sql_add

	response.redirect "banners.asp"
end if
%>
