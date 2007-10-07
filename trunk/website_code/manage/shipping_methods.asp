<!--#include file="Global_Settings.asp"-->

<%

if Shipping_Class <> "" then

	if Shipping_Class>=1 and Shipping_Class<=6 then
		Response.Redirect "shipping_class"&Shipping_Class&"_list.asp"&sAddString
	else
		Response.Redirect "shipping_class.asp"&sAddString
	end if
else
	response.write "Please define a shipping method first."
end if
%>

