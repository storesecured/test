<%

sOrigUrl = "www.easystorecreator.com"
sDefaultUrl = fn_url_rewrite (sOrigUrl)
Name="EasyStoreCreator"
logoimagename="logo_inner.gif"

dim strget,planrate,i,planid
redim planrate(5),planid(0)

planrate(1)="9.95"
planrate(2)="19.95"
planrate(3)="29.95"
planrate(4)="39.95"
planrate(5)="67"
Name = "EasyStoreCreator"

host = lcase(Request.ServerVariables("SERVER_NAME"))

if host<>sDefaultUrl and "www."&host<>sDefaultUrl then
    host = replace(host,sLocalAddName,"")
    host = replace(host,"www.","")
    %>
    <!--#include file="reseller_include.asp"-->
    <%
else
    intresellerid=8   
end if

dim Yearplanrate
redim Yearplanrate(5)

for i=1 to 5
	Yearplanrate(i) = formatnumber(planrate(i) * 12 * .75,2)
next 	

 CurrAffiliate = Request.Cookies("EASYSTORECREATOR")("Affiliate")
if CurrAffiliate = "" or CurrAffiliate="None" then
  Affiliate = request.querystring("Id")
  if Affiliate <> "" then
		response.cookies("EASYSTORECREATOR")("Affiliate") = Affiliate
		response.cookies("EASYSTORECREATOR").expires = DateAdd("d",60,Now())
  else
		response.cookies("EASYSTORECREATOR")("Affiliate") = "None"
		response.cookies("EASYSTORECREATOR").expires = DateAdd("d",60,Now())
  end if
end if
Referrer = request.ServerVariables("HTTP_REFERER")
if instr(Referrer,sDefaultUrl) > 0 or Referrer = "" then
else
   CurrReferrer = Request.Cookies("EASYSTORECREATOR")("Referrer")
   if CurrReferrer <> "" then
      Referrer = CurrReferrer & ", " & Referrer
   end if
	response.cookies("EASYSTORECREATOR")("Referrer") = Referrer
	response.cookies("EASYSTORECREATOR").expires = DateAdd("d",60,Now())
end if

if quicktext="" then
   quicktext ="<a class=title href=shopping_cart.asp><strong>Shopping Cart Software</strong></a>: Online store and <a class=title href=ecommerce_solution_samples.asp>ecommerce solutions</a>."
end if
if title="" then
   title="Shopping Cart Software, Ecommerce Solutions, Online Store Builder, "&Name
end if
if description="" then
  description="Easy Ecommerce solutions! The "&Name&" shopping cart software allows you to easily build an online store from any web browser"
end if
if keywords="" then
   keywords="Shopping cart software, Ecommerce solutions, Online store Builder, Shopping cart, E-commerce software"
end if
if keyword1="" then
   keyword1="Shopping cart software"
end if
if keyword2="" then
   keyword2="Ecommerce solutions"
end if
if keyword3="" then
   keyword3="Online store Builder"
end if

response.charset="ISO-8859-1"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= title %></title>
<meta name="description" content="<%= description %>">
<meta name="keywords" content="<%= keyword1 %>, <%= keyword2 %><% if keyword3 <> "" then %>, <%= keyword3 %><% end if %><% if keyword4 <> "" then %>, <%= keyword4 %><% end if %><% if keyword5 <> "" then %>, <%= keyword5 %><% end if %>">
<link rel="stylesheet" href="http://<%=sDefaultUrl %>/style.css" type="text/css">
<script src="include/script.js" language="JavaScript" type="text/javascript"></script>
</head>

