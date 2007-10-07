<%

if Request.ServerVariables("SERVER_NAME")<>"www.easystorecreator.com" then
        arPathInfo = split(Request.ServerVariables("PATH_INFO"), "/")
        arFileNameWithQueryString = split(arPathInfo(Ubound(arPathInfo,1)), "?")
        CurrentFileName = LCase(arFileNameWithQueryString(0))

        response.status="301 Moved Permanently"
        Response.AddHeader "Location", "http://www.easystorecreator.com/"&CurrentFileName
end if

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
if instr(Referrer,"easystorecreator.com") > 0 or instr(Referrer,"prosera.com") > 0 or Referrer = "" then
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
   title="Shopping Cart Software, Ecommerce Solutions, Online Store Builder, EasyStoreCreator"
end if
if description="" then
  description="Easy Ecommerce solutions! The EasyStoreCreator shopping cart software allows you to easily build an online store from any web browser"
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
<link rel="stylesheet" href="http://www.easystorecreator.com/style.css" type="text/css">
<script src="script.js" language="JavaScript" type="text/javascript"></script>
<link rel="alternate" type="application/atom+xml" title="Atom" href="http://www.makingyourway.com/atom.xml" />
<link rel="alternate" type="application/rss+xml" title="RSS 1.0" href="http://www.makingyourway.com/index.rdf" />
<link rel="alternate" type="application/rss+xml" title="RSS 2.0" href="http://www.makingyourway.com/rss.xml" />
</head>

<body>
<div align="center">
<table border="0" cellspacing="0" cellpadding="0" width="780">

<tr>
<td valign="top" width="14"><img src="images/leftborder.gif" width="14" height="720" border="0" alt="<%= keyword1 %>"></td>

<td valign="top" width="752" bgcolor="#ffffff">
<table border="0" cellspacing="0" cellpadding="0" width="752">

<tr>
<td valign="top" width="752">
<table border="0" cellspacing="0" cellpadding="0" width="752">
<tr>
<td width="5" bgcolor="#ffffff" nowrap></td>

<td valign="middle" align="left" style="padding-left: 15px; background-image: URL(images/top_brown_back.gif);" class="whitebold" width="511" height="49"><h1><%= quicktext %></h1></td>

<td width="231" valign="top">
<table border="0" cellspacing="0" cellpadding="0" width="231">
<tr>
<td><a href="http://www.easystorecreator.com/"
            onmouseout="onoff('home','off')" 
            onmouseover="onoff('home','over')"><img src="images/homeoff.gif" width="50" height="27" border="0" alt="home" name="home"></a></td>
<td><a href="http://www.easystorecreator.com/aboutus.asp"
            onmouseout="onoff('about','off')" 
            onmouseover="onoff('about','over')"><img src="images/aboutoff.gif" width="59" height="27" border="0" alt="about us" name="about"></a></td>
<td><a href="http://www.easystorecreator.com/privacy.asp"
            onmouseout="onoff('priv','off')" 
            onmouseover="onoff('priv','over')"><img src="images/privoff.gif" width="65" height="27" border="0" alt="privacy" name="priv"></a></td>
<td><a href="http://www.easystorecreator.com/sitemap.asp"
            onmouseout="onoff('site','off')" 
            onmouseover="onoff('site','over')"><img src="images/siteoff.gif" width="57" height="27" border="0" alt="sitemap" name="site"></a></td>
</tr>

<tr>
<td valign="top" colspan="4" width="231"><img alt="<%= keyword2 %>" src="images/top_menu_inner_bot.gif" width="231" height="22" border="0"></td>
</tr>
</table>
</td>


<td width="5" bgcolor="#ffffff" nowrap></td>
</tr>
</table>
</td>
</tr>

<tr>
<td valign="top" width="752">
<table border="0" cellspacing="0" cellpadding="0" width="752">
<tr>
<td width="5" bgcolor="#ffffff" nowrap></td>

<td valign="top" width="197"><a href="http://www.easystorecreator.com/"><img alt="<%= keyword3 %>" src="images/logo_inner.gif" width="197" height="119" border="0"></a></td>

<td valign="top" width="545">
<table border="0" cellspacing="0" cellpadding="0" width="545">
<tr>
<td><img alt="<%= keyword4 %>" src="images/menuleft_home.gif" width="58" height="27" border="0"></td>
<td><a href="http://www.easystorecreator.com/shopping_cart_features.asp"
            onmouseout="onoff('features','off')" 
            onmouseover="onoff('features','over')"><img src="images/featuresoff.gif" width="92" height="27" border="0" alt="features" name="features"></a></td>
<td><a href="http://www.easystorecreator.com/shopping_cart_signup.asp"
            onmouseout="onoff('signup','off')" 
            onmouseover="onoff('signup','over')"><img src="images/signupoff.gif" width="92" height="27" border="0" alt="sign up now" name="signup"></a></td>
<td><a href="http://www.easystorecreator.com/ecommerce_solution_samples.asp"
            onmouseout="onoff('cust','off')"
            onmouseover="onoff('cust','over')"><img src="images/custoff.gif" width="150" height="27" border="0" alt="featured customers" name="cust"></a></td>
<td><a href="http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/front_page!PAGETYPE" target=_blank
            onmouseout="onoff('faq','off')" 
            onmouseover="onoff('faq','over')"><img src="images/faqoff.gif" width="54" height="27" border="0" alt="faq" name="faq"></a></td>
<td><a href="http://www.easystorecreator.com/ecommerce_contactus.asp"
            onmouseout="onoff('contact','off')" 
            onmouseover="onoff('contact','over')"><img src="images/contactoff.gif" width="99" height="27" border="0" alt="contact us" name="contact"></a></td>
</tr>

<tr>
<td valign="top" colspan="6"><img alt="<%= keyword5 %>" src="images/inner_header1.jpg" width="545" height="92" border="0"></td>
</tr>
</table>
</td>

<td width="5" bgcolor="#ffffff" nowrap></td>
</tr>
</table>
</td>
</tr>





<!-- body area begins -->
<tr>
<td valign="top" width="752">
<table border="0" cellspacing="0" cellpadding="0" width="752">
<tr>
<td width="5" bgcolor="#ffffff" nowrap></td>

<td valign="top" style="background-image: URL(images/bodyback_inner.gif);" width="547">
<img src="images/bodytop_inner.gif" width="547" height="28" border="0" alt=""><br>

<br>

<!-- content area begins -->

<table border="0" cellspacing="0" cellpadding="0" width="520">
<tr>
<td valign="top" class="bodytext" width="520" style="padding-left: 10px;">
