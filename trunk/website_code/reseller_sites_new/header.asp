<%
host = replace(lcase(request.servervariables("HTTP_HOST")),"www.","")
 site_host = "easystorecreator.com"
  logo = "easystorecreator.jpg"
  Name = "easystorecreator"
if host = "clickmerchant.com" then
   site_host = "easystorecreator.com"
   logo = "logoclickmerchant.jpg"
   Name = "ClickMerchant"
   Affiliate_ID = "35"
elseif host = "abetterwaystore.com" then
   'site_host = "abetterwaystore.com"
   'logo = "abetterway.gif"
   'Name = "ABetterWayStore"
   'Affiliate_ID = "36"
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
if instr(Referrer,host) > 0 or instr(Referrer,"prosera.com") > 0 or Referrer = "" then
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

'***************************************************STARTS HERE**********************************************************
'code added by DEvki Anote here to reterive the keywords corresponding to the page name and the reseller's site 

'code here to set up the connection in the database
'------------------------------------------------------------------------------------------------------------------------
%>
<!--#include file = "include/headeroutside.asp"-->
<%	




'retreiing the name of the page that is being excecuted
tracking_page_name = trim(tracking_page_name)&".asp"  'this will be the hardcoded name in each page


'code here to retriive the resellerid and name of the site
'------------------------------------------------------------------------------------------------------------------------
'code commented here for the local
'**********************************************
if trim(Request.QueryString("reseller"))<>"" then 
	intresellerid =trim(Request.QueryString("reseller"))
	session("ResellerSiteID") = intresellerid	
end if

if intresellerid="" then
	set rsName  = conn.Execute("select fld_Reseller_ID,fld_Website from tbl_Reseller_Master where fld_Website ='"&host&"'")
	if not rsName.eof then
		intresellerid = trim(rsName("fld_Reseller_ID"))
		fld_Website  = trim(rsName("fld_Website"))
	end if
end if

if intresellerid ="" then
	 intresellerid	= session("ResellerSiteID")
end if
if session("ResellerSiteID")="" then
	session("ResellerSiteID")= intresellerid
end if

'------------------------------------------------------------------------------------------------------------------------
dim rsName,intresellerid,rsPageName,pageid,keyword1,keyword2,keyword3,keyword4,keyword5,sqlProc
dim sqlProcDesc,getimagename,rsProcDesc,rsProc,description,logoimagename
Dim strQuery

strQuery = "select fld_Reseller_ID,fld_Website,fld_Display_Name from tbl_Reseller_Master where fld_Reseller_ID ="&session("ResellerSiteID")
'set rsName  = conn.Execute("select fld_Reseller_ID,fld_Website from tbl_Reseller_Master where fld_Website = '"&sPath(ubound(sPath)-1)&"'")
if intresellerid <> "" then
'	set rsName  = conn.Execute("select fld_Reseller_ID,fld_Website from tbl_Reseller_Master where fld_Reseller_ID ="&session("ResellerSiteID")&"")
	set rsName  = conn.Execute(strQuery)
	if not rsName.eof then
		Site_Name = trim(rsName("fld_Website"))
		Name = trim(rsName("fld_Display_Name"))
		site_host = Name
		intresellerid = trim(rsName("fld_Reseller_ID"))
		session("SiteResellerSiteID") = intresellerid 
		 session("ResellerSiteID") = intresellerid

	else
		Response.Redirect "http://www.storesecured.com/notfound.asp"
		Response.End
		
		
	end if
	
end if


'Response.Write "Name"&Name
'if trim(Request.QueryString("action")) = "" then
'	intresellerid="1"
'	session("ResellerSiteID") = intresellerid
'end if 
'**********************************************


'code here to retreive the page id corresponing to that particula page name


Set rsPageName = Server.CreateObject("ADODB.Recordset")
With rsPageName
	.Source = "Get_Page_Id '"&trim(tracking_page_name)&"'"
	.ActiveConnection = strConn
	.CursorType = adOpenForwardOnly
	.Open
End With

if not rsPageName.eof then 
	pageid = trim(rsPageName("fld_page_id"))
end if  
'------------------------------------------------------------------------------------------------------------------------

'code here to reterive the keywords corresponding to that particular id 

'------------------------------------------------------------------------------------------------------------------------



sqlProc = "Get_Reseller_Page_Key " &intresellerid&" ,"&pageid
set rsProc = conn.Execute(sqlProc)
if not rsProc.eof then
		keyword1 = trim(rsProc("fld_Reseller_Page_Keyword1"))
		keyword2 = trim(rsProc("fld_Reseller_Page_Keyword2"))
		keyword3 = trim(rsProc("fld_Reseller_Page_Keyword3"))
		keyword4 = trim(rsProc("fld_Reseller_Page_Keyword4"))
		keyword5 = trim(rsProc("fld_Reseller_Page_Keyword5"))
end if 
'------------------------------------------------------------------------------------------------------------------------

'code here to retreive the description corresponding to that particular page id

'------------------------------------------------------------------------------------------------------------------------
sqlProcDesc = "Get_Reseller_Page_Desc  " &intresellerid&" ,"&pageid
set rsProcDesc  = conn.Execute(sqlProcDesc)
if not rsProcDesc.eof then 
	description = trim(rsProcDesc("fld_desc_text"))
end if 
'------------------------------------------------------------------------------------------------------------------------


'code here to retreive the LOGO image set by the reseller
'------------------------------------------------------------------------------------------------------------------------
getimagename="select fld_title_image from tbl_reseller_logo where fld_reseller_id="&intresellerid
set rsgetlogo=conn.execute(getimagename)
if not rsgetlogo.eof then
	logoimagename=trim(rsgetlogo("fld_title_image"))
end if 
'------------------------------------------------------------------------------------------------------------------------

if logo_image_name="" then
   logo_image_name=logo
end if
'*********************************************** ENDS HERE**************************************************************

'********************************Code starts here******************************
dim strget,planrate,i,planid
redim planrate(0),planid(0)

'***********************code added by rashmi*******************************************************
redim preserve planrate(1)
redim preserve planrate(2)
redim preserve planrate(3)
redim preserve planrate(4)
redim preserve planrate(5)
planrate(1)="9.95"
planrate(2)="19.95"
planrate(3)="29.95"
planrate(4)="39.95"
planrate(5)="59.95"

'************************code ends here******************************************************
'Code added by Sudha Ghogare on 13 th August 2004 to show reseller rates
if intResellerId <> "" then 
for i=1 to 5
			strget="Get_reseller_plan "&intResellerId&" ,"&i&""
			set rsget=conn.execute(strget)
			if not rsget.eof then
				redim preserve planrate(i)
				planrate(i)=formatnumber(trim(rsget("fld_rate")),2)
			end if
	
next	
end if
'*********************************Code ends here******************************

'code added here to calculate the yearly pricing 
'*******************code starts here****************************************
term = 12

dim Yearplanrate
redim Yearplanrate(0)
redim preserve Yearplanrate(1)
redim preserve Yearplanrate(2)
redim preserve Yearplanrate(3)
redim preserve Yearplanrate(4)
redim preserve Yearplanrate(5)

if cint(term) = 12 then
	for i=1 to 5
		Yearplanrate(i) = formatnumber(planrate(i) * 12 * .75,2)
	next 	
end if
'*******************code Ends here****************************************
%>
<html>
<head>
<title><%= title %></title>
<meta name="description" content="<%= description %>">
<meta name="keywords" content="<%= keyword1 %>, <%= keyword2 %><% if keyword3 <> "" then %>, <%= keyword3 %><% end if %><% if keyword4 <> "" then %>, <%= keyword4 %><% end if %><% if keyword5 <> "" then %>, <%= keyword5 %><% end if %>">
<link rel="stylesheet" href="http://www.<%= host %>/style.css" type="text/css">
<script src="script.js" language="JavaScript" type="text/javascript"></script>
</head>
<!-- BEGIN HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG --><script language='javascript' src='https://server.iad.liveperson.net/hc/7400929/x.js?cmd=file&file=chatScript3&site=7400929&&category=en;woman;1'> </script><!-- END HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG -->

<body style="margin: 0px; background-image: URL(images/bg.gif);">
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
<td><a href="http://www.<%= host %>"
            onmouseout="onoff('home','off')" 
            onmouseover="onoff('home','over')"><img src="images/homeoff.gif" width="50" height="27" border="0" alt="home" name="home"></a></td>
<td><a href="http://www.<%= host %>/aboutus.asp"
            onmouseout="onoff('about','off')" 
            onmouseover="onoff('about','over')"><img src="images/aboutoff.gif" width="59" height="27" border="0" alt="about us" name="about"></a></td>
<td><a href="http://www.<%= host %>/privacy.asp"
            onmouseout="onoff('priv','off')" 
            onmouseover="onoff('priv','over')"><img src="images/privoff.gif" width="65" height="27" border="0" alt="privacy" name="priv"></a></td>
<td><a href="http://www.<%= host %>/sitemap.asp"
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

<td valign="top" width="197"><a href=http://www.<%= host %>/><img alt="<%= keyword3 %>" src="logos/<%=logoimagename%>" width="197" height="119" border="0"></a></td>

<td valign="top" width="545">
<table border="0" cellspacing="0" cellpadding="0" width="545">
<tr>
<td><img alt="<%= keyword4 %>" src="images/menuleft_home.gif" width="58" height="27" border="0"></td>
<td><a href="http://www.<%= host %>/shopping_cart_features.asp"
            onmouseout="onoff('features','off')" 
            onmouseover="onoff('features','over')"><img src="images/featuresoff.gif" width="92" height="27" border="0" alt="features" name="features"></a></td>
<td><a href="http://www.<%= host %>/shopping_cart_signup.asp"
            onmouseout="onoff('signup','off')" 
            onmouseover="onoff('signup','over')"><img src="images/signupoff.gif" width="92" height="27" border="0" alt="sign up now" name="signup"></a></td>
<td><a href="http://www.<%= host %>/ecommerce_solution_samples.asp"
            onmouseout="onoff('cust','off')"
            onmouseover="onoff('cust','over')"><img src="images/custoff.gif" width="150" height="27" border="0" alt="featured customers" name="cust"></a></td>
<td><a href="http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/front_page!PAGETYPE" target=_blank
            onmouseout="onoff('faq','off')" 
            onmouseover="onoff('faq','over')"><img src="images/faqoff.gif" width="54" height="27" border="0" alt="faq" name="faq"></a></td>
<td><a href="http://www.<%= host %>/ecommerce_contactus.asp"
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
<img src="images/bodytop_inner.gif" width="547" height="28" border="0"><br>

<br>

<!-- content area begins -->

<table border="0" cellspacing="0" cellpadding="0" width="520">
<tr>
<td valign="top" class="bodytext" width="520" style="padding-left: 10px;">
