<%
host = replace(lcase(request.servervariables("HTTP_HOST")),"www.","")

host = split(host,".")
if ubound(host)> 1 then 
	site = host(1)
else
			site = host(0)
end if 
%>
<!--#include file = "include/headeroutside.asp"-->
<HTML>
<HEAD>
<title><%= title %></title>
<meta name="description" content="<%= description %>">
<meta name="keywords" content="<%= keyword1 %>, <%= keyword2 %><% if keyword3 <> "" then %>, <%= keyword3 %><% end if %><% if keyword4 <> "" then %>, <%= keyword4 %><% end if %><% if keyword5 <> "" then %>, <%= keyword5 %><% end if %>">
<link href="images/styles.css" rel="stylesheet" type="text/css">
<% if includejs =1 then %>
<script src="script.js" language="JavaScript" type="text/javascript"></script>
<% end if %>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>

</HEAD>
<!-- BEGIN HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG --><script language='javascript' src='https://server.iad.liveperson.net/hc/7400929/x.js?cmd=file&file=chatScript3&site=7400929&&category=en;woman;1'> </script><!-- END HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG -->

<BODY bgcolor="#C0C0C0" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/nav_overview_on.gif','images/nav_included_on.gif','images/nav_earn_on.gif','images/nav_rebranding_on.gif','images/nav_login_on.gif','images/nav_signup_on.gif')">

<table width="775" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="1" rowspan="2" background="images/pxl_black.gif"><img src="images/pxl_black.gif" width="1" height="1"></td>
    <td background="images/tile_navigation.gif">
      <table width="773" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="195"><a href="default.asp"><img src="images/logo.gif" border="0" width="195" height="82"></a></td>
          <td align="right" valign="middle">
            <table width="543" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="62"><a href="overview.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image26','','images/nav_overview_on.gif',1)"><img src="images/nav_overview_off.gif" alt="Overview" name="Image26" border="0" width="62" height="20"></a></td>
                <td width="100"><a href="whatsincluded.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image27','','images/nav_included_on.gif',1)"><img src="images/nav_included_off.gif" alt="Whats Included" name="Image27" border="0" width="100" height="20"></a></td>
                <td width="128"><a href="earnings.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image28','','images/nav_earn_on.gif',1)"><img src="images/nav_earn_off.gif" alt="How Much can I earn" name="Image28" border="0" width="128" height="20"></a></td>
                <td width="82"><a href="sample-sites.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image29','','images/nav_samplesites_on.gif',1)"><img src="images/nav_samplesites_off.gif" alt="Sample Sites" name="Image29" border="0" width="82" height="20"></a></td>
                <td width="76"><a href="rebranding.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image30','','images/nav_rebranding_on.gif',1)"><img src="images/nav_rebranding_off.gif" alt="Rebranding" name="Image30" border="0" width="76" height="20"></a></td>
                <td width="44"><a href="http://managereseller.<%=site%>.com" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image31','','images/nav_login_on.gif',1)"><img src="images/nav_login_off.gif" alt="Login" name="Image31" border="0" width="44" height="20"></a></td>
                <td><a href="https://reseller.<%=site%>.com/resellersignup.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image32','','images/nav_signup_on.gif',1)"><img src="images/nav_signup_off.gif" alt="SignUp" name="Image32" border="0" width="51" height="20"></a></td>
              </tr>
            </table>
          </td>
          <td width="5"><img src="images/tile_navigation.gif" width="5" height="82"></td>
        </tr>
      </table>
    </td>
    <td width="1" rowspan="2" background="images/pxl_black.gif"><img src="images/pxl_black.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td align="left" valign="top" background="images/tile_silver_undernav.gif"><img src="images/tile_silver_undernav.gif" width="9" height="18"></td>
  </tr>
</table>
<table width="775" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="1" background="images/pxl_black.gif"><img src="images/pxl_black.gif" width="1" height="1"></td>
    <td>
      <table width="773" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="284"><a href="whatsincluded.asp"><img src="images/img_top_1.gif" width="284" height="69" border="0"></a></td>
          <td width="316"><a href="earnings.asp"><img src="images/img_top_2.gif" width="316" height="69" border="0"></a></td>
          <td rowspan="5"><img src="images/img_basket.jpg" width="173" height="170"></td>
        </tr>
        <tr> 
          <td><a href="contactus.asp"><img src="images/img_toll_free.gif" width="284" height="35" border="0"></a></td>
          <td><a href="ecommerce-tutorial.asp"><img src="images/img_10day_free.gif" width="316" height="35" border="0"></a></td>
        </tr>

        <tr>
          <td><img src="images/rounded_left.gif" width="284" height="15"></td>
          <td><img src="images/rounded_right.gif" width="316" height="15"></td>
        </tr>
        <tr>
          <td><a href=https://reseller.<%=site%>.com/resellersignup.asp><img src="images/img_join_now.gif" border=0 width="284" height="51"></a></td>
          <td><a href="earnings.asp"><img src="images/img_how_much_earn.gif" width="316" height="51" border=0></a></td>
        </tr>
      </table>
      <table width="773" border="0" cellspacing="0" cellpadding="0" bgcolor=#D0CEA8>
        <tr>
          <td align="left" valign="top" bgcolor="#D0CEA8">
           <table wdith=773 border=0 cellspacing=0 cellpadding=10><tr><td>

