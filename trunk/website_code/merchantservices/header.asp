<%
 sFileName = lcase(request.ServerVariables("script_name"))
 CurrAffiliate = Request.Cookies("EASYSTORECREATOR")("Affiliate")
if CurrAffiliate = "" then
  Affiliate = request.querystring("Id")
  if Affiliate <> "" then
		response.cookies("EASYSTORECREATOR")("Affiliate") = Affiliate
		response.cookies("EASYSTORECREATOR").expires = DateAdd("d",60,Now())
  else
		response.cookies("EASYSTORECREATOR")("Affiliate") = "None"
		response.cookies("EASYSTORECREATOR").expires = DateAdd("d",60,Now())
  end if
end if
%>
<HTML>
<HEAD>
<title><%= title %></title>
<meta name="description" content="<%= description %>">
<meta name="keywords" content="<%= keyword1 %>, <%= keyword2 %><% if keyword3 <> "" then %>, <%= keyword3 %><% end if %><% if keyword4 <> "" then %>, <%= keyword4 %><% end if %><% if keyword5 <> "" then %>, <%= keyword5 %><% end if %>">
<link rel="stylesheet" type="text/css" href="include/style.css">

<script language="javascript" type="text/javascript">
<!--
function menu_goto( menuform )
{
   // see http://www.thesitewizard.com/archive/navigation.shtml
   // for an explanation of this script and how to use it on your
   // own site

   selecteditem = menuform.newurl.selectedIndex ;
   newurl = menuform.newurl.options[ selecteditem ].value ;
   if (newurl.length != 0) {
      location.href = newurl ;
   }
}
//-->
</script>
</HEAD>
<!-- BEGIN HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG --><script language='javascript' src='https://server.iad.liveperson.net/hc/7400929/x.js?cmd=file&file=chatScript3&site=7400929&&category=en;woman;1'> </script><!-- END HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG -->

<body bgcolor="#EFEFE7">
<form name="form1">
  <table width="780" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#EFEFE7">
    <tr> 
      <td><table width="780" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="322"><a href="default.asp"><img src="images/logo.gif" width="322" height="76" border="0"></a></td>
            <td align="right" valign="middle"><table width="254" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td></td>
<td width="5">&nbsp;</td>
                  <td width="47"><a href="default.asp"><img src="images/btn_home_up.gif" width="47" height="17" border="0"></a></td>
                  <td width="15">&nbsp;</td>
                  <td width="78"><a href="contactus.asp"><img src="images/btn_contact_up.gif" width="78" height="17" border="0"></a></td>
                  <td width="5">&nbsp;</td>
                  <td width="72"><a href="affiliate-program.asp"><img src="images/btn_affiliates_up.gif" width="72" height="17" border="0"></a></td>
                  <td width="5">&nbsp;</td>
                  <td width="42"><a href="links.asp"><img src="images/btn_links_up.gif" width="42" height="17" border="0"></a></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td align="left" valign="top" background="images/tile_orange_large.jpg" class="topmenutext"><img src="images/tile_orange_large.jpg" width="8" height="19" align="absmiddle"> 
        <a href="merchant-accounts.asp" class="topmenutext">Merchant Account Basics</a> | <a href="merchant-account-fees.asp" class="topmenutext">Processing
        Fees</a> | <a href="merchant-account-calculator.asp" class="topmenutext">Comparison Calculator</a> |
        <a href="merchant-account-fraud.asp" class="topmenutext">Fraud Prevention</a> | <a href="electronic-payments.asp" class="topmenutext">Electronic
        Payments</a> | <a href="recommended-merchant-accounts.asp" class="topmenutext">Our Rates</a> | 
        <a href="merchant-account-terminology.asp" class="topmenutext">Common Terms</a></td>
    </tr>
    <tr> 
      <td align="center" valign="top" background="images/tile_blue_large.jpg"> 
        <table width="760" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="514" align="left" valign="middle" class="topmenutext"><strong>Alternatives:</strong> 
              <a href="third-party-merchant-accounts.asp" class="topmenutext">Third Party Accounts</a> | <a href="
              person-to-person-accounts.asp" class="topmenutext">Person 2 Person Accounts</a></td>
            <td width="2" rowspan="2"><img src="images/blue_separator.gif" width="2" height="57"></td>
            <td width="243" align="right" valign="middle" class="topmenutext"><strong>EasyMerchantServices</strong> 
              - Quick Navigation!</td>
          </tr>
          <tr> 
            <td width="514" align="left" valign="top"><table width="296" border="0" cellspacing="0" cellpadding="0">
                <tr class="topmenutext"> 
                  <td width="138">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td width="138"><a href="compare-merchant-accounts.asp"><img src="images/btn_whybuy_up.gif" width="138" height="17" border="0"></a></td>
                  <td><a href="low-rate-guarantee.asp"><img src="images/btn_lowrates.gif" width="158" height="17" border="0"></a></td>

                </tr>
              </table></td>
            <td align="right" valign="middle"> <select name="newurl" onChange="menu_goto(this.form)">
                <option value="merchant-accounts.asp"
                  <% if sFileName = "/merchant-accounts.asp" then
                   response.write " selected "
                end if %>> Merchant Account Basics</option>
                <option value="merchant-account-fees.asp"
                  <% if sFileName = "/merchant-account-fees.asp" then
                   response.write " selected "
                end if %>> Processing Fees</option>
                <option value="merchant-account-calculator.asp"
                  <% if sFileName = "/merchant-account-calculator.asp" then
                   response.write " selected "
                end if %>>Comparison Calculator</option>
                <option value="merchant-account-fraud.asp"
                  <% if sFileName = "/merchant-account-fraud.asp" then
                   response.write " selected "
                end if %>>Fraud Prevention</option>
                <option value="electronic-payments.asp"
                  <% if sFileName = "/electronic-payments.asp" then
                   response.write " selected "
                end if %>>Electronic Payments</option>
                <option value="recommended-merchant-accounts.asp"
                  <% if sFileName = "/recommended-merchant-accounts.asp" then
                   response.write " selected "
                end if %>>Our Rates</option>
                <option value="merchant-account-terminology.asp"
                  <% if sFileName = "/merchant-account-terminology.asp" then
                   response.write " selected "
                end if %>>Common Terms</option>
                <option>===================</option>
                <option value="third-party-merchant-accounts.asp"
                  <% if sFileName = "/third-party-merchant-accounts.asp" then
                   response.write " selected "
                end if %>>Third Party Accounts </option>
                <option value="person-to-person-accounts.asp"
                  <% if sFileName = "/person-to-person-accounts.asp" then
                   response.write " selected "
                end if %>>Person 2 Person Accounts</option>
                <option>===================</option>
                <option value="default.asp"
                  <% if sFileName = "/default.asp" or sFileName = "/" then
                   response.write " selected "
                end if %>>Home</option>
                <option value="contactus.asp"
                  <% if sFileName = "/contactus.asp" then
                   response.write " selected "
                end if %>>Contact Us</option>
                <option value="affiliate-program.asp"
                  <% if sFileName = "/affiliate-program.asp" then
                   response.write " selected "
                end if %>>Affiliates</option>
                <option value="links.asp"
                  <% if sFileName = "/links.asp" then
                   response.write " selected "
                end if %>>Links</option>
                <option value="privacy.asp"
                  <% if sFileName = "/privacy.asp" then
                   response.write " selected "
                end if %>>Privacy Policy</option>
              </select> </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
<table width="780" align="center" cellpadding="0" cellspacing="0"><TR><TD>




