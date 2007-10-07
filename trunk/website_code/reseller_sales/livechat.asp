<%
title = "Easystorecreator Contact Form"
description = "Merchant account provider offers merchant account processor functions. Integrates with many merchant account service provider companies. Apply for your small business merchant account today. We'll completely setup your online store for you."
keyword1="merchant account provider"
keyword2="merchant account processor"
keyword3="merchant account service provider"
keyword4="small business merchant account"
keyword5=""

include_extra_links = false
include_credit = false
tracking_page_name="contactus"

sDate = now()
iDay = datepart("w",sDate)
if iDay = 1 then
   sDay = "SUN"
elseif iDay = 2 then
   sDay = "MON"
elseif iDay = 3 then
   sDay = "TUE"
elseif iDay = 4 then
   sDay = "WED"
elseif iDay = 5 then
   sDay = "THU"
elseif iDay = 6 then
   sDay = "FRI"
elseif iDay = 7 then
   sDay = "SAT"
end if
sTime=formatdatetime(sDate,4)
sDateTime= sDay & " " & sTime
   %>
<!--#include file="header.asp"-->
<p class="bodytext">

<table>
      <TR><TD colspan=5 align=center><BR>Times listed below are not a guarantee but rather a guideline to when a support technician is normally available for live chat.<BR></td></tr>
      <TR><TD colspan=5 align=center><BR><BR></td></tr>
      <tr><td><B>Available 5am-11pm 7 days a week</td><TD rowspan=3 align=left>
	  

	  <!-- BEGIN LivePerson Button Code --><table border='0' cellspacing='2' cellpadding='2'><tr><td align="center"></td><td align='center'><a href='https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&byhref=1' target='chat7400929'  onClick="javascript:window.open('https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&referrer='+escape(document.location),'chat7400929','width=472,height=320');return false;" ><img src='https://server.iad.liveperson.net/hc/7400929/?cmd=repstate&site=7400929&&ver=1&category=en;woman;5' name='hcIcon' width=165 height=60 border=0></a></td></tr></table><!-- END LivePerson Button code -->
      </td></tr>
      <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>

      <TR><TD>&nbsp;</td><td></td><td rowspan=2>Click on the livechat graphic<BR>to the left to start the chat</td></tr>


      </table>
      </p>

<!--#include file="footer.asp"-->
