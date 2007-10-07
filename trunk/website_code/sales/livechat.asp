<%
title = "Live Chat"
description = "Build a website, with automatic daily froogle updates."
keyword1="froogle"
keyword2="google shopping"
keyword3="froogle intregration"
keyword4="shopping directory"
keyword5=""

include_extra_links = true
include_credit = false
tracking_page_name="froogle"

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
      <table>
      <tr><td><TD>
	  

<!-- BEGIN LivePerson Button Code --><a href='https://server.iad.liveperson.net/hc/7400929/?cmd=file&amp;file=visitorWantsToChat&amp;site=7400929&amp;byhref=1&amp;imageUrl=https://manage.storesecured.com/images/liveperson' target='chat7400929'  onClick="javascript:window.open('https://server.iad.liveperson.net/hc/7400929/?cmd=file&amp;file=visitorWantsToChat&amp;site=7400929&amp;imageUrl=https://manage.storesecured.com/images/liveperson&amp;referrer='+escape(document.location),'chat7400929','width=472,height=320');return false;" ><img src='https://server.iad.liveperson.net/hc/7400929/?cmd=repstate&amp;site=7400929&amp;ver=1&amp;imageUrl=https://manage.storesecured.com/images/liveperson' name='hcIcon' width=165 height=60 border=0 alt=""></a><!-- END LivePerson Button code -->
<BR>

</td></tr>
      

      </table>

<!--#include file="footer.asp"-->
