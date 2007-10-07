<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!-- BEGIN Monitor Tracking Variables  -->
<script language="JavaScript1.2">
if (typeof(tagVars) == "undefined") tagVars = "";
tagVars += "&VISITORVAR!Store_Id=<%= Store_Id %>&VISITORVAR!Service_Type=<%= Service_Type %>&VISITORVAR!Site_Name=<%= Site_Name %>&VISITORVAR!Trial_Version=<%= Trial_Version %>&VISITORVAR!Server=0&VISITORVAR!Store_Domain=<%= Store_Domain %>&VISITORVAR!Store_Email=<%= Store_Email %>";
</script>
<!-- End Monitor Tracking Variables  -->

<!-- BEGIN HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG --><script language='javascript' src='https://server.iad.liveperson.net/hc/7400929/x.js?cmd=file&file=chatScript3&site=7400929&&category=en;woman;5'> </script><!-- END HumanTag Monitor. DO NOT MOVE! MUST BE PLACED JUST BEFORE THE /BODY TAG -->

<%

sTitle = "Get Help"
thisRedirect = "livechat.asp"
Site_Name = Site_Name2
Secure = Replace(Site_Name,"http://","https://")
sMenu="none"
createHead thisRedirect  

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

if Trial_Version then
   Trial="-Trial"
else
   Trial=""
end if
   %>




      <tr bgcolor='#FFFFFF'><td align=center colspan=2><B>What kind of help would you like?</b></td>
	  <tr bgcolor='#FFFFFF'><TD align=center>Chat with an operator</td><td>
	  
	  
	 <!-- BEGIN LivePerson Button Code --><a href='https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&byhref=1&imageUrl=https://manage.storesecured.com/images/liveperson' target='chat7400929'  onClick="javascript:window.open('https://server.iad.liveperson.net/hc/7400929/?cmd=file&file=visitorWantsToChat&site=7400929&imageUrl=https://manage.storesecured.com/images/liveperson&referrer='+escape(document.location),'chat7400929','width=472,height=320');return false;" ><img src='https://server.iad.liveperson.net/hc/7400929/?cmd=repstate&site=7400929&&ver=1&imageUrl=https://manage.storesecured.com/images/liveperson' name='hcIcon' width=165 height=60 border=0></a></a><!-- END LivePerson Button code -->
</td></tr>
      <tr bgcolor='#FFFFFF'><td align=center>Search the knowledgebase</td><td><a href=http://server.iad.liveperson.net/hc/s-7400929/cmd/kbresource/front_page!PAGETYPE class=link target=_blank>Search</a></td></tr>
      <tr bgcolor='#FFFFFF'><td align=center>Post a technical request</td><td><a href=support_list.asp>Post Request</a></td></tr>
      

      

<% createFoot thisRedirect, 0 %>


