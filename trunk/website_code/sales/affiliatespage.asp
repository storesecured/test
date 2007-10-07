<!--#include virtual="common/connection.asp"-->
<!--#include file="include\sub.asp"-->
<% on error goto 0 %>
<% Affiliate_ID = Session("Affiliate_ID") 
tracking_page_name="affiliatespage"
%>

<TABLE><TR><TD>Copy and paste the text in the box below onto your site to begin earning money today.</TD></TR>
<TR><TD><a href="http://<%=sDefaultUrl %>?Id=<%= Affiliate_ID %>">Put your business on the web today at <%=Name %></a><BR>
<textarea rows=4 cols=55><a href="http://<%=sDefaultUrl %>?Id=<%= Affiliate_ID %>">Put your business on the web today at <%=Name %></a></textarea><BR>
</td></tr>
<tr><td>
<a href="http://<%=sDefaultUrl %>?Id=<%= Affiliate_ID %>">Website builder, domain name and hosting all for one low monthly fee, at <%=Name %></a><BR>
<textarea rows=4 cols=55><a href="http://<%=sDefaultUrl %>?Id=<%= Affiliate_ID %>">Website builder, domain name and hosting all for one low monthly fee at <%=Name %></a></textarea>
</td></tr></table>

<table border=1 cellspacing=0 cellpadding=1>
<tr><td><B>Type</b></td><td><b>Sites</b></td></td></tr>
<%
sql_select = "select count(Site_Name) as sites from Store_Settings where Affiliate_Id = "&Affiliate_ID&" and Trial_Version=0 and Service_Type >= 7 and overdue_payment=0"
rs_Store.open sql_select,conn_store,1,1
accounts1 = rs_Store("sites")
rs_store.close

sql_select = "select count(Site_Name) as sites from Store_Settings where Affiliate_Id = "&Affiliate_ID&" and Trial_Version=0 and Service_Type < 7 and overdue_payment=0"
rs_Store.open sql_select,conn_store,1,1
accounts2 = rs_Store("sites")
rs_store.close

sql_select = "select count(Site_Name) as sites from Store_Settings where Affiliate_Id = "&Affiliate_ID&" and Trial_Version=1"
rs_Store.open sql_select,conn_store,1,1
accounts3 = rs_Store("sites")
rs_store.close
%>


	<tr><td>Gold and Premium Accounts</td><td><%= accounts1 %></td></tr>
	<tr><td>Bronze and Silver Accounts</td><td><%= accounts2 %></td></tr>
	<tr><td>Free Accounts*</td><td><%= accounts3 %></td></tr>
	<tr><td colspan=2>*Please note that free accounts which are inactive for a period
	of 4 weeks or more are removed from our servers and are not counted in the stats 
   above.</td></tr>
</table>
<BR><BR>
<% if accounts1 + accounts2 > 2 then
	payment_due = (accounts1 * 5) + (accounts2 * 2)
else
	 payment_due = 0
end if %>
<BR><BR>
<table><tr><td>Your affiliate payout for next month will be: $<%= payment_due %></td></tr></table>
<!-- BEGIN WEBSIDESTORY CODE v3.0.0 (pro)-->
<!-- COPYRIGHT 1997-2003 WEBSIDESTORY, INC. ALL RIGHTS RESERVED. U.S.PATENT No. 6,393,479 B1. Privacy notice at: http://websidestory.com/privacy -->
<script language="javascript">
var _pn="<%= tracking_page_name %>";	//page name(s)
var _mlc="CONTENT+CATEGORY";	//multi-level content category
var _cp="null";	//campaign
var _acct="WR531104J9MB93EN3";	//account number(s)
var _pndef="title";	//default page name
var _ctdef="full";	//default content category
var _prc="";	//commerce price
var _oid="";	//commerce order
var _dlf="";	//download filter
var _elf="";	//exit link filter
var _epg="";	//event page identifier
var _mn="wp110"; //machine name
var _gn="phg.hitbox.com"; //gateway name
var _hcv=68;function _wn(n){return((n.indexOf("NAME")>0&&n.indexOf("PUT")>-1)||
(n.indexOf("CONTENT")>-1&&n.indexOf("CATEGORY")>0))};function _gd(x,w){var _ed=
x.lastIndexOf("/");var _be=(w!="full")?x.lastIndexOf("/",_ed-2):x.indexOf("/");
return(_be==_ed)?"/":x.substring(_be,_ed);};function _gf(x){return x.substring(
x.lastIndexOf("/")+1,x.length);};var _hl=location;var _lp=_hl.protocol.indexOf(
'https')>-1?'https://':'http://';var _zo=(new Date()).getTimezoneOffset();
function _ps(_ip,_pml){return(_ip&&_wn(_pml))?((_pndef=="title"&&
document.title!="")?document.title:_gf(_hl.pathname)?_gf(_hl.pathname):_pndef):
(_wn(_pml)?_gd(_hl.pathname,_ctdef):_pml)};function _pm(m,_fml,h){if(m.indexOf(
";")!=-1){_nml=m.substring(0,m.indexOf(";"));_rm=m.substring(m.indexOf(";")+1,
m.length);_fml+=_ps(h,_nml)+";";return _pm(_rm,_fml,h);}else{return _fml+_ps(h,
m);}};var _sv=10;var _hn=navigator;var _bn=_hn.appName;if(_bn.substring(0,9)==
"Microsoft"){_bn="MSIE";};var _bv=(Math.round(parseFloat(_hn.appVersion)*100));
if((_bn=="MSIE")&&(parseInt(_bv)==2))_bv=301;function _ex(v){return escape(v)};
var _rf=_ex(document.referrer);_mlc=_pm(_mlc,"",0);_pn=_pm(_pn,"",1);</script>
<script language="javascript1.1">_sv=11</script><script language="javascript1.1" 
src="http://stats.hitbox.com/js/hbp.js" defer></script><script 
language="javascript">if(_sv<11){if(document.cookie.indexOf("CP=")>-1){_ce="y";}
else{document.cookie="CP=null*; path=/; expires=Wed, 1 Jan 2020 00:00:00 GMT";
_ce=(document.cookie.indexOf("CP=")!=-1)?"y":"n";};if((_rf=="undefined")||
(_rf=="")){_rf="bookmark";};_x2="<img src='"+_lp+_gn+"/HG?hc="+_mn+"&hb="+
_ex(_acct)+"&n="+_ex(_pn);_x3="&cd=1&hv=6' border=0 height=1 width=1>";_ar=
"&bn="+_ex(_bn)+"&bv="+_bv+"&ce="+_ce+"&ss=na&sc=na&ja=na&sv="+_sv+"&con="+
_ex(_mlc)+"&vcon="+_ex(_mlc)+"&epg="+_epg+"&hp=u&cy=u&dt=&ln=na&cp="+_ex(_cp)+
"&pl=&prc="+_ex(_prc)+"&oid="+_ex(_oid)+"&rf="+_rf;document.write(_x2+_ar+_x3);}
if(_bn!='Netscape'&&_bv>399)document.write("<\!"+"--");</script><noscript><img 
src="http://phg.hitbox.com/HG?hc=wp110&cd=1&hv=6&ce=u&hb=WR531104J9MB93EN3&n=<%= tracking_page_name %>&vcon=CONTENT+CATEGORY" border='0' width='1' height='1'></noscript><!--//-->
<!-- END WEBSIDESTORY CODE  -->

