<%
host = replace(lcase(request.servervariables("HTTP_HOST")),"www.","")
   Affiliate_ID = "35"


 CurrAffiliate = Request.Cookies("EASYSTORECREATOR")("Affiliate")
if CurrAffiliate = "" or CurrAffiliate="None" then
  Affiliate = Affiliate_ID
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
	
	str_site_host = split(host,".")

'***************************************************STARTS HERE**********************************************************
'code added by DEvki Anote here to reterive the keywords corresponding to the page name and the reseller's site 

'code here to set up the connection in the database
'------------------------------------------------------------------------------------------------------------------------
%>
<!--#include file = "code/include/headeroutside.asp"-->
<%	


'code here to retriive the resellerid and name of the site
'------------------------------------------------------------------------------------------------------------------------
Response.Write "host"&host

Response.Write "here"&str_site_host(0)
Response.End



set rsName  = conn.Execute("select fld_Reseller_ID,fld_Website  from tbl_Reseller_Master where fldWebsite = '"&str_site_host(0)&"'")
Response.Write ("select fld_Reseller_ID,fld_Website  from tbl_Reseller_Master where fldWebsite = '"&str_site_host(0)&"'")

if rsName.eof then 
	Response.Write "Not Found"
	
end if


if not rsName.eof then 
	
		Name = trim(rsName("fld_Website"))
		intresellerid = trim(rsName("fld_Reseller_ID"))
end if
'Response.Write "Name"&Name
'Response.Write "intresellerid"&intresellerid
'------------------------------------------------------------------------------------------------------------------------
%> 