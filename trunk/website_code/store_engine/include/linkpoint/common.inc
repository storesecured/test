<!-- #include file=cc_defs.inc -->
<%
' Opaque contexts
Dim OrderCtx
Dim ItemCtx
Dim OptionCtx
Dim ReqCtx
Dim ShippingCtx
Dim TaxCtx
Dim CheckCtx
Dim CheckorderCtx
Dim ApiDriver

' Deaclarations

Dim bname
Dim bcompany
Dim baddr1
Dim baddr2
Dim bstate
Dim bcity
Dim bcountry
Dim bzip

Dim sname
Dim scompany
Dim saddr1
Dim saddr2
Dim scity
Dim sstate
Dim scountry
Dim szip
Dim scarrier

Dim phonenumber 
Dim faxnumber
Dim emailaddress
Dim comments

Dim cardnumber
Dim expmonth
Dim expyear

Dim subtotal
Dim tax
Dim shipping
Dim total

Dim startdate
Dim periodicity
Dim installments
Dim pbthreshold
Dim pbcomments

Dim accttype
Dim checknum
Dim trouting
Dim micr

Dim prog
Dim progqty

' Common functions

Function SetLocalVars

bname	= Request.Form("bname")
bcompany= Request.Form("bcompany")
baddr1	= Request.Form("baddr1")
baddr2	= Request.Form("baddr2")
bstate	= Request.Form("bstate")
bcity	= Request.Form("bcity")
bcountry= Request.Form("bcountry")
bzip	= Request.Form("bzip")

sname	= Request.Form("sname")
scompany= Request.Form("scompany")
saddr1	= Request.Form("saddr1")
saddr2	= Request.Form("saddr2")
scity	= Request.Form("scity")
sstate	= Request.Form("sstate")
scountry= Request.Form("scountry")	
szip	= Request.Form("szip")
scarrier= Request.Form("scarrier")

phonenumber = Request.Form("phone")
faxnumber   = Request.Form("fax")	
emailaddress= Request.Form("emailaddress")
comments    = Request.Form("comments")

cardnumber  = Request.Form("cardnumber")
expmonth    = Request.Form("expmonth")	
expyear     = Request.Form("expyear")	

subtotal    = Request.Form("subtotal")	
tax         = Request.Form("tax")
shipping    = Request.Form("shipping")
total       = Request.Form("total")

pbthreshold = Request.Form("pbthreshold")
startdate   = Request.Form("startdate")	
periodicity = Request.Form("periodicity")
installments= Request.Form("installments")
pbcomments  = Request.Form("pbcomments")

accttype    = Request.Form("accttype")
checknum    = Request.Form("checknum")
trouting    = Request.Form("trouting")
micr        = Request.Form("micr")

prog        = Request.Form("prog")
progqty     = Request.Form("progqty")
End Function

Function LeadingZero( InpStr )
if Len(InpStr) = 1 Then
   LeadingZero = "0"&InpStr
Else
   LeadingZero = InpStr
end if   
End Function
%>
