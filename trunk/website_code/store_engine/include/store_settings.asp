<!--#include virtual="common/connection.asp"-->
<%
sServerPath=lcase(server.MapPath("/"))
sSitePathArray=split(sServerPath,"\")
for each sFolder in sSitePathArray
    if instr(sFolder,"files_")>0 then
        Store_id = replace(sFolder,"files_","")
        exit for    
    end if
next

%>
<!--#include virtual="common/get_settings.asp"-->
<%

If strServerPort<> 443 then
	Switch_Name = Site_Name
	sHttp="http://"
else
	Switch_name = Secure_Site
	sHttp="https://"
end if

sURL = sHttp&Request.ServerVariables("SERVER_NAME")&"/"
sPathInfo = Request.ServerVariables("PATH_INFO")
arPathInfo = split(sPathInfo, "/")
fn_print_debug "sPathInfo="&sPathInfo
sRewritePath = Request.ServerVariables("HTTP_X_REWRITE_URL")
sLength=len(sRewritePath)
if sLength>0 then
sRewritePath=right(sRewritePath,sLength-1)
end if
fn_print_debug "sRewritePath="&sRewritePath
arFileNameWithQueryString = split(arPathInfo(Ubound(arPathInfo,1)), "?")
CurrentFileName = LCase(arFileNameWithQueryString(0))
if CurrentFileName="custom_page.asp" then
   CurrentFileName=fn_get_querystring("page_name")&".asp"
elseif CurrentFileName="custom.asp" then
   CurrentFileName=fn_get_querystring("page_name")&".asp"
end if
sQueryString=Request.ServerVariables("QUERY_STRING")

ReturnTo_Orig = Switch_Name&sRewritePath
fn_print_debug "ReturnTo_Orig="&ReturnTo_Orig
ReturnTo = Server.URLEncode(ReturnTo_Orig)
fn_print_debug "ReturnTo="&ReturnTo

%><!--#include file="sub.asp"-->
<!--#include file="session.asp"--><%

if Overdue_Payment>60 and overdue_payment<>100 then
    fn_redirect "http://www.easystorecreator.com/notfound.asp"
elseif (Store_active =0 or Overdue_Payment>7) and CurrentFileName <> "store_closed.asp" and CurrentFileName <> "404.asp" and CurrentFileName <> "preview.asp" and Store_id <> "" then
	'store is closed, not permanent
	fn_redirect Switch_name&"Store_Closed.asp"
elseif request.form("Shopper_Id")<>"" and (request.form("Shopper_Id")<>Session("Shopper_Id")) and isNumeric(request.form("Shopper_Id")) then
    'shopper doesnt have cookies enabled and cant be tracked, not permanent
    fn_redirect Switch_name&"cookies.asp"
elseif CurrentFileName = "store.asp" and Store_Homepage <> "" and not isNull(Store_Homepage) and store_homepage<>"http://"&Request.ServerVariables("server_name") and instr(Store_Homepage,"store.asp")=0 and instr(Store_Homepage,"http")=1 and store_homepage<>"http://"&Request.ServerVariables("server_name")&"/" then
	ReturnTo_Orig = Store_Homepage
	sRedirect=1
elseif lcase(Switch_Name)<>lcase(sURL) then
    sRedirect=1
else
    sRedirect=0
end if

'if the site is not using the default url permanently redirect to new url
if sRedirect=1 and ReturnTo_Orig<>"http://store.asp" and ReturnTo_Orig<>"http://preview.easystorecreator.com/store.asp"  then
    'response.write ReturnTo_Orig
    ReturnTo_Orig=replace(ReturnTo_Orig,"store.asp","")
    fn_redirect_perm ReturnTo_Orig
end if

Page_id=fn_get_querystring("page_id")
if not isNumeric(Page_id) or Page_id = "" then
	Page_id = 999
end if

'end reset of session variables
const EXPRESS_CHECKOUT_CUSTOMER = "--ExpressCheckOut Customer--"



Url_string = "Store_id/"&Store_id
'Url_string=""

' The following pages require a valid cid and we wont let the user 
' in unless such one exists ...

sLoginPages="budget_view_cust.asp,before_payment.asp,before_payment_action.asp,esd_download.asp,past_order_detail.asp,print_receipt.asp,recipiet.asp,register_thank_you.asp,modify_my_account.asp,modify_my_billing.asp,modify_my_shipping.asp,past_orders.asp"
fn_print_debug "cid="&cid

if (cid = 0) and Is_In_Collection(sLoginPages,CurrentFileName,",") then
	' redirect user to login page ...
	fn_redirect Switch_name&"check_out.asp?ReturnTo="&ReturnTo
elseif cid > 0 and (CurrentFileName = "register.asp" or CurrentFileName = "check_out.asp") then
	'redirect to thank you for registering page
	if fn_get_querystring("Protected") = "" then
	    fn_redirect Switch_name&"login_thank_you.asp"
	end if
end if

' in case shopping cart is empty redirect back to cart ...
if (CurrentFileName = "payment.asp" or CurrentFileName = "before_payment.asp" ) then
    if fn_Is_Cart_Full()=0 then
        fn_redirect Switch_name&"Show_Big_Cart.asp?Cart_Empty=True&Shopper_Id="&Shopper_id
    end if
end if

if (cid = 0) and (store_public=0 and CurrentFileName <> "check_out.asp" and CurrentFileName <> "check_out_action.asp" and CurrentFileName <> "store.asp" and CurrentFileName <> "contactus.asp" and CurrentFileName <> "forgot.asp" and CurrentFileName <> "forgot_thank_you.asp" and CurrentFileName <> "contact_action.asp" and CurrentFileName <> "check_out_action.asp" and CurrentFileName <> "error.asp" and CurrentFileName <> "privacy.asp" and CurrentFileName <> "returns.asp") and Request.Form("Form_Name") <> "Check_Out" and Store_active <> 0 and CurrentFileName <> "register.asp" and CurrentFileName <> "register_action.asp" and CurrentFileName <> "preview.asp" and CurrentFileName <> "form_email.asp" and CurrentFileName <> "form_error.asp" then
	fn_redirect Secure_Site&"check_out.asp?store_public="&Store_public&"&ReturnTo="&server.urlencode(CurrentFilename)
end if

%>
