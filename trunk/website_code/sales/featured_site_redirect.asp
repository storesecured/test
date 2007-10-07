Please wait while we load the customers site.
<%


if request.querystring("site_url")<>"" then
   response.redirect request.querystring("site_url")
else
   response.redirect "shopping_cart_software.html"
end if
%>

