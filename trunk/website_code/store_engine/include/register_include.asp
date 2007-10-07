<%
function fn_is_logged_in()
    if cid > 0 and fn_get_querystring("Protected") = "" then
        fn_redirect Switch_Name&"login_thank_you.asp"
    end if
end function

%>