 <%
const COOKIE_NAME = "EASYSTORECREATOR"

'SAVE LOGIN AND SHOPPER_ID INFORMATION USING A COOKIE
sub saveUserCookie(theUserName, theUserPassword)
	response.cookies(COOKIE_NAME)("UserSet") = "TRUE"
	response.cookies(COOKIE_NAME)("UserName") = theUserName
	response.cookies(COOKIE_NAME)("UserPassword") = theUserPassword
	response.cookies(COOKIE_NAME).Expires = DateAdd("d",30,Now())
end sub

'CHECK IF THERE IS A COOKIE SAVED ON CUSTOMER COMPUTER
function isUserCookie()
	if request.cookies(COOKIE_NAME)("UserSet") <>"" then
		isUserCookie = true
	else
		isUserCookie = false
	end if
end function

'REMOVE AN EXISTING COOKIE SAVED ON CUSTOMER COMPUTER
sub deleteUserCookie()
	response.cookies(COOKIE_NAME).Expires = DateAdd("d",-1,now()) 
end sub

'LOAD INFORMATION STORED USING COOKIE
sub getUserCookie(theUserName, theUserPassword)
	theUserName = request.cookies(COOKIE_NAME)("UserName")
	theUserPassword = request.cookies(COOKIE_NAME)("UserPassword")
end sub

%>
