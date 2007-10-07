
<!--#include file="include/sub.asp"-->
<%

str = Request.Form

Send_Mail "mblack@easystorecreator.com","mblack@easystorecreator.com","Store " & Store_ID&" Cancel","Store has cancelled recurring paypal payment.<BR>"&str


%>

