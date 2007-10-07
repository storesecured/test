<!--#include virtual="common/connection.asp"-->


<%

  'LOG OUT THE STORE ADMINISTRATOR FROM STORE MANAGER
  Session("Store_Id") = ""
  Session("Super_User") = 0
  Session("Login_Privilege") = 0
  Response.Redirect "login_store.asp"

%>
