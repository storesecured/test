<!--#include virtual="common/connection.asp"-->


<%

  'LOG OUT THE STORE ADMINISTRATOR FROM STORE MANAGER
  Session("Store_Id") = ""
  Session("Super_User") = 0
  Session("Login_Privilege") = 0
  Session("Department_List:"&Store_Id)=""
  Session("Location_List:"&Store_Id)=""
  Response.Redirect "login_store.asp"

%>
