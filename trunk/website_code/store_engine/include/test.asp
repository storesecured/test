<%=Request.ServerVariables("SERVER_NAME")%>
<%
Response.Write( Session.SessionID )
response.write "1session shopper_id="&Session("Shopper_Id")
 %>