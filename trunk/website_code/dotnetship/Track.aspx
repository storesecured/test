
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="dotnetSHIP" %>
<script language="VB" runat="server">

Dim objShip As dotnetSHIP.Ship

Dim TrackNo, Shipper As String

Sub Page_Load(Src As Object, E As EventArgs)

	
		objShip = new dotnetSHIP.Ship()
		objShip.USPSLogin = "876EASYS2859,610EY45TV868"  ' Please enter  your USPS login information here.
    objShip.UPSLogin = "melanieeasystore,ankle237,FBBA98A10FDA0F48"

            
End Sub

</script>

<html>
 
  <tr><td>Tracking Data:</td></tr>
  
   <%
    objShip.Track(Request.Params("shipcompany"), Request.Params("TrackNumber"))			

    If Len(objShip.Error) < 1 Then
      response.redirect(Request.Params("Site")+"trackdisplay.asp?Result="+server.urlencode(objShip.DisplayTracking()))
    Else
      response.redirect(Request.Params("Site")+"trackdisplay.asp?Error="+server.urlencode(objShip.Error))
    End If
   %>


</html>
