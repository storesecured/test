
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="dotnetSHIP" %>
<%@ Page Language="VB" debug="true" %>
<script language="VB" runat="server" debug="true">

Dim objShip As dotnetSHIP.Ship

Dim TrackNo, Shipper As String

Sub Page_Load(Src As Object, E As EventArgs)

	
		objShip = new dotnetSHIP.Ship()
		objShip.USPSLogin = "876EASYS2859,610EY45TV868"  ' Please enter  your USPS login information here.
    'objShip.UPSLogin = Request.Params("UPS")

            
End Sub

</script>

<html>
 
  <tr><td>Tracking Data:</td></tr>
  
   <%
    objShip.Track("USPS", "7777")

    response.write(objShip.DisplayTracking())

   %>


</html>
