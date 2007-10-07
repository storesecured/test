<%@ Page Language="VB" Debug="true" Trace="false" Description="NetShip Component"%>
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="dotnetSHIP" %>
<script language="VB" runat="server">

Dim objShip As Ship ' dotnetSHIP object name

Sub Page_Load(sender As [Object], e As EventArgs)
   
   
   objShip = New dotnetSHIP.Ship() ' dotnetSHIP.ship
   
   'A meter number request is a one time operation.  Make note of this return value
   'and use the number returned in the FedExLogin property.
   'Before you can use the test servers your FedEx account number
   'must be setup on the test systems for the XML API.
   'For the test server please use following line:
   'objShip.FedExURL="https://gatewaybeta.fedex.com/GatewayDC"
   Dim strMeter As String = objShip.FedExGetMeterNumber(request.form("Account_No"), request.form("Name"), request.form("Phone"), request.form("Address"), request.form("City"), request.form("State"), request.form("Zip"), request.form("Country"))
   
   
   ' FedEx Account number
   ' Your name
   ' Phone number, no space or '-', only numbers
   ' Address
   ' City name
   ' State or Province
   ' Postal code
   ' Country code
   
   Response.Write(("Your Meter Number is " + strMeter))
End Sub 'Page_Load 

</script>


