<!--#include file="include/header.asp"-->
<%
shipCompany = fn_get_querystring("shipCompany")
tracking = fn_get_querystring("tracking")


%>


<form method="POST" action="http://dotnetshiporig.easystorecreator.com/Track.aspx">
<input type=hidden name=Site value="<%= Site_Name %>">
 <table>
      <tr><td>Track Number: <input type="text" name="TrackNumber" size="15" value="<%= tracking %>"></td</tr>

      <% 
      if shipCompany="Fedex" then
        fn_redirect "http://www.fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers="&tracking
      elseif shipCompany="DHL" then
        fn_redirect "http://track.dhl-usa.com/TrackRslts.asp?nav=TrackBynumber&ShipmentNumber="&tracking
      elseif shipCompany = "" then %>
     <tr><td>Shipping Company:
           <select size="1" name="shipCompany">
	         <option value="USPS">USPS</option>
           <option value="UPS">UPS</option>
           <option value="CanadaPost">CanadaPost</option>
           </select>
         </td>
      </tr>

      <% else %>
      <input type=hidden name=shipCompany value=<%= shipCompany %>>

      <% end if %>
      <tr><td width="397" colspan="2">
         <input type="submit" value="Get Tracking data" name="TrackBut">
        </td>
      </tr>
  </table>
</form>

<!--#include file="include/footer.asp"-->
