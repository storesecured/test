<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Canada Post Instructions"
thisRedirect = "canada_post_instructions.asp"
sMenu="shipping"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up Canada Post</td></tr>
		<TR bgcolor='#FFFFFF'><TD>
           Email eparcel@canadapost.ca to request a merchant ID for the XML system.
           <BR><BR>
           Make sure that your merchant ID is setup for the Production server and not Test server.
		 <BR><BR>  
           Canada Post sets up most accounts on the test server to start so you must specifically request to have it moved.


      </td></tr>


<% createFoot thisRedirect, 0 %>


