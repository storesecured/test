<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Fedex"
thisRedirect = "fedex.asp"
sMenu="none"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>Fedex realtime rates re-enabled</td></tr>
                <TR bgcolor='#FFFFFF'><TD>
                <B>Update 7/23 11:20 AM PST, FedEx rates now updated and working</b><BR>
                FedEx rates are now enabled.  Please note that FedEx now requires all users to 
                have a account number and meter number and be certified.  Please see the FedEx instructions
                on the <a href=shipping_class_realtime.asp class=link>realtime shipping</a> setup page for more information.
                <BR><BR>
                <B>Update 7/20 8:30 AM PST</b><BR>
                The application has now been approved by Fedex and we are completing final testing 
                of the new integration.  We expect to implement the new Fedex rate API this weekend.
                <BR><BR>
                <B>Update 6/29 11:30 AM PST</b><BR>
                Please note that fedex has updated their programming API.  We are currently 
                working on getting reapproved to provide their rates.  The certification process 
                should take another 1-2 weeks, at which time the Fedex rates will be reactivated again.
                <BR><BR>
                <B>Important Update 6/18 9:00PM PST</b><BR>
                Fedex realtime rates are not working at the moment.  
                It appears that the Fedex api has changed.  We are contacting 
                Fedex for updated details and if needed a new release will be
                issued as soon as it is available.  Any additional information 
                will be posted here as soon as it is available.  
                <BR><BR>
                <B>Recommendations for those using Fedex</b>
                <BR>
                It is unknown at this time how many updates will be required or 
                what changes Fedex has made.  It is recommended to disable Fedex 
                realtime rates for the time being.  They should be temporarily 
                replaced with realtime rates for another provider or calculated rates.
                We have no further information at this time.
                </td></tr>


<% createFoot thisRedirect, 0 %>


