<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "FedEx Instructions"
thisRedirect = "fedex_instructions.asp"
sMenu="shipping"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>Instructions for setting up FedEx</td></tr>
		<TR bgcolor='#FFFFFF'><TD>
                <UL><LI>Get a FedEx account number to obtain rates from FedEx. 
                <LI>Email websupport@fedex.com requesting to have your account number added to the 
                test server, include in this email your FedEx account number,
                full company and contact details.  The subject line should be FedEx Account.
                <LI>Within 1 business day you will receive an email from FedEx stating that your account
                is going to be added to the test servers within 24 hours.
                <LI>After the 24 hours have passed please request a FedEx Meter Number by filling out the 
                <a href=fedex_request.asp target=_blank class=link>Fedex Meter Request Form</a>
                <LI>Take the meter number received and your FedEx Account number and enter them 
                in the required boxes on the realtime setup screen and enable FedEx
                <LI>Please ensure that you have entered the state/province for your store as the 2 
                character state abbreviation, ie CA not California for your Store Contact information and 
                for any extra shipping locations you have added.  FedEx will return an error if 
                you have entered an invalid state abbreviation.
                </UL>
                <a href=shipping_class_realtime.asp>Return to shipping setup</a>


      </td></tr>


<% createFoot thisRedirect, 0 %>


