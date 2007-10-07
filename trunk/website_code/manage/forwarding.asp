<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%


sTitle = "Forward blocking to AOL.com anf Yahoo.com"
thisRedirect = "forwarding.asp"
sMenu="none"
createHead thisRedirect  %>


      
		<TR bgcolor='#FFFFFF'><td><B>No forwarding allowed to AOL and Yahoo</td></tr>
		<TR bgcolor='#FFFFFF'><TD>
                <B>Notice posted August 15, 2006 2:35PM PST</b>
                <BR>Please note that effective immediately our mailservers will no longer
                allow mail to be forwarded to any email address at AOL.com or at Yahoo.com
                <BR><BR>
                If you currently have a forward setup to an address at either AOL.com or 
                Yahoo.com your forwards have been ceased indefinately and you will need to 
                either remove the forward or choose a new email address to forward to.  
                <BR><BR>
                <B>What is a email forward?</b>
                <BR>
                An email forward is when you indicate that any mail received at a certain address 
                is to be automatically sent on to another address.  Ie similar to a forward when you 
                move houses.  Ie perhaps I have an address at aol, widgets@aol.com and I have an
                email address with storesecured, and its support@widgets.com.  Now at this point I
                have 2 separate email inboxes.  Some people only want to check 1 email inbox and 
                because of this they ask us to forward a copy of any mail received at the support@widgets.com
                address to their widgets@aol.com address, now they are only looking inside of the aol 
                inbox for their mail.  If this is not familiar to you then you most likely never setup 
                a forward and would not be impacted by this action.
                <BR><BR>
                <B>Why was this done?</b>
                <BR>
                Both AOL and Yahoo have mechanisms for reporting email as spam.  If you ask 
                StoreSecured to forward mail to your AOL or Yahoo address and then subsequently 
                mark that email that we have forwarded, at your request, as spam that is a black 
                mark against StoreSecured.com and we are seen as a spammer.  For everytime that 
                this occurs our spam score at these companies goes higher and the possibility of 
                legitamate emails being received is reduced.  AOL and Yahoo refuse to indicate who marks
                these emails as spam and the problem has become very prevalent recently with 
                thousands of such emails being returned each day and our status as a legitamate email
                sender in jeopardy.  At this point we have taken the step to not allow any forwards to 
                both domain names to protect ourselves and you our customer.
                <BR><BR>
                <B>Does this effect my ability to send emails to AOL and Yahoo?</b>
                <BR>
                No, it does not affect regular email sending to either of these companies in 
                any way, shape or form you can still send regular emails to your clients at 
                AOL and Yahoo just as you always could before.  This does not effect your 
                ability to send mail to those providers including automatic notifications 
                emails.  This covers only direct forwards.  If you are not sure what a mail forward is
                then most likely this does not apply to you.
                <BR><BR>
                <B>What do I do now if I was forwarding to AOL or Yahoo?</b>
                <BR>
                If you currently have a forward to AOL or Yahoo it has been disabled and no emails will be 
                sent.  You may either forward to a new email address or check your mail directly on the 
                StoreSecured servers using either webmail or a mail client like Outlook etc.
                <BR><BR>
                <B>Is this action permanent?</b>
                <BR>
                In the future if either AOL or Yahoo change this policy of who to mark as the spammer we 
                will revisit this issue at that time.  However under the current AOL and Yahoo policies 
                we are unable to allow forwarding.
                <BR><BR>
                <B>My mail isnt being forwarded to AOL or Yahoo but its still not working?</b>
                <BR>
                Please note that some email providers do use AOL and/or Yahoo services for their mail
                without you even realizing it is occuring.   If you are unsure if this is what is happening 
                in your situation you may perform a short test to see for sure.  Go to 
                <a href=http://www.webmaster-toolkit.com/mx-record-lookup.shtml target=_blank>http://www.webmaster-toolkit.com/mx-record-lookup.shtml</a>
                and type in your email address and select the button labelled Lookup Mx Records, it 
                will give a list of servers, if you see aol or yahoo listed then that email is 
                provided by that company and any forwards to it would also be blocked.
                <BR><BR>
                <B>How do I view or modify my forwarding information?</b>
                <BR>
                To view or modify your forward please visit the StoreSecured Webmail at 
                <a href="http://mail.storesecured.com" target=_blank>http://mail.storesecured.com</a> 
                Enter your username and password to login.  Under the Settings menu at the top please 
                choose My Settings, then select the Forwarding Tab.  If you are forwarding your email 
                the forward address will be shown in the box marked Forwarding Address.  Please make any 
                required changes and hit the Save button.
                <BR><BR>
                <B>What will happen to my mail if it isnt forwarding?</b>
                <BR>
                If you have a forward setup to AOL or Yahoo, your mail will no longer be forwarded it will
                be as if the forward does not exist and your mail can be read directly from the StoreSecured 
                servers.
                <BR><BR>
                <B>Can I request an exception?</b>
                <BR>
                Unfortunately no, the block that is in place on forwarding is a global block (meaning it affects all customers).  
                We do not have the ability to open it up for individual users at this time.  If in the future the 
                mailserver technology allows exceptions to this block we may offer the ability to bypass this setting to customers
                willing to guarantee in a written statement that they would not mark such email as spam.   However, until such time
                as that is possible within our mailserver any such individual requests for exceptions will be denied.
                <BR><BR>
                <B>Are there any other alternatives?</b>
                <BR>
                Yes, with Yahoo you may be able to setup your account at
                tto pull the mail off of our servers and into your account.  This capability is provided under the heading
                check other mail.  Please direct
                any and all questions regarding how to do this, to Yahoo support.  We are
                unable to provide assistance in how to setup this sort of configuration as it is done from
                the Yahoo site only.  At this time AOL has indicated that they do not provide this facility.
                <BR><BR>
                <B>Note to customers</b>
                <BR>
                While we realize that this may be seen as an inconvenience for some of our merchants
                our end goal is to ensure that your emails and all of our valued customers emails 
                reach their destination and are not impeded in any way shape or form.  Your ability to 
                email your customers and keep in contact with them is of the utmost concern to our staff 
                and this measure was taken to ensure that this indeed happens.  AOL and Yahoo are both 
                extremely large players in the email market and a blockage by either company would be
                determental to all of our merchants including yourself.   This was seen as a necessary 
                step to ensure the smooth flow of emails between our company, all alternative avenues 
                were investigated and discussed with support representatives at both Yahoo and AOL before proceeding
                with this extreme measure.
                <BR><BR>
                Sincerely,
                <BR><BR>
                StoreSecured Support Staff
                </td></tr>


<% createFoot thisRedirect, 0 %>


