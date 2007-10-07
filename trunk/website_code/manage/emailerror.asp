<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Email Server Upgrade"
thisRedirect = "2checkout.asp"
Secure = "https://"&Secure_name&"/"
sMenu="advanced"
createHead thisRedirect  %>


      
		<TR><TD>
		<B>Mailserver Re-Enabled</b><BR><BR>
                 Sat 11/3 7:45AM<BR>For those of you who have been waiting the mailserver is now up and accepting connections.
		We will be shortly turning on the backup mailservers to deliver mail which has been collected over
		the evening.  Please give the backups serveral hours to catch up and deliver all mail that is waiting.
                We will also
		be performing final tuning on the server during the day today. If you have any problems,issues or concerns please dont
                hesistate to contact us.  We are sorry for the delay but unfortunately we experienced a  variety of
                problems moving over some mail accounts which cause the entire process to take much longer then anticipated.
                <BR><BR>Please check all of your settings on the new server by sending yourself a test email to ensure that 
                everything is setup properly with your new configuration.
		<BR><BR>
                <B>Frequently Asked Questions</b>
                <BR>1. Why am I re-receiving emails from earlier?
                <BR>If you are using a program like outlook to retrieve mail and leaving a copy of that mail on the 
                server outlook can tell if its a new mail based on the mailserver id and the email id.  Since we now 
                have a new mailserver outlook will be unable to match to any emails it had previously downloaded and 
                you may get copies.  Unfortunately there is no way to prevent this if using pop mail.  You may want to 
                consider imap instead if you regularly leave mail on the servers.
                <BR><BR>
                2. What is my login to manage my email accounts?
                <BR>This has changed slightly and is now managedomain@yourdomainname.com with the same password as before. 
                There is a login button which will take you directly there from General-->Email-->Add Email
                <BR><BR>
                3. I think I am missing some mails from last night?
                <BR>Please note that it will take several hours this morning for the backups to deliver all mail that 
                was captured over the time the main server was in maintenance.  Please be patient.  Most emails will come 
                fairly quickly but you may see some stragglers.
                <BR><BR>
                4. Why did it take so much longer then anticipated?
                <BR>We are very sorry for the inconvenience.  Unfortunately we did encounter several problems while 
                performing the switch which had to be dealt with individually and as you know it always takes longer 
                when you are learning something new.
                <BR><BR>5. Whats new about this server?
                <BR>The new servers webmail interface is quicker loading and has many more options.  You can setup spam
                filtering rules, view statistics, search for particular emails, and automatically delete messages when your mailbox
                hits a certain size.  In addition there is an online help manual to help you out when you get stuck.
                <BR>
		<HR>
                Original announcement posted since 11/19, this is just a reminder.
                <BR><BR>
                <B>What:</b> Mail Server Upgrade and Scheduled Downtime TONIGHT
                <BR><BR><B>Why:</b> The current mailserver cannot handle the load imposed on it and has been experiencing issues lately with uptime and duplication of emails.
                <BR><BR><B>When:</b> Fri December 2nd 10:00pm PST (tonight)
                <BR><BR><B>How long:</b> We are unsure of the exact timeframe but you should expect a downtime of at least 2-3 hours
                <BR><BR><B>Effect:</b> The main mailserver will not be accepting any connections during this time either
                for reading mail or accepting mail.  The backup mailserver will continue to collect your mail
                during this time so that it is not lost.  However, you will be unable to retrieve this mail
                until after the upgrade is complete.  Once the upgrade is complete the backup mailserver will begin to send 
                all queued mail to your main inbox for collection.
                <BR><BR><B>Outcome:</b> All settings and existing mail will be transferred to the new mailserver, those
                using an SMTP mail client like outlook or outlook express should see no difference.  Those 
                using the webbased client will see a new interface but will continue to use the same username
                and password to login.  We will also try to copy forwarding addresses, filters, auto replies etc but recommend full 
                checking on all extra services after the upgrade.
                <BR><BR><B>Other Services:</b> All other services including the admin interface, stores, ftp etc will
                continue to operate as normal and will be unaffected by this mailserver upgrade.
                <BR><BR><B>What can I do to help:</b> Please remove any unused mailboxes in your account and delete any old mail
                on the servers which you do not need.  The upgrade time length will be dependent on the total amount of 
                data to transfer.  By reducing your total number of emails saved on the server you will be helping to make the upgrade go quicker and more smoothly.
                <BR><BR><B>Common Questions:</b>
                <BR>Will my mailservers address change?  No, the address will not change
                <BR>Do I need to do anything before the change?  No, you are not required to do anything
                <BR>Will I lose any mail during this upgrade?  No, we will be copying over all existing mail
                <BR>Will my site be down during this time?  No, your site will not be down during this time
                <BR>Why was this time chosen?  Nights see the lowest level of activity and weekend nights are even lower then 
                weekdays.  Thus we have tried to choose a time and day which will affect the fewest number of people.
                </td></tr>


<% createFoot thisRedirect, 0 %>


