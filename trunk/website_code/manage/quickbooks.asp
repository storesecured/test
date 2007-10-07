<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

sTitle = "Quickbooks"
thisRedirect = "quickbooks.asp"
sMenu="general"
createHead thisRedirect  %>



		<TR bgcolor='#FFFFFF'><td>
                <B>One-time Setup Instructions</b>
                <UL>
                <LI>1. Ensure Quickbooks is installed on your local computer
                <LI>2. Install the <B><a href=http://videos.storesecured.com/quickbooks/qbsdk50.exe target=_blank class=link>Quickbooks SDK 5.0</a></b>
                <LI>3. Extract the files in this <B><a href=http://videos.storesecured.com/quickbooks/bin.zip target=_blank class=link>zip file</a></b> in your c:\bin directory
                <LI>4. In your browser go to Tools ---> Internet Options --->security(tab) --> click on custom level button
                <BR>When you click custom level button security setting dialog box will open from that
                <BR>Enable  ActiveX controls and Plug-Ins
                <BR>Enable  Initialize and script Activex Controls not marked as safe.
                <BR>Enable  Run ActiveX controls and Plug-Ins
                <BR>Enable  Scripts Activex Controls marked safe for Scripting
                </UL>
                <B>Each Time You Export You Must</b>
                <UL>
                <LI>1. Ensure that the Quickbooks Program is running on your local computer
                <LI>2. When asked whether to give permission to access your Quickbooks files choose, Yes, this time
                <LI>3. Ensure that related customers and items are exported to Quickbooks BEFORE exporting orders.
                </UL>
                Once you have completed the above please hit your browsers back button to return to the previous screen and select the Insert to Quickbooks button to continue.
                If you are exporting a large amount of data it may take a while before you see anything happening.
                </td></tr>


<% createFoot thisRedirect, 0 %>


