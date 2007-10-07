<%
title = "Shopping Cart Signup"
description = "Shopping cart software package"
keyword1="shopping cart"
keyword2="shopping cart software"
keyword3="ecommerce package"
keyword4="website builder"
keyword5="ecommerce store"

include_extra_links = false
include_credit = false
tracking_page_name="premium-signup"
includejs=1
%>
<!--#include file="header.asp"-->


<p class="header1">Store Signup</p>
Congratulations! You've taken the first step towards building a fantastic online site for yourself.
<BR><BR>Please completely fill out the form below to get started with your <%= Name %> site.
							&nbsp;<BR><BR>


			<form method="POST" target="_top" name=signup action="http://manage.storesecured.com/new_Store_Action.asp">
			<input type="Hidden" name="Form_Name" value="company">
                        <table border=1 cellspacing=0 cellpadding=3>
			<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Site Name</b></font></td>

			 <td width="98%" height="1">
					<input type="text" name="Store_name" size="30" maxlength=200>
					<input type="hidden" name="Store_name_C" value="Re|String|0|200|||Site Name"></td>
				</tr>
                        <tr>
					
					<td colspan=2 bgcolor=#DDDDDD>
                                        The site name will be placed at the top of all pages as part of your logo and on invoices.  This can be changed at anytime after signing up.</td>
                           </tr>

			<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Web Address</b></font></td>
					<td width="98%" height="1"><font size=1>http://
					<input type="text" name="Site_name" size="25" maxlength=63 onKeyPress="return goodchars(event,'0123456789abcdefghijklmnopqrstuvwxyz-')">.storesecured.com
					<input type="hidden" name="Site_name_C" value="Re|String|0|63||.,http,www,\,/,?,:,*,&,@,_|Web Address">
               <input type="hidden" name="site_host" value="storesecured.com"></td>
				</tr>
                            <tr>
					
					<td colspan=2 bgcolor=#DDDDDD>
                                        This is your FREE WEB ADDRESS to be used until the time you choose a domain name.
                                        Letters and numbers only please. No spaces or punctuation. &nbsp;
                                        Once your store is fully activated you may use your own domain name such as www.mystore.com</td>
                           </tr>

			<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Choose a Username</b></font></td>
					<td width="98%">
					<input type="text" name="Store_User_id" size="30">
					 <input type="hidden" name="Store_User_id_C" value="Re|String|0|20|||Username"></td>
				</tr>


			<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Choose a Password</b></font></td>
					<td width="98%">
					<input type="password" name="Store_password" size="30" maxlength=20>
					<input type="hidden" name="Store_password_C" value="Re|String|0|20|||Password"></td>
				</tr>


			<tr>

			<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Confirm the Password</b></font></td>
					<td width="98%">
					<input type="password" name="Store_password_confirm" size="30" maxlength=20>
					<input type="hidden" name="Store_password_confirm_C" value="Re|String|0|20|||Password Confirmation"></td>
				</tr>
			<tr>
					
					<td colspan=2 bgcolor=#DDDDDD>
                                        The Username and Password you choose will be used in future visits when you wish to login and make changes to your site.  Please keep this information in a safe place.
                                        Passwords must be at least 5 characters.</td>
                           </tr>

			 <tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Enter your Email</b></font></td>
					<td width="98%">
					<input type="text" name="Email" size="30" maxlength=50>
					<input type="hidden" name="Email_C" value="Re|Email|0|50|||Email">
					<input type="hidden" name="Affiliate" value="<%= Request.Cookies("EASYSTORECREATOR")("Affiliate") %>">
					<input type="hidden" name="Referrer" value="<%= left(Request.Cookies("EASYSTORECREATOR")("Referrer"),1000) %>">
</td>
				</tr>
			<tr>
					
					<td colspan=2 bgcolor=#DDDDDD>
                                        Your email will be placed on the contact us page of the site for customers to use to contact you.  It will also be used to send important notfications regarding your site 
                                        including notification on new orders.  This information can be changed at anytime.</td>
                           </tr>

				<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Payment Type</b></td>
					<td width="98%">
					<select name="Payment_Method">
                                        <option value="">Choose One</option>
										<option value="Visa">Visa</option>
										<option value="Mastercard">Mastercard</option>
										<option value="American Express">American Express</option>
										<option value="Discover">Discover</option>
										<option value="eCheck">eCheck (US Only)</option>
										<option value="Paypal">PayPal</option>
                                        
                                        
                                        </td>
				</tr>

				<input type=hidden name=Service_Level value=store>
					
				
				<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>Choose the Term</b></td>
					<td width="98%">
					<select name="term">
                                        <option value="">Choose one</option>
										<option value="1">Monthly</option>
										<option value="3">Quarterly (5% off)</option>
										<option value="6">Semi-Annual (10% off)</option>
										<option value="12">Yearly (25% off)</option>
					</option></td>
				</tr>
				<tr>
					
					<td colspan=2 bgcolor=#DDDDDD>
                                        No Risk, 45 Day Money Back Guarantee.  No questions asked.</td>
                           </tr>
                             <tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>How did you hear about us?</b></font></td>
					<td width="98%">
					<input type="text" name="Referred_By" size="30" maxlength=200>
					<input type="hidden" name="Referred_By_C" value="Op|String|0|200|||How did you hear about us"></td>
				</tr>
				<tr>
					
					<td width="1%" nowrap><font face="Arial" size="2"><B>I agree to the terms of service</b></font></td>
					<td width="98%">
					<input type="checkbox" name="Accept_Terms">
                                        <input type="hidden" name="hidResellerID" value="<%=intResellerID%>">
					<input type="hidden" name="Service" value="premium">
					</td>
				</tr>
				<tr>

				<td colspan=3 align=center><textarea cols=50 rows=6><%= Name %> Terms of Service
Last Updated: November 27, 2005

1. ACCEPTANCE OF TERMS 

Welcome to <%= Name %>. <%= Name %> provides its service to you, subject to the following Terms of Service ("TOS"), which may be updated by us from time to time without notice to you. 

You can review the most current version of the TOS at any time at: http://www.<%= host %>/tos.asp

In addition, when using particular <%= Name %> services, you and <%= Name %> shall be subject to any posted guidelines or rules applicable to such services, which may be posted from time to time. All such guidelines or rules are hereby incorporated by reference into the TOS.

2. DESCRIPTION OF SERVICE

<%= Name %> provides users with access to a collection of on-line resources, including, various communications tools, software, online forums, shopping services, personalized content and branded programming through its network of properties (the "Service"). Unless explicitly stated otherwise, any new features that augment or enhance the current Service, including the release of new <%= Name %> properties, shall be subject to the TOS. You understand and agree that the Service is provided "AS-IS" and that <%= Name %> assumes no responsibility for the timeliness, deletion, mis-delivery or failure to store any communications or data. 

In order to use the Service, you must obtain access to the World Wide Web, either directly or through devices that access web-based content, and pay any service fees associated with such access. In addition, you must provide all equipment necessary to make such connection to the World Wide Web, including a computer and modem or other access device. Please be aware that <%= Name %> has created certain areas that provide the ability to create adult or mature content. You must be at least 18 years of age to access and use such services. 

3. YOUR REGISTRATION OBLIGATIONS 

In consideration of your use of the Service, you agree to: (a) provide true, accurate, current and complete information about yourself and your company as prompted by the Service's registration form (such information being the "Registration Data") and (b) maintain and promptly update the Registration Data to keep it true, accurate, current and complete. If you provide any information that is untrue, inaccurate, not current or incomplete, or <%= Name %> has reasonable grounds to suspect that such information is untrue, inaccurate, not current or incomplete, <%= Name %> has the right to suspend or terminate your account and refuse any and all current or future use of the Service (or any portion thereof). <%= Name %> is not liable for any resulting loss of service due to untruthful or incomplete information. <%= Name %> is concerned about the safety and privacy of all its users, particularly minors. For this reason, parents, or the legal guardian of minors who wish to allow their children access to <%= Name %> must agree to supervise, and otherwise completely indemnify <%= Name %> of any content the minor may come into contact with during operation of the <%= Name %> service. Please remember that the Service is designed to appeal to a mature business audience. Accordingly, as the legal guardian, it is your responsibility to determine whether any of the Services and/or Content (as defined in Section 6 below) are appropriate for your child. 

4. <%= Name %> PRIVACY POLICY 

Registration Data and certain other information about you is subject to our Privacy Policy. For more information, please see our full privacy policy at http://www.<%= host %>/privacy.asp 

5. MEMBER ACCOUNT, PASSWORD AND SECURITY 

You will receive a password and account designation upon completing the Service's registration process. You are responsible for maintaining the confidentiality of the password and account, and are fully responsible for all activities that occur under your password or account. You agree to:

(a) immediately notify <%= Name %> of any unauthorized use of your password or account or any other breach of security. 
(b) ensure that you exit from your account at the end of each session.


<%= Name %> cannot and will not be liable for any loss or damage arising from your failure to comply with this Section 5. Furthermore <%= Name %> is not responsible for any damages, which may occur to you, your community standing, or your business in the event of a known or unknown person(s) gain unauthorized access to your account. 

6. MEMBER CONDUCT 

You understand that all information, data, text, software, music, sound, photographs, graphics, video, messages or other materials ("Content"), whether publicly posted or privately transmitted, are the sole responsibility of the person from which such Content originated. This means that you, and not <%= Name %>, are entirely responsible for all Content that you upload, post, email or otherwise transmit via the Service. <%= Name %> does not control the Content posted via the Service and, as such, does not guarantee the accuracy, integrity or quality of such Content. You understand that by using the Service, you may be exposed to Content that is offensive, indecent or objectionable. Under no circumstances will <%= Name %> be liable in any way for any Content, including, but not limited to, for any errors or omissions in any Content, or for any loss or damage of any kind incurred as a result of the use of any Content posted, emailed or otherwise transmitted via the Service. 

You agree to not use the Service to: 

A. Upload, post, email or otherwise transmit any Content that is unlawful, harmful, threatening, abusive, harassing, tortious, defamatory, vulgar, obscene, libelous, invasive of another's privacy, hateful, or racially, ethnically or otherwise objectionable
B. Harm minors in any way
C. Impersonate any person or entity,; including, but not limited to, a <%= Name %> official, forum leader, guide or host, or falsely state or otherwise misrepresent your affiliation with a person or entity
D. Forge headers or otherwise manipulate identifiers in order to disguise the origin of any Content transmitted through the Service
E. Upload, post, email or otherwise transmit any Content that you do not have a right to transmit under any law or under contractual or fiduciary relationships (such as inside information, proprietary and confidential information learned or disclosed as part of employment relationships or under nondisclosure agreements) 
F. Upload, post, email or otherwise transmit any Content that infringes any patent, trademark, trade secret, copyright or other proprietary rights ("Rights") of any party
G. Upload, post, email or otherwise transmit any unsolicited or unauthorized advertising, promotional materials, "junk mail," "spam," "chain letters," "pyramid schemes," or any other form of solicitation, except in those areas (such as shopping rooms) that are designated for such purpose
H. Upload, post, email or otherwise transmit any material that contains software viruses or any other computer code, files or programs designed to interrupt, destroy or limit the functionality of any computer software or hardware or telecommunications equipment; or that may interfere with or disrupt the normal flow and operation of any system. 
J. Interfere with or disrupt the Service or servers or networks connected to the Service, or disobey any requirements, procedures, policies or regulations of networks connected to the Service
K. Intentionally or unintentionally violate any applicable local, state, national or international law, including, but not limited to, regulations promulgated be the U.S. Securities and Exchange Commission, any rules of any national or other securities exchange, including, without limitation, the New York Stock Exchange, the American Stock Exchange or the NASDAQ, and any regulations having the force of law
L. "Stalk" or otherwise harass another individual
M. Collect or store personal data about other users of this service


You acknowledge that <%= Name %> does not pre-screen Content, but that <%= Name %> and its designees shall have the right (but not the obligation) in their sole discretion to refuse or move any Content that is available via the Service. Without limiting the foregoing, <%= Name %> and its designees shall have the right to remove any Content that violates the TOS or is otherwise objectionable. You agree that you must evaluate, and bear all risks associated with, the use of any Content, including any reliance on the accuracy, completeness, or usefulness of such Content. In this regard, you acknowledge that you may not rely on any Content created by <%= Name %> or submitted to <%= Name %>, including without limitation information in <%= Name %> and in all parts of the Service.

You acknowledge and agree that <%= Name %> may preserve Content and may also disclose Content if required to do so by law or in the good faith belief that such preservation or disclosure is reasonably necessary to: (a) comply with legal process; (b) enforce the TOS; (c) respond to claims that any Content violates the rights of third-parties; or (d) protect the rights, property, or personal safety of <%= Name %>, its users and the public.

You understand that the technical processing and transmission of the Service, including your Content, may involve (a) transmissions over various networks; and (b) changes to conform and adapt to technical requirements of connecting networks or devices.

7. SPECIAL ADMONITIONS FOR INTERNATIONAL USE 

Recognizing the global nature of the Internet, you agree to comply with all local rules in your jurisdiction regarding online conduct and acceptable Content. Specifically, you agree to comply with all applicable laws regarding the transmission of technical data exported from or to the United States or the country in which you reside. 

8. INDEMNITY 

You agree to indemnify and hold <%= Name %>, and its subsidiaries, affiliates, officers, agents, co-branders or other partners, and employees, harmless from any claim or demand, including reasonable attorneys' fees, made by any third party due to or arising out of Content you submit, post to or transmit through the Service, your use of the Service, your connection to the Service, your violation of the TOS, or your violation of any rights of another.

9. NO RESALE OF SERVICE 

You agree not to reproduce, duplicate, copy, sell, resell or exploit for any commercial purposes, any portion of the Service, use of the Service, or access to the Service. 

10. GENERAL PRACTICES REGARDING USE AND STORAGE 

You acknowledge that <%= Name %> may establish general practices and limits concerning use of the Service, including without limitation the maximum number of days that email messages, maximum sizes of images, total number of products, sales history, or other uploaded Content will be retained by the Service, the maximum number of email messages that may be sent from or received by an account on the Service, the maximum size of any email message that may be sent from or received by an account on the Service, the maximum disk space that will be allotted on <%= Name %>'s servers on your behalf, and the maximum number of times (and the maximum duration for which) you may access the Service in a given period of time. You agree that <%= Name %> has no responsibility or liability for the deletion or failure to store any communications or other Content maintained or transmitted by the Service. You acknowledge that <%= Name %> reserves the right to log off accounts that are inactive for an extended period of time. You further acknowledge that <%= Name %> reserves the right to change these general practices and limits at any time, in its sole discretion, with or without notice. 

11. MODIFICATIONS TO SERVICE 

<%= Name %> reserves the right at any time and from time to time to modify or discontinue, temporarily or permanently, the Service (or any part thereof) with or without notice. You agree that <%= Name %> shall not be liable to you or to any third party for any modification, suspension or discontinuance of the Service. 

12. TERMINATION 

You agree that <%= Name %>, in its sole discretion, may terminate your password, account (or any part thereof) or use of the Service, and remove and discard any Content within the Service, for any reason, including, without limitation, for lack of use or if <%= Name %> believes that you have violated or acted inconsistently with the letter or spirit of the TOS. <%= Name %> may also in its sole discretion and at any time discontinue providing the Service, or any part thereof, with or without notice. You agree that any termination of your access to the Service under any provision of this TOS may be effected without prior notice, and acknowledge and agree that <%= Name %> may immediately deactivate or delete your account and all related information and files in your account and/or bar any further access to such files or the Service. Further, you agree that <%= Name %> shall not be liable to you or any third-party for any termination of your access to the Service. 

13. DEALINGS WITH ADVERTISERS 

Your correspondence or business dealings with, or participation in promotions of, advertisers found on or through the Service, including payment and delivery of related goods or services, and any other terms, conditions, warranties or representations associated with such dealings, are solely between you and such advertiser. You agree that <%= Name %> shall not be responsible or liable for any loss or damage of any sort incurred as the result of any such dealings or as the result of the presence of such advertisers on the Service. 

14. LINKS 

The Service may provide, or third parties may provide, links to other World Wide Web sites or resources. Because <%= Name %> has no control over such sites and resources, you acknowledge and agree that <%= Name %> is not responsible for the availability of such external sites or resources, and does not endorse and is not responsible or liable for any Content, advertising, products, or other materials on or available from such sites or resources. You further acknowledge and agree that <%= Name %> shall not be responsible or liable, directly or indirectly, for any damage or loss caused or alleged to be caused by or in connection with use of or reliance on any such Content, goods or services available on or through any such site or resource. 

15. <%= Name %>'S PROPRIETARY RIGHTS 

You acknowledge and agree that the Service and any necessary software used in connection with the Service ("Software") may contain proprietary and confidential information that is protected by applicable intellectual property and other laws. You further acknowledge and agree that Content contained in sponsor advertisements or information presented to you through the Service or advertisers is protected by copyrights, trademarks, service marks, patents or other proprietary rights and laws. Except as expressly authorized by <%= Name %> or advertisers, you agree not to modify, rent, lease, loan, sell, distribute or create derivative works based on the Service or the Software, in whole or in part. 

<%= Name %> grants you a personal, non-transferable and non-exclusive right and license to use the object code of its Software on a single computer; provided that you do not (and do not allow any third party to) copy, modify, create a derivative work of, reverse engineer, reverse assemble or otherwise attempt to discover any source code, sell, assign, sublicense, grant a security interest in or otherwise transfer any right in the Software. You agree not to modify the Software in any manner or form, or to use modified versions of the Software, including (without limitation) for the purpose of obtaining unauthorized access to the Service. You agree not to access the Service by any means other than through the interface that is provided by <%= Name %> for use in accessing the Service, or by through interfaces which have been appropriately licensed and endorsed by <%= Name %>, Inc. 

16. DISCLAIMER OF WARRANTIES 

YOU EXPRESSLY UNDERSTAND AND AGREE THAT: 

A. YOUR USE OF THE SERVICE IS AT YOUR SOLE RISK. THE SERVICE IS PROVIDED ON AN "AS IS" AND "AS AVAILABLE" BASIS. <%= Name %> EXPRESSLY DISCLAIMS ALL WARRANTIES OF ANY KIND, WHETHER EXPRESS OR IMPLIED, INCLUDING, BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.

B. <%= Name %> MAKES NO WARRANTY THAT (i) THE SERVICE WILL MEET YOUR REQUIREMENTS, (ii) THE SERVICE WILL BE UNINTERRUPTED, TIMELY, SECURE, OR ERROR-FREE, (iii) THE RESULTS THAT MAY BE OBTAINED FROM THE USE OF THE SERVICE WILL BE ACCURATE OR RELIABLE, (iv) THE QUALITY OF ANY PRODUCTS, SERVICES, INFORMATION, OR OTHER MATERIAL PURCHASED OR OBTAINED BY YOU THROUGH THE SERVICE WILL MEET YOUR EXPECTATIONS, AND (V) ANY ERRORS IN THE SOFTWARE WILL BE CORRECTED.

C. ANY MATERIAL DOWNLOADED OR OTHERWISE OBTAINED THROUGH THE USE OF THE SERVICE IS DONE AT YOUR OWN DISCRETION A ND RISK AND THAT YOU WILL BE SOLELY RESPONSIBLE FOR ANY DAMAGE TO YOUR COMPUTER SYSTEM OR LOSS OF DATA THAT RESULTS FROM THE DOWNLOAD OF ANY SUCH MATERIAL.

D. NO ADVICE OR INFORMATION, WHETHER ORAL OR WRITTEN, OBTAINED BY YOU FROM <%= Name %> OR THROUGH OR FROM THE SERVICE SHALL CREATE ANY WARRANTY NOT EXPRESSLY STATED IN THE TOS.



17. LIMITATION OF LIABILITY 

YOU EXPRESSLY UNDERSTAND AND AGREE THAT <%= Name %> SHALL NOT BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL OR EXEMPLARY DAMAGES, INCLUDING BUT NOT LIMITED TO, DAMAGES FOR LOSS OF PROFITS, GOODWILL, USE, DATA OR OTHER INTANGIBLE LOSSES (EVEN IF <%= Name %> HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES), RESULTING FROM: (i) THE USE OR THE INABILITY TO USE THE SERVICE; (ii) THE COST OF PROCUREMENT OF SUBSTITUTE GOODS AND SERVICES RESULTING FROM ANY GOODS, DATA, INFORMATION OR SERVICES PURCHASED OR OBTAINED OR MESSAGES RECEIVED OR TRANSACTIONS ENTERED INTO THROUGH OR FROM THE SERVICE; (iii) UNAUTHORIZED ACCESS TO OR ALTERATION OF YOUR TRANSMISSIONS OR DATA; (iv) STATEMENTS OR CONDUCT OF ANY THIRD PARTY ON THE SERVICE; OR (v) ANY OTHER MATTER RELATING TO THE SERVICE. 

18. INDEMNIFICATION OF ILLEGAL ACTIONS 

By using <%= Name %> services you represent that the goods or services you promote on <%= Name %> belong to you, and that such items were obtained legally, through sanctioned reputable channels, and that such goods or services are legal in your jurisdiction, and that you have the authorization by all necessary government or regulatory boards to sell, issue, authorize, or otherwise transfer said goods and services. Furthermore you represent that you will take all necessary measures to insure said goods or services do not reach areas, or individuals who are not allowed to own, handle, or otherwise poses such goods or services. Furthermore you represent you will take reasonable steps to insure that said goods or services do not transit an area where said goods and services are illegal or prohibited, unless explicitly allowed through other applicable laws such as in the case of interstate commerce. And that in the case of dangerous goods, all necessary preparations were made to avoid any accidental release, leakage, discharge, or other action which may expose persons responsible for shipping, transporting, or receiving goods any harm. You further agree to indemnify <%= Name %>, and its subsidiaries, affiliates, officers, agents, co-branders or other partners, and employees, harmless from any claim or demand, including reasonable attorneys' fees, made by any third party due to or arising out of your sale and transportation of goods and or services.

19. NOTICE 

Notices to you may be made via either email or regular mail. The Service may also provide notices of changes to the TOS or other matters by displaying notices or links to notices to you generally on the Service.

20. TRADEMARK INFORMATION 

<%= Name %>, and <%= host %> are trademarks of <%= Name %>. all rights reserved. 

21. COPYRIGHTS and COPYRIGHT AGENTS 

<%= Name %> respects the intellectual property of others, and we ask our users to do the same. If you believe that your work has been copied in a way that constitutes copyright infringement, please provide <%= Name %>'s Copyright Agent the following information: 

An electronic or physical signature of the person authorized to act on behalf of the owner of the copyright interest; 

a description of the copyrighted work that you claim has been infringed; 

a description of where the material that you claim is infringing is located on the site; 

your address, telephone number, and email address; 

a statement by you that you have a good faith belief that the disputed use is not authorized by the copyright owner, its agent, or the law; 

a statement by you, made under penalty of perjury, that the above information in your Notice is accurate and that you are the copyright owner or authorized to act on the copyright owner's behalf. 

22. SEARCH ENGINE POLICY


All free <%= Name %> sites are hosted under the "<%= host %>" domain name. Many search engines do not distinquish that each site is separate and therefore the actions of one customer, can affect all customers. For this reason <%= Name %> has developed a strict policy regarding Search Engine "Spamming" techniques. Customers are explicitly prohibited from:

Cloaking - creating deceptive pages which attempt to trick search engines into giving higher ranking.

Writing text or creating links that can be seen by search engines but not by visitors.

Participating in link exchanges for the sole purpose of increasing your ranking in search engines. 

Sending automated queries to Google in an attempt to monitor your site's ranking.

Use programs that generate lots of generic doorway pages. 



22. GENERAL INFORMATION

The TOS constitute the entire agreement between you and <%= Name %> and govern your use of the Service, superceding any prior agreements between you and <%= Name %>. You also may be subject to additional terms and conditions that may apply when you use affiliate services, third-party content or third-party software. The TOS and the relationship between you and <%= Name %> shall be governed by the laws of the State of California without regard to its conflict of law provisions. You and <%= Name %> agree to submit to the personal and exclusive jurisdiction of the courts located within the county of San Diego, California. The failure of <%= Name %> to exercise or enforce any right or provision of the TOS shall not constitute a waiver of such right or provision. If any provision of the TOS is found by a court of competent jurisdiction to be invalid, the parties nevertheless agree that the court should endeavor to give effect to the parties' intentions as reflected in the provision, and the other provisions of the TOS remain in full force and effect. You agree that regardless of any statute or law to the contrary, any claim or cause of action arising out of or related to use of the Service or the TOS must be filed within one (1) year after such claim or cause of action arose or be forever barred. 

The section titles in the TOS are for convenience only and have no legal or contractual effect. 


24. CANCELLATION OF SERVICE

Service may be cancelled online at anytime and all recurring billings will be immediately terminated.  If cancellation of service is requested within 45 days of initial activation customer may request a full refund of payment.  After the intial 45 day testing period all payments are final.  Discounts are given for paying for multiple months at a time.  If a customer pays for multiple months and then cancels before using up the aforementioned months no credit will be given for unused months unless the cancellation occurs within the first 45 days of service.


25. VIOLATIONS

Please report any violations of the TOS to our Customer Care group at support@<%= host %>.

</textarea></td></tr>

			<tr>
					

					<td colspan=2 align=center>

						<input type="submit" border="0" value="Create Site Now"></td>
				</tr>
				</table>

			</form>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Service_Level","req","Please choose a service level.");
 frmvalidator.addValidation("Store_name","req","Please enter a store name.");
 frmvalidator.addValidation("Site_name","req","Please enter a site name.");
 frmvalidator.addValidation("Site_name","alnumhyphen","Your site name cannot contain special characters.");
 frmvalidator.addValidation("Store_User_id","req","Please enter a login.");
 frmvalidator.addValidation("Store_password","req","Please enter your password.");
 frmvalidator.addValidation("Store_password","minlength=4","For security purposes your password must be at least 5 characters long.");
 frmvalidator.addValidation("Store_password_confirm","req","Please enter your password again.");
 frmvalidator.addValidation("Email","req","Please enter a valid email.");
 frmvalidator.addValidation("Email","email","Please enter a valid email.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");

function DoCustomValidation()
{
  var frm = document.forms[0];
  if (document.forms[0].Service_Level.value == "")
    {
    alert('Please choose a service level.');
          return false;
    }


    if (document.forms[0].Accept_Terms.checked == true)
  {
    if (document.forms[0].Store_password.value == document.forms[0].Store_password_confirm.value)
    {
      if (document.forms[0].Accept_Terms.checked == true)
      {
        if (document.forms[0].Site_name.value == document.forms[0].Store_password.value)
        {
          alert('Your password cannot be the same as your login or site name.');
          return false;
        }
        else
        {
          if (document.forms[0].Store_User_id.value == document.forms[0].Store_password.value)
          {
            alert('Your password cannot be the same as your login or site name.');
            return false;
          } 
          else
          {
          return true;
          }
        }
      }
    }
    else
    {
      alert('The 2 passwords must match.');
      return false;
    }
  }
  else
  {
    alert('You must accept the terms of service in order to proceed.');
    return false;

  }
}


</script>

<!--#include file="footer.asp"-->
