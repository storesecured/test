<%
title = "eCommerce Tutorial eCommerce Tutorials On eCommerce"
description = "eCommerce tutorial helps clients realize maximum profits. Discover effective ecommerce tutorials on ecommerce fundamental principle applications. Take our free ecommerce tutorial today."
keyword1="ecommerce tutorial"
keyword2="ecommerce tutorials"
keyword3="tutorials on ecommerce"
keyword4=""
keyword5=""

include_extra_links = true
include_credit = true
tracking_page_name="ecommerce-tutorial"
quicktext="Ecommerce Tutorial"
%>
<!--#include file="header.asp"-->
		<% if request.form("ecommerce") <> "" and request.form("Email_Address") <> "" then

	sEmail = request.form("Email_Address")
	if instr(sEmail,"@") > 0 and instr(sEmail," ") = 0 then
	iDay = 0
	sAttach_Folder = fn_get_sites_folder(Store_Id,"Attachments")
	 For iDay = 0 to 9
	     Set Mail = Server.CreateObject("Persits.MailSender")
	   	Mail.From = sNoReply_email
		Mail.AddAddress sEmail
		Mail.Subject = "Ecommerce Tutorial Day "&iDay
		Mail.AppendBodyFromFile sAttach_Folder & "day"&iDay&".txt"
		Mail.Timestamp = Now() + iDay
		Mail.Queue = True
		Mail.Send
		set Mail = Nothing
	  next
	  
	  iDay=15
	  Set Mail = Server.CreateObject("Persits.MailSender")
	  Mail.From = sNoReply_email
	  Mail.AddAddress sEmail
	  Mail.Subject = "Ecommerce Tutorial Day "&iDay
	 Mail.AppendBodyFromFile sAttach_Folder & "day"&iDay&".txt"
	 Mail.Timestamp = Now() + iDay
	 Mail.Queue = True
	 Mail.Send

	  response.write "<BR>Congratulations, you are now receiving our ecommerce tutorial, please check your email for your first installment of our Ecommerce Tutorial."
    else
        response.write "You have entered an invalid email address.  Please use your browsers back button and try again."    end if
else %>
		<h2>Free eCommerce Tutorial</h2>
		  How To Create a Successful eCommerce Website in 10 Easy Steps
		<p>Going online with your first ebusiness is easy with our EasyStoreCreator 
		  system. </p>
		<p>It's <b>even easier</b> with our FREE ten-day, ten-lesson ecommerce tutorial. 
		</p>
		<p>Most ecommerce tutorials rush you through three or four superficial lessons 
		  that just touch on the basics. Our self-paced ecommerce tutorial provides 
		  you with everything you need to know to start and succeed in your new 
		  online business. </p>
		<p>Our tutorials on ecommerce will arrive in your inbox every day for ten 
		  days. You can study them as they arrive or compile them for future reference. 
		  They're so packed with useful information that you'll refer to them often 
		  -- even after your online store is up and running. </p>
		<p>The ecommerce tutorials you receive will give you valuable tips on subjects 
		  like building a home page that works, selecting a domain name, marketing
		  your online store, accepting credit cards, and much more. </p>
		<p>Here is an outline of our ten-day ecommerce tutorial:</p>
		<ul>
		  <li>Day 1: Sixteen Reasons to Put Your Business On The Web <br>
		  </li>
		  <li>Day 2: Choosing a Vendor to Process Your Transactions <br>
		  </li>
		  <li>Day 3: Is Ecommerce Right for Your Small Business? <br>
		  </li>
		  <li>Day 4: Level the Playing Field with eCommerce <br>
		  </li>
		  <li>Day 5: Building a Successful Home Page <br>
		  </li>
		  <li>Day 6: Tools Needed to Sell Online <br>
		  </li>
		  <li>Day 7: Accepting Credit Cards <br>
		  </li>
		  <li>Day 8: Choosing a Domain Name <br>
		  </li>
		  <li>Day 9: Twenty-nine Ways to Promote Your Website <br>
		  </li>
		  <li>Day 10: Conclusion </li>
		</ul>
		<p>Sign up now by entering your email address in the box below. The first 
		  of your ecommerce tutorials will arrive shortly and you'll receive one 
		  lesson a day for ten days. We hope you enjoy our EasyStoreCreator free 
		  tutorials on ecommerce!</p>
		<form method=post action=ecommerce-tutorial.asp?Signup=Yes name=ecommerce>
		  <input type=text name=Email_Address size=22>
		  <input type=submit name=ecommerce value=Signup>
		</form>
		<p>
		Questions? Please <a href="contactus.asp">click
		here to contact us</a>.</p> 
		<% end if %>
		<!--#include file="footer.asp"-->

