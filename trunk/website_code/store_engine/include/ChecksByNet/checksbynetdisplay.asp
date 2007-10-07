<%

sub create_form_post_checksNet ()

    sql_real_time = "exec wsp_real_time_property "&Store_Id&",33;"
    rs_Store.open sql_real_time,conn_store,1,1
    rs_Store.MoveFirst
    Do While Not Rs_Store.EOF
        select case Rs_store("Property")
            case "PaytoID"
                PaytoID = decrypt(Rs_store("Value"))
        end select
     Rs_store.MoveNext
    Loop
    Rs_store.Close


 ' TEST TRANSACTION

'	 response.write "<form method='POST' action='https://cross.checksbynet.com/response.asp' name='Payment' onSubmit=""return checkFields();""><INPUT type='hidden' name='paytoid' value='"&PaytoID&"'><INPUT type='hidden' name='payto' value='Valued CrossCheck Merchant'>"
	 response.write "<form method='POST' action='https://cross.checksbynet.com/response.asp' name='Payment' onSubmit=""return checkFields();""><INPUT type='hidden' name='paytoid' value='"&PaytoID&"'>"
end sub

sub create_form_content_checksNet ()

	Secure_Name = Replace(Site_Name,"http://","https://")
	Return_Addr = Secure_Name&"include/checksByNet/checksbynetResponse.asp"
	Return_Addr = Return_Addr & "?OrderID="&oid&"&Shopper_id="&Shopper_id
	%>

<!-- start checks by net -->
</tr>
<tr><td>

  <input type="hidden" name="ScriptURL" value="<%= Return_Addr %>">
	<input type="hidden" name="checkamt" value="<%= formatnumber(GGrand_Total,2) %>">

  <div align="center">
		<center>
			<p>
				<img src="include/checksbynet/check.gif" border="0" width="390" height="222">
			</p>
		</center>
	</div>
	
	<div align="left">

		<table border="0" cellpadding="0" width="103%" height="181">


	 <tr>
      <td valign="top" height="25"><font face="Verdana, Trebuchet MS, Geneva" size="1"><strong> 
        <input size="13" name="checknbr" value="">
        </strong></font><font face="Verdana, Trebuchet MS, Geneva" size="1"> <strong> 
        </strong> </font> </td>
      <td align="right" valign="top" height="30">&nbsp;</td>
      <td valign="top" height="25">
				<font face="Verdana, Trebuchet MS, Geneva" size="2">
					Enter the next <strong>check</strong> <strong>number</strong> in your check book.
				</font>
			</td>
    </tr>
	
    <tr>
      <td valign="top" height="34">
				<font face="Verdana, Trebuchet MS, Geneva" size="1">
					<input type="text" size="19" name="idnbr" value="">
					<input type="text" size="3" name="idst" value="">
				</font>
			</td>
      <td align="right" valign="top" height="34">
				<font face="Verdana, Trebuchet MS, Geneva" size="1">&nbsp;</font>
			</td>
      <td valign="top" height="34">
				<font face="Verdana, Trebuchet MS, Geneva" size="2">
				 Enter your <strong>drivers</strong> <strong>license</strong> <strong>number</strong>
         and its 2-letter <strong>state</strong> code.<br>
         Do <u>not</u> include dashes or spaces.
				 </font>
				</td>
    </tr>

    <tr>
      <td valign="top" height="27">
				<font face="Verdana, Trebuchet MS, Geneva" size="1">
					<input type="text" size="25" name="bankname">
				</font>
			</td>
      <td align="right" valign="top" height="27">&nbsp;</td>
      <td valign="top" height="27">
				<font face="Verdana, Trebuchet MS, Geneva" size="2">
					&nbsp;Enter your <strong>bank's</strong> <strong>name</strong>.
				</font>
			</td>
    </tr>

		<tr>
      <td valign="top" height="22">
				<font face="Verdana, Trebuchet MS, Geneva" size="1">
					<strong>
					  <input type="text" size="12" name="bankcity">
						<input type="text" size="3" name="bankst" value="">
						<input type="text" size="6" name="bankzip">
					</strong>
				</font>
			</td>
      <td align="right" valign="top" height="22">&nbsp;</td>
      <td valign="top" height="22">
				<font face="Verdana, Trebuchet MS, Geneva" size="2">
				 Enter your <strong>bank's</strong>
				 <strong>city</strong>, 2-letter
				 <strong>state</strong> code, and
				 <strong>zip</strong>
         <strong>code</strong>.</font></td>
    </tr>

		 <tr>
      <td valign="top" height="47">
				<font face="Verdana, Trebuchet MS, Geneva" size="1">
					<strong>
						<input size="32" name="Micr">
					</strong>
				</font>
			</td>
      <td align="right" valign="top" height="47">&nbsp;</td>
      <td valign="top" height="47">
				<font FACE="Verdana, Trebuchet MS, Helvetica, Arial" SIZE="2">Enter 
				<b>ALL NUMBERS </b>
				    from the
            bottom of your check starting from left to right.<br>
            For each symbol encountered, enter<b> &quot;
						<font color="#00AA00">S</font>&quot;<br>
        </b>
        </font>
				<font face="Verdana, Trebuchet MS, Geneva, Arial" size="1">
				   BE SURE TO INCLUDE SPACES (Estimation on the number of spaces is
           allowed)
				</font>
			</td>
    </tr>

</table>
</div>



<table border="0" cellpadding="0" cellspacing="0" width="80%" height="179">

		<tr>
      <td align="center" width="334" height="62">
				<font size="2" face="Verdana, Trebuchet MS, Geneva, Arial">
					<strong>
						<div align="left"><p>

								Bank routing number and checking account number may vary. Below are two
                common examples of placement.
							
					</strong>
				</font>
			</td>
    </tr>

    <tr>
      <td align="center" width="682" height="22">
				<div align="left"><p>
					<font FACE="Verdana,Trebuchet MS,Helvetica,Arial" SIZE="2">
						<b>Example 1:</b>
					</font>
        </div>
        <div align="left"><p>
					<font FACE="Verdana,Trebuchet MS,Helvetica,Arial" SIZE="2">
						<b>Enter as-&nbsp;
						 <font color="#FF0000">S</font>987654321
						 <font color="#FF0000">S</font>67895432
						 <font color="#FF0000">S</font>  00250
					 </b>
					</font>
        </div>
      </td>
    </tr>

    <tr>
      <td align="center" width="682" height="22">
				<font FACE="Verdana,Trebuchet MS,Helvetica,Arial" SIZE="2">
					<b>
						<div align="left">
							<p>
								<img src="include/checksbynet/example1.gif" alt="example1.gif (4651 bytes)" width="302" height="26">
					</b>
				</font>
        </div>
      </td>
    </tr>

    <tr align="center">
      <td align="left" font size="2" face="Verdana, Trebuchet MS, Geneva, Arial" height="37" width="682" valign="bottom">
				<font FACE="Verdana,Trebuchet MS,Helvetica,Arial" SIZE="2">
					<b>Example 2:</b>
				</font>
        <p>
					<font FACE="Verdana,Trebuchet MS,Helvetica,Arial" SIZE="2">
						<b>Enter as-&nbsp;
							<font color="#FF0000">S</font>987654321
							<font color="#FF0000">S</font>00250        6789
							<font color="#FF0000">S</font>5432
							<font color="#FF0000">S</font>
						</b>
					</font>
				</p>
      </td>
    </tr>

    <tr align="center">
      <td align="center" font size="2" face="Verdana, Trebuchet MS, Geneva, Arial" width="639" height="30">
				<img src="include/checksbynet/example2.gif" alt="example2.gif (4726 bytes)" align="left" WIDTH="288" HEIGHT="29">
					<p>&nbsp;
				</td>
    </tr>

 </table>

  <p><br>
  <font face="Verdana, Trebuchet MS, Geneva">
	</font>
	</p>
  
	<table border="0">
    <tr>
      <td><div align="right"><p>
				<font face="Verdana, Trebuchet MS, Geneva" size="2">
					<strong>Please
      read and approve the following authorization:</strong></font></td>
      <td valign="top" width="75%"><font face="Verdana, Trebuchet MS, Geneva" size="1">I
      authorize ChecksByNet to duplicate the preceeding information into a bank draft form. I
      understand that I will receive by email, a check authorization notice, notifying me that a
      bank draft has been issued on my behalf for said purchase. I will retain my original check
      for my record of the transaction.</font></td>
    </tr>
    <tr>
      <td></td>
      <td width="75%"><font face="Verdana, Trebuchet MS, Geneva" size="1">I understand that the
      Payee or authorized agent of Payee, will sign the bank draft as my agent for this
      transaction only. This authorization is valid for this transaction only. No other bank
      drafts will be created without my direct written or verbal authorization. All returned
      checks are subject to a fee of $22.50 or the maximum allowed by law plus returned bank
      debit fee.</font></td>
    </tr>
    <tr><td colspan=2>If preferred, you can include a memo to print on the check (30 chars max):
	<input type="text" name="Memo" size="30" maxlength="30"></td></tr>
  </table>






	<% rs_store.open "select * from store_customers where record_type=0 and cid="&cid, conn_store, 1, 1 %>
	<% If not rs_Store.eof then %>
		<INPUT type="hidden" NAME="writerfirst" value="<%= rs_store("First_name") %>">
		<INPUT type="hidden" NAME="writerlast" value="<%= rs_store("Last_name") %>">
		<INPUT type="hidden" NAME="writeraddr" value="<%= rs_store("Address1") %> <%= rs_store("Address2") %> ">
		<INPUT type="hidden" NAME="writercity" value="<%= rs_store("City") %>">
		<INPUT type="hidden" NAME="writerst" value="<%= rs_store("State") %>">
		<INPUT type="hidden" NAME="writerzip" value="<%= rs_store("Zip") %>">
		<INPUT type="hidden" NAME="email" value="<%= rs_store("Email") %>">
		<INPUT type="hidden" NAME="phone" value="<%= rs_store("Phone") %>">
	<% End If %>
	<% rs_store.close %>
	</td></tr>
<!-- end checks by net -->
<% end sub %>
