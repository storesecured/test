<% sub put_paypal () %>
<input type="hidden" name="cancel_return" value="<%= Cancel_Addr %>">
<input type="hidden" name="return" value="<%= Return_Addr %>">
<input type="hidden" name="item_number" value="<%= Store_Id %>">
<input type="hidden" name="no_note" value="1">
<input type="hidden" name="no_shipping" value="1">

<input type="hidden" name="custom" value="<%= Store_id %>">
<INPUT type="submit" NAME="submit" value="Pay Now">
</td></tr></form>
<%
end sub

sub put_normal () %>

<tr bgcolor='#FFFFFF'><td colspan=3>The address information you provide below will be used to verify your transaction
and prevent unauthorized charges.</td></tr>
   <input type=hidden name=payment_method value='<%= payment_method %>'>
       <tr bgcolor='#FFFFFF'>
       <td width="24%" height="23" class="inputname"><B>Company Name</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="company" value="<%= company %>" size="60" maxlength=250>
       <input type="hidden" name="company_C" value="Op|String|0|250|||Company Name">
       <% small_help "Company Name" %></td>
       </tr>
       <tr bgcolor='#FFFFFF'>
       <td width="24%" height="23" class="inputname"><B>First Name</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="first" value="<%= first_name %>" size="60" maxlength=100>
       <input type="hidden" name="first_C" value="Re|String|0|100|||First">
       <% small_help "First" %></td>
       </tr>
       <tr bgcolor='#FFFFFF'>
       <td width="24%" height="23" class="inputname"><B>Last Name</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="last" value="<%= last_name %>" size="60" maxlength=100>
       <input type="hidden" name="last_C" value="Re|String|0|100|||Last">
       <% small_help "Last" %></td>
       </tr>
       <tr bgcolor='#FFFFFF'>
       <td width="24%" height="23" class="inputname"><B>Address</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="address" value="<%= address %>" size="60" maxlength=250>
       <input type="hidden" name="address_C" value="Re|String|0|250|||Address">
       <% small_help "Address" %></td>
       </tr>


       <tr bgcolor='#FFFFFF'>
       <td width="24%" height="23" class="inputname"><B>City</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="city" value="<%= city %>" size="60" maxlength=100>
       <input type="hidden" name="city_C" value="Re|String|0|100|||City">
       <% small_help "City" %></td>
       </tr>
       <tr bgcolor='#FFFFFF'>
       <td width="24%" height="23" class="inputname"><B>State</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="State" value="<%= state %>" size="60" maxlength=100>
       <input type="hidden" name="State_C" value="Re|String|0|100|||State">
       <% small_help "State" %></td>
       </tr>
       <tr bgcolor='#FFFFFF'>
          <td width="24%" height="23" class="inputname"><B>Country</B></td>
       <td width="76%" height="23" class="inputvalue">
             <select size="1" name="country">
             <%
            sql_region = "SELECT Country FROM Sys_Countries where country <> 'All Countries' ORDER BY Country;"
            set myfields=server.createobject("scripting.dictionary")
            Call DataGetrows(conn_store,sql_region,mydata,myfields,noRecords)

            if noRecords = 0 then
            FOR rowcounter= 0 TO myfields("rowcount")

                ' set the selected flag
                if country=mydata(myfields("country"),rowcounter) then
                   selected = "selected"
                else
                   selected = ""
                end if
                response.write "<Option "&selected&" value='"&mydata(myfields("country"),rowcounter)&"'>"&mydata(myfields("country"),rowcounter)&"</option>"
             Next
             end if
             set myfields= Nothing %>
             </select>
             <% small_help "Country" %></td>
       </tr>



       <tr bgcolor='#FFFFFF'>
    <td width="24%" height="23" class="inputname"><B>Zip Code</B></td>
    <td width="76%" height="23" class="inputvalue">
             <input type="text" name="zip" value="<%= zip %>" size="5" maxlength=10>
             <INPUT type="hidden"  name=zip_C value="Re|String|0|20|||Zip Code">
             <% small_help "Zip Code" %></td>
       </tr>

       <tr bgcolor='#FFFFFF'>
    <td width="24%" height="23" class="inputname"><B>Phone</B></td>
       <td width="76%" height="23" class="inputvalue">
             <input type="text" name="Phone" value="<%= Phone %>" size="60" maxlength="14">
       <INPUT type="hidden"  name=Phone_C value="Re|String|0|14|||Phone">
       <% small_help "Phone" %></td>
          </tr>

       <tr bgcolor='#FFFFFF'>
    <td width="24%" height="23" class="inputname"><B>Fax</B></td>
    <td width="76%" height="23" class="inputvalue">
             <input type="text" name="Fax" value="<%= Fax %>" size="60" maxlength="14">
       <INPUT type="hidden"  name=Fax_C value="Op|String|0|14|||Fax">
       <% small_help "Fax" %></td>
       </tr>

       <tr bgcolor='#FFFFFF'>
    <td width="24%" height="17" class="inputname"><B>Email</B></td>
    <td width="76%" height="17" class="inputvalue">
             <input type="text" name="email" value="<%= email %>" size="60" maxlength=50>
       <INPUT type="hidden"  name=email_C value="Re|String|0|50|@,.||Email">
       <% small_help "Email" %></td>
       </tr>

<%
if payment_method = "eCheck" then
%>
<tr bgcolor='#FFFFFF'>
               <TD width="176" class="inputname"><B>Bank Name</B></TD>
               <TD width="219" class="inputvalue">
               <INPUT type="text" name="BankName" size="60" value='<%= Bank_Name %>' maxlength=50>
               <INPUT name=BankName_C type=hidden value="Re|String|0|50|||Bank Name"><% small_help "Bank Name" %></TD>
            </TR>

            <tr bgcolor='#FFFFFF'>
               <TD width="176" class="inputname"><B>Routing #</B></TD>
               <TD width="219" class="inputvalue">
               <table><tr bgcolor='#FFFFFF'><td><img src=images/symbol_route.gif></td><td><INPUT type="text" name="BankABA" size="60" maxlength=9 value='<%= Bank_ABA %>' onKeyPress="return goodchars(event,'0123456789')"></td><td><img src=images/symbol_route.gif>
               <INPUT name=BankABA_C type=hidden value="Re|String|0|50|||Bank Routing Number"></td></tr></table><% small_help "Bank Routing" %></TD>
            </TR>
            <tr bgcolor='#FFFFFF'>
               <TD width="176" class="inputname"><B>Account #</b></TD>
               <TD width="219" class="inputvalue" nowrap><table><tr bgcolor='#FFFFFF'><td><INPUT type="text" name="BankAccount" size="60" value='<%= Bank_Account %>' onKeyPress="return goodchars(event,'0123456789')"></td><td><img src=images/symbol_account.gif>
               <INPUT name=BankAccount_C type=hidden value="Re|String|0|252|||Account Num"></td></tr></table><% small_help "Bank Account" %></TD>
            </TR>
            <tr bgcolor='#FFFFFF'>
               <TD width="176"  class="inputname"><B>Account Type</b></TD>
               <TD width="219" class="inputvalue">
                  <select name="acct_type" size="1">
                  <option
                  <% if Acct_Type = "CHECKING" then %>
                     selected
                  <% end if %>
                  value="CHECKING">Checking</option>
                  
                  </select>
                  <select name="org_type" size="1">
                  <option
                  <% if Org_Type = "I" then %>
                     selected
                  <% end if %>
                  value="I">Individual</option>
                  <option
                  <% if Org_Type = "B" then %>
                     selected
                  <% end if %>
                  value="B">Business</option>
                  </select>
               <% small_help "Account Type" %></TD>
            </TR>
            <tr bgcolor='#FFFFFF'>
            <TD width="176" class="inputname"><B>Check #</b></TD>
               <TD width="219" class="inputvalue"><INPUT type="text" name="CheckSerial" size="5" maxlength=6 value='<%= Check_Num %>' onKeyPress="return goodchars(event,'0123456789')"><a href="JavaScript:goCheck();"><img src=images/echeck.jpg border=0></a>
               <INPUT name=CheckSerial_C type=hidden value="Re|String|0|6|||Check Serial #"><% small_help "Check Num" %></TD>
            </TR>


            </TR>
            <tr bgcolor='#FFFFFF'>
               <TD width="176" class="inputname"><B>Drivers Lic #</b></TD>
               <TD width="219" class="inputvalue"><INPUT type="text" name="DrvNumber" size="15" value='<%= License_Num %>' maxlength=50>
               <INPUT name=DrvNumber_C type=hidden value="Re|String|0|50|||License Number">

               <select name=DrvState>
               <option  value='<%= License_State %>'><%= License_State %></option>
               <option  value='AK'>AK</option><option  value='AL'>AL</option><option  value='AR'>AR</option><option  value='AZ'>AZ</option><option  value='CA'>CA</option><option  value='CO'>CO</option><option  value='CT'>CT</option><option  value='DC'>DC</option><option  value='DE'>DE</option><option  value='FL'>FL</option><option  value='GA'>GA</option><option  value='HI'>HI</option><option  value='IA'>IA</option><option  value='ID'>ID</option><option  value='IL'>IL</option><option  value='IN'>IN</option><option  value='KS'>KS</option><option  value='KY'>KY</option><option  value='LA'>LA</option><option  value='MA'>MA</option><option  value='MD'>MD</option><option  value='ME'>ME</option><option  value='MI'>MI</option><option  value='MN'>MN</option><option  value='MO'>MO</option><option  value='MS'>MS</option><option  value='MT'>MT</option><option  value='NC'>NC</option><option  value='ND'>ND</option><option  value='NE'>NE</option><option  value='NH'>NH</option><option  value='NJ'>NJ</option><option  value='NM'>NM</option><option  value='NT'>NT</option><option  value='NV'>NV</option><option  value='NY'>NY</option><option  value='OH'>OH</option><option  value='OK'>OK</option><option  value='OR'>OR</option><option  value='PA'>PA</option><option  value='PR'>PR</option><option  value='RI'>RI</option><option  value='SC'>SC</option><option  value='SD'>SD</option><option  value='TN'>TN</option><option  value='TX'>TX</option><option  value='UT'>UT</option><option  value='VA'>VA</option><option  value='VT'>VT</option><option  value='WA'>WA</option><option  value='WI'>WI</option><option  value='WV'>WV</option><option  value='WY'>WY</option>
               </select>
               <INPUT name=DrvState_C type=hidden value="Re|String|0|2|||Lic State">

               <% small_help "License Num" %></TD>

            </TR>
            <% Current_year = year(now())
               iYearStart = Current_year
               iYearEnd = Current_year + 20 %>
            <tr bgcolor='#FFFFFF'>
               <TD width="176" class="inputname"><B>Drivers Lic Exp</b></TD>
               <TD width="219" class="inputvalue">
                  <select name="dobm" size="1">
                  <option  value='<%= License_Exp_Month %>'><%= License_Exp_Month %></option>
                  <option value="01">January</option>
                  <option value="02">February</option>
                  <option value="03">March</option>
                  <option value="04">April</option>
                  <option value="05">May</option>
                  <option value="06">June</option>
                  <option value="07">July</option>
                  <option value="08">August</option>
                  <option value="09">September</option>
                  <option value="10">October</option>
                  <option value="11">November</option>
                  <option value="12">December</option>
                  </select>
                  <INPUT name=dobm_C type=hidden value="Re|integer|1|12|||Exp Month">

            <INPUT type="text" name="dobd" size="2" maxlength=2 value='<%= License_Exp_Day %>' onKeyPress="return goodchars(event,'0123456789')">
                  <INPUT name=dobd_C type=hidden value="Re|integer|1|31|||Exp Day">
                  <%  %>
                  <select name="doby" size="1">
                  <option  value='<%= License_Exp_Year %>'><%= License_Exp_Year %></option>

                  <% for go_year =  iYearStart to iYearEnd %>
                        <option value="<%= go_year %>"><%= go_year %>
                  <% next %>

               </select>
               <INPUT name=doby_C type=hidden value="Re|integer|2004|2030|||Exp Year">

               <% small_help "License Exp" %></TD>
            </TR>
<%
else
%>



    <tr bgcolor='#FFFFFF'><TD class="inputname"><B>Credit Card Number</B></TD><TD class="inputvalue">
    <Input name=current_cc value="<%= Card_Number %>" type=hidden>
    <INPUT NAME="cc_num" SIZE=60 value="" maxlength=20 onKeyPress="return goodchars(event,'0123456789')" >
    <% if payment_method = "Visa" then %>
      <img src=images/icon_visa.gif>
    <% elseif payment_method ="Mastercard"then %>
      <img src=images/icon_mastercard.gif>
    <% elseif payment_method = "Discover" then %>
      <img src=images/icon_discover.gif>
   <% elseif payment_method = "American Express" then %>
      <img src=images/icon_amex.gif>
   <% end if %><BR>
   <% if Card_Number <> "" then %>
   <font size=1>Leave empty to charge card on file ending in <%= right(decrypt(Card_Number),4) %></font>
   <% end if %>
    <% small_help "CC Number" %></TD>
    </TR>


    <tr bgcolor='#FFFFFF'>
    <TD class="inputname"><B>Expiration Date</B></TD><TD class="inputvalue"><select name="mm" size="1">
       <option selected value="<%= Exp_Month %>"><%= Exp_Month %></option>
       <option value="01">January</option>
       <option value="02">February</option>
       <option value="03">March</option>
       <option value="04">April</option>
       <option value="05">May</option>
       <option value="06">June</option>
       <option value="07">July</option>
       <option value="08">August</option>
       <option value="09">September</option>
       <option value="10">October</option>
       <option value="11">November</option>
       <option value="12">December</option>
       </select>
       <input type=hidden name=mm_C value="Re|integer|1|12|||Exp Month">
       <% Current_year = year(now()) %>
       <select name="yy" size="1">
       <option selected value="<%= Exp_Year %>"><%= Exp_Year %></option>
       <% for go_year =     Current_year to Current_year + 10 %>
             <% go_year_two_digit = right(go_year,2) %>
             <option value="<%= go_year_two_digit %>"><%= go_year %>
          <% next %>
    </select>
    <input type=hidden name=yy_C value="Re|integer|2004|2024|||Exp Year">

    <% small_help "Expiration Date" %></TD>
    </TR>
    <tr bgcolor='#FFFFFF'>
    <TD class="inputname"><B>Card Code</B></TD><TD class="inputvalue"><INPUT NAME="CardCode" SIZE=4 maxlength=4 onKeyPress="return goodchars(event,'0123456789')">
    <a href="JavaScript:goCardCode();"><img src=images/mini_cvv2.gif border=0></a><% small_help "Card Code" %></TD>
    </TR>
   <% end if %>
<% end sub %>
