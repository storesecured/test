<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<% 
calendar=1
op=Request.QueryString("op")
if op="edit" then
	promotion_id = Request.QueryString("id")
	if promotion_id = "" then
	promotion_id = Request.QueryString("promotion_id")
	end if
	if not isNumeric(promotion_id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	' IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlSelPromo="select * from store_promotions_rules where promotion_id=" & promotion_id & " and store_id=" & store_id
	rs_Store.open sqlSelPromo,conn_store,1,1
	if rs_Store.bof or rs_Store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	PromoName=rs_store("promotion_name")
	PromoStart=rs_store("promotion_start_date")
	PromoEnd=rs_store("promotion_end_date")
	DiscountedItemsIds=rs_store("Discounted_items_ids")
        DiscountedItemsAmount=rs_store("Discounted_items_Amount")
	FreeItems_ids=rs_store("Free_items_ids")
	CustomerGroup=rs_store("Customer_group")
	TotalFrom=rs_store("Total_From")
	TotalTo=rs_store("Total_To")
	Is_Exclusion=rs_store("Is_Exclusion")
	rs_Store.close
	
	if DiscountedItemsIds<>"" then
                sql_select="select distinct item_sku from store_items WITH (NOLOCK) where store_id="&store_id&" and item_id in ("&DiscountedItemsIds&") order by item_sku"
                set myfields=server.createobject("scripting.dictionary")
	       Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	       if noRecords = 0 then
	       FOR rowcounter= 0 TO myfields("rowcount")
	               if DiscountedItemsSKUs="" then
	                       DiscountedItemsSKUs =mydata(myfields("item_sku"),rowcounter)
	               else
	                       DiscountedItemsSKUs =DiscountedItemsSKUs&", "&mydata(myfields("item_sku"),rowcounter)
	               end if
	       Next
	       end if
	end if
	
	if FreeItems_ids<>"" then
                sql_select="select distinct item_sku from store_items WITH (NOLOCK) where store_id="&store_id&" and item_id in ("&FreeItems_ids&") order by item_sku"
                set myfields=server.createobject("scripting.dictionary")
	       Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
	       if noRecords = 0 then
	       FOR rowcounter= 0 TO myfields("rowcount")
	               if FreeItems="" then
	                       FreeItems =mydata(myfields("item_sku"),rowcounter)
	               else
	                       FreeItems =FreeItems&", "&mydata(myfields("item_sku"),rowcounter)
	               end if
	       Next
	       end if
	end if
end if

if FreeItems <> "" then
	 sDivFree="block"
else
	 sDivFree="none"
end if

if DiscountedItemsIds="" then
   sDivskus="none"
else
   sDivskus="block"
end if

End_date =  DateAdd("m", +1, Now())
Start_date = Now()
sInstructions="The discounted items sku field and free item sku fields can each contain a maximum of 4000 characters.  If you need to include more items then you can fit in 4000 characters please create multiple promotions and split your item list as appropriate.   This will ensure that your checkout runs smoothly and quickly.  You will be unable to save if your promotions item lists are greater then 4000 characters."


if op<>"edit" then
      sTitle ="Add Promotion"
      sFullTitle ="Marketing > <a href=promotion_manager.asp class=white>Promotions</a> > Add"
else
      sTitle = "Edit Promotion - "&PromoName
      sFullTitle ="Marketing > <a href=promotion_manager.asp class=white>Promotions</a> > Edit - "&PromoName
end if
sCommonName="Promotion"
sCancel="promotion_manager.asp"
sFormAction = "Promotion_action.asp"
sSubmitName = "Store_Activation_Update"
thisRedirect = "new_promotion.asp"
sFormName = "Create_Promotion"
sMenu = "marketing"
sQuestion_Path = "marketing/promotions.htm"
createHead thisRedirect
sDiv="Skus"
%>
<script>
function showHideRadio(div,field,nest){
  if (document.forms[0].elements(field)[1].checked == true)
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	  obj.display='block'
  }
  	else
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	  obj.display='none'
  }

}
</script>

<%
if Service_Type<5 then
%>
	<tr bgcolor='#FFFFFF'>
		<td colspan=2>
			This feature is not available at your current level of service.<BR><BR>
			SILVER Service or higher is required.
			<a href=billing.asp class=link>Click here to upgrade now.</a>
		</td>
	</tr>
	<% createFoot thisRedirect, 0%>
<%
else 
%>
<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Promotion_Id" value="<%=promotion_id%>">
	


				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Name</b></td>
				<td width="60%" class="inputvalue" colspan=2><input type="text" name="Promotion_name" maxlength=255 size="60" value="<%= PromoName %>">
					<INPUT type="hidden"  name=Promotion_name_C value="Re|String|0|255|||Name">
			 <% small_help "Name" %></td>
				</tr>
	  
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Start Date</b></td>
				<td width="60%" class="inputvalue" colspan=2>
        <SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
      <input name="Promotion_Start_Date" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')"
							<% if op="edit" then %>
								value="<%= PromoStart %>"
							<% else %>
								value="<%= FormatDateTime(Start_date,2) %>"
							<% end if %>			      
						>
						<A HREF="#" onClick="cal1.select(document.forms[0].Promotion_Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Promotion_Start_Date.value=='')?document.forms[0].Promotion_Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>

					<INPUT type="hidden"  name="Promotion_Start_Date_C" value="Re|date|||||Start Date">
			 <% small_help "Start Date" %></td>
			</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>End Date</b></td>
				<td width="60%" class="inputvalue" colspan=2><input name="Promotion_End_Date" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')"
							<% if op="edit" then %>
								value="<%= PromoEnd %>"
							<% else %>
								value="<%= FormatDateTime(End_date,2) %>"
							<% end if %>			      
						>
					<A HREF="#" onClick="cal1.select(document.forms[0].Promotion_End_Date,'anchor2','M/d/yyyy',(document.forms[0].Promotion_End_Date.value=='')?document.forms[0].Promotion_End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
					<INPUT type="hidden"  name="Promotion_end_Date_C" value="Re|date|||||End Date">
			 <% small_help "End Date" %></td>
			</tr>
	  
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Activate Discount</b></td>
				<td width="1%" class="inputvalue"><input class="image" type="checkbox" name="Discounted_Items_Want" value="1"
							<% if op="edit" then %>
								<% if DiscountedItemsAmount<>0 then %>
									checked
								<% end if %>
							<% else %>
								checked
							<% end if %>								
						></td><td width="50%">
						<input type="text" name="Discounted_Items_Amount" size="3" onKeyPress="return goodchars(event,'0123456789.')"
							<% if op="edit" then %>
								value="<%= DiscountedItemsAmount %>"
							<% else %>
								value="0"
							<% end if %>				    
						>
						<INPUT type="hidden"  name=Discounted_Items_Amount_C value="Re|Integer|0|100|||Discount">%
			      <% small_help "Discount" %></td>
		       </tr>
				
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname" rowspan=2><b>Discount Available on</b></td>
				<td width="1%" class="inputvalue" colspan=2><input class="image" type="radio" value="All" name="Discount_Items"
							<% if op="edit" then %>
								<% if DiscountedItemsIds="" then %>
									checked
								<% end if %>
							<% else %>
								checked
							<% end if %>
						>All Items<% small_help "Discount Available On" %></td></tr>
						<tr bgcolor='#FFFFFF'><td nowrap>
						<input class="image" type="radio" value="part" name="Discount_Items"
							<% if op="edit" then %>
								<% if DiscountedItemsIds<>"" and Is_Exclusion=0 then %>
									checked
								<% end if %>
							<% end if %>
						>Selected Items
                                                <BR>
                                                <input class="image" type="radio" value="part_exclude" name="Discount_Items"
							<% if op="edit" then %>
								<% if DiscountedItemsIds<>"" and Is_Exclusion<>0  then %>
									checked
								<% end if %>
							<% end if %>
						>Excluded Items
                                                </td><td>Discounted Item Skus
				<a class="link" href="JavaScript: itemflag=1; goItemsSelect('Discounted_Items_Skus');">Sku Picker</a>
			    <BR>
				<textarea rows="4" name="Discounted_Items_Skus" cols="50"><%= DiscountedItemsSKUs %></textarea>
			    <input type="hidden" name="Discounted_Items_Skus_Value" value="<%= DiscountedItemsSKUs %>">
				<INPUT type="hidden"  name=Discounted_Items_Skus_C value="Op|String|0|4000|||Discounted SKU's">

                                <% small_help "Discounted SKU's" %>
                                </td>
			</tr>

       

				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Activate Free Items</b></td>
				<td width="1%" class="inputvalue"><input class="image" type="checkbox" name="Free_items_want" value="1"
							<% if op="edit" then %>
								<% if FreeItems_Ids<>"" then %>
									checked
								<% end if %>
							<% else %>
								checked
							<% end if %>
						></td><td width="50%">
						Free Item Skus <a class="link" href="JavaScript: itemflag=1; goItemsSelect('Free_items');">Sku Picker</a>
				<BR>
        <% if op="edit" then %>
							<textarea rows="4" name="Free_items" cols="50"><%= FreeItems %></textarea>
						<% else %>
							<textarea rows="4" name="Free_items" cols="50"></textarea>
						<% end if %>
			    <INPUT type="hidden"  name=Free_items_C value="Op|String|0|4000|||Free SKU's">
					<% small_help "Free SKU's" %></td>
			      

			</tr>
        
				
			
	  

				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Customer Group</b></td>
				<td width="60%" class="inputvalue" colspan=2><select size="1" name="Customer_group">
						<option
								 <% if op="edit" then %>
									<% if CustomerGroup=0 then %>
										selected
									<% end if %>
								<% else %>
									selected
								<% end if %>
							value="0">All Customers</option>
						<% 
						sql_groups = "select Group_Name, Group_Id from Store_Customers_Groups where Store_id = "&Store_id&""
						set myfields=server.createobject("scripting.dictionary")
						Call DataGetrows(conn_store,sql_groups,mydata,myfields,noRecords)
						if noRecords = 0 then
						FOR rowcounter= 0 TO myfields("rowcount")

							response.write "<option value='"&mydata(myfields("group_id"),rowcounter)&"'"
							if op="edit" then
								if cint(CustomerGroup)=mydata(myfields("group_id"),rowcounter) then
									response.write "selected"
								end if
							end if
							response.write ">"&mydata(myfields("group_name"),rowcounter)&"</option>"
						Next
						end if
						set myfields= nothing
						%>
						</select><% small_help "Customer Groups" %></td>
				</tr>

				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Purchase Min</b></td>
				<td width="60%" class="inputvalue" colspan=2><%= Store_Currency %><input name="Total_From" size="10" onKeyPress="return goodchars(event,'0123456789.')"
							<% if op="edit" then %>
								value="<%= TotalFrom %>"
							<% else %>
								value="1"
							<% end if %>
						>
						<INPUT type="hidden"  name=Total_From_C value="Re|Integer|0||||Purchase Min">
			      <% small_help "Purchase Min" %></td>
				</tr>
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Purchase Max</b></td>
				<td width="60%" class="inputvalue" colspan=2><%= Store_Currency %><input name="Total_To" size="10" onKeyPress="return goodchars(event,'0123456789.')"
							<% if op="edit" then %>
								value="<%= TotalTo %>"
							<% else %>
								value="99999"
							<% end if %>
						>
						<INPUT type="hidden"  name=Total_To_C value="Re|Integer|0|||Purchase Max">
			      <% small_help "Purchase Max" %></td>
				</tr>



<% createFoot thisRedirect, 1%>
<SCRIPT language="JavaScript">
 var frmvalidator  = new Validator(0);
 frmvalidator.addValidation("Promotion_name","req","Please enter a name for this promotion.");
 frmvalidator.addValidation("Promotion_Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("Promotion_End_Date","date","Please enter a valid end date.");
 frmvalidator.addValidation("Discounted_Items_Amount","req","Please enter a discount percentage between 0 and 100%.");
 frmvalidator.addValidation("Discounted_Items_Amount","lessthan=100.01","Please enter a discount percentage between 0 and 100%.");
 frmvalidator.addValidation("Discounted_Items_Amount","greaterthan=-.01","Please enter a discount percentage between 0 and 100%.");
 frmvalidator.addValidation("Total_From","req","Please enter a minimum purchase amount.");
 frmvalidator.addValidation("Total_To","req","Please enter a maximum purchase amount.");

var globlist;
var DiscVal = "<% = DiscountedItemsSKUs %>";
if(DiscVal != "") {
	DiscVal = DiscVal + ",";
}
var FreeVal = "<% = FreeItems %>";
if(FreeVal != "") {
	FreeVal = FreeVal + ",";
}

</script>
<% end if %>
