<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<% 
calendar=1
op=Request.QueryString("op")
if op="edit" then
  if request.querystring("Id") <> "" then
  	Coupon_Id = Request.QueryString("Id")
	if not isNumeric(Coupon_Id) then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
	sqlCoupons="select * from store_coupons where Coupon_Line_Id=" & Coupon_Id & " and store_id=" & store_id
	rs_Store.open sqlCoupons,conn_store,1,1
	if rs_Store.bof or rs_Store.eof then
		Response.Redirect "admin_error.asp?message_id=1"
	end if
  elseif request.querystring("couponid") <> "" then
    Coupon_Id = Request.QueryString("CouponId")
  	'IF EDIT THEN LOAD CURRENT VALUES FROM THE DATABASE
  	sqlCoupons="select * from store_coupons where Coupon_Id='" & Coupon_Id & "' and store_id=" & store_id
  	rs_Store.open sqlCoupons,conn_store,1,1
  	if rs_Store.bof or rs_Store.eof then
  		Response.Redirect "admin_error.asp?message_id=1"
  	end if
  end if
   Coupon_Line_Id=rs_store("Coupon_Line_Id")
	CouponName=rs_store("Coupon_Name")
	CouponIdUser=rs_store("Coupon_Id")
	CouponStartDate=rs_store("Coupon_Start_Date")
	CouponEndDate=rs_store("Coupon_End_Date")
	CouponType=rs_store("Coupon_Type")
	CouponAmount=rs_store("Coupon_Amount")
	CouponTotalUsage=rs_store("Coupon_Total_Usage")
	CouponCustomerUsage=rs_store("Coupon_Customer_Usage")
	DiscountedItemsIds=rs_store("Discounted_items_ids")
	Total_From=rs_store("Total_From")
	Total_To=rs_store("Total_To")
	Is_Exclusion=rs_store("Is_Exclusion")
  rs_Store.close
        DiscountedItemsSKUs=""
        
        if isNull(DiscountedItemsIds) then
                DiscountedItemsIds=""
        end if

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
end if

if DiscountedItemsIds="" then
   sDivskus="none"
else
   sDivskus="block"
end if


End_date =  DateAdd("m", +1, Now())
Start_date = Now()

if op="edit" then
   sTitle = "Edit Coupon - "&CouponName
   sFullTitle = "Marketing > <a href=coupon_manager.asp class=white>Coupons</a> > Edit - "&CouponName
else
   sTitle = "Add Coupon"
   sFullTitle = "Marketing > <a href=coupon_manager.asp class=white>Coupons</a> > Add"
end if
sCommonName="Coupon"
sCancel="coupon_manager.asp"
sFormAction = "Coupon_action.asp"
thisRedirect = "new_coupon.asp"
sFormName ="Create_Coupon"
sMenu = "marketing"
sQuestion_Path = "marketing/coupons.htm"
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



<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Coupon_Line_Id" value="<%=Coupon_Line_Id%>">


			<tr bgcolor='#FFFFFF'>
				    <td width="100%" colspan="4" height="11">
					    <input type="button" OnClick=JavaScript:self.location="Coupon_settings.asp" class="Buttons" value="Coupon Settings" name="Create_new_Coupon">
						</td>
			</tr>
			<% if Show_Coupon =0 then %>
			<tr bgcolor='red'>
				    <td width="100%" colspan="4" height="11">
					    <font class=white>Coupons are currently disabled in your store.  Your customers will be unable to redeem this coupon unless coupons are enabled.  Go to the coupon settings page to enable coupons.</font></td>
			</tr>
			<% end if %>
		
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Name</b></td>
				<td width="60%" colspan=2 class="inputvalue">
					<input type="text" name="Coupon_name" size="60" value="<%= CouponName %>">
						<INPUT type="hidden"  name=Coupon_name_C value="Re|String|0|150|||Name">
                              <% small_help "Name" %></td>
			</tr>
    
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>ID</B></td>
				<td width="60%" colspan=2 class="inputvalue">
					<input type="text" name="Coupon_Id" size="60" value="<%= CouponIdUser %>">
						<INPUT type="hidden"  name=Coupon_Id_C value="Re|String|0|255|||ID">
                              <% small_help "ID" %></td>
			</tr>
    
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Start Date</B></td>
				<td width="60%" colspan=2 class="inputvalue">
					<SCRIPT LANGUAGE="JavaScript" ID="jscal1">
      var now = new Date();
      var cal1 = new CalendarPopup("testdiv1");
      cal1.showNavigationDropdowns();
      </SCRIPT>
      <input name="Coupon_Start_Date" size="10" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')"
							<% if op="edit" then %>
								value="<%= CouponStartDate %>"
							<% else %>
								value="<%= FormatDateTime(Start_date,2) %>"
							<% end if %>		
						>
						<A HREF="#" onClick="cal1.select(document.forms[0].Coupon_Start_Date,'anchor1','M/d/yyyy',(document.forms[0].Coupon_Start_Date.value=='')?document.forms[0].Coupon_Start_Date.value:null); return false;" TITLE="Start Date" NAME="anchor1" ID="anchor1"><img src=images/calendar.gif border=0></A>
						<INPUT type="hidden"  name="Coupon_Start_Date_C" value="Re|date|||||Start Date">
                              <% small_help "Start Date" %></td>
			</tr>
			       <tr bgcolor='#FFFFFF'>
			       <td width="40%" class="inputname"><B>End Date</B></td>
			       <td width="60%" colspan=2 class="inputvalue" maxlength=10 onKeyPress="return goodchars(event,'0123456789//')">
						<input name="Coupon_End_Date" size="10"
							<% if op="edit" then %>
								value="<%= CouponEndDate %>"
							<% else %>
								value="<%= FormatDateTime(End_date,2) %>"
							<% end if %>		      
						>
									<A HREF="#" onClick="cal1.select(document.forms[0].Coupon_End_Date,'anchor2','M/d/yyyy',(document.forms[0].Coupon_End_Date.value=='')?document.forms[0].Coupon_End_Date.value:null); return false;" TITLE="End Date" NAME="anchor2" ID="anchor2"><img src=images/calendar.gif border=0></A>
						<INPUT type="hidden"  name="Coupon_End_Date_C" value="Re|date|||||End Date">
                              <% small_help "End Date" %></td>
			</tr>
    
				<tr bgcolor='#FFFFFF'>
				<td width="20%" class="inputname" rowspan=2><B>Type</B></td>
				<td width="20%" class="inputvalue">
						<input class="image" type="radio" value="1" name="Coupon_Type"
							<% if op="edit" then %>
								<% if CouponType<>0 then %>
									checked
								<% end if %>
							<% else %>
								checked
							<% end if %>
						>Fixed Amount</td>
				<td width="60%" class="inputvalue"><%= Store_currency %>
						<input type="text" name="Coupon_Fixed_Amount"
							<% if op="edit" then %>
								<% if CouponType<>0 then %>
									value="<%= CouponAmount %>"
								<% else %>
									value="0"
								<% end if %>
							<% else %>
								value="10"
							<% end if %>			    
						size="10">
						<INPUT type="hidden" name=Coupon_Fixed_Amount_C value="Re|Integer|||||Amount">
                              <% small_help "Fixed Amount" %></td>
			</tr>
    
				<tr bgcolor='#FFFFFF'>
				<td width="20%" class="inputvalue">
						<input class="image" type="radio" value="0" name="Coupon_Type"
							<% if op="edit" then %>
								<% if CouponType=0 then %>
									checked
								<% end if %>
							<% end if %>				  
						>Percent Off</td>
				<td width="20%" class="inputvalue">
						<input type="text" name="Coupon_Percent_of_Sale"
							<% if op="edit" then %>
								<% if CouponType=0 then %>
									value="<%= CouponAmount %>"
								<% else %>
									value="0"
								<% end if %>
							<% else %>
								value="10"
							<% end if %>				  
						size="5">%
						<INPUT type="hidden"  name=Coupon_Percent_of_Sale_C value="Re|Integer|||||Percent Off">
                              <% small_help "Percent Off" %></td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname" rowspan=2><b>Discount Available on</b></td>
				<td width="60%" class="inputvalue" colspan=2><input class="image" type="radio" value="All" name="Discount_Items"
							<% if op="edit" then %>
								<% if DiscountedItemsIds="" then %>
									checked
								<% end if %>
							<% else %>
								checked
							<% end if %>
						>All Items<% small_help "Discount Available On" %></tr>
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
                                                </td><td>
			      <a class="link" href="JavaScript: itemflag=1; goItemsSelect('Discounted_Items_Skus');">Sku Picker</a>
			    <BR>
        <% if op="edit" then %>
								<% if ucase(DiscountedItemsSKUs)<>"ALL" then %>
									<textarea rows="2" name="Discounted_Items_Skus" cols="50"><%= DiscountedItemsSKUs %></textarea>
								<% else %>
									<textarea rows="2" name="Discounted_Items_Skus" cols="50"></textarea>
								<% end if %>
							<% else %>
								<textarea rows="2" name="Discounted_Items_Skus" cols="50"></textarea>
							<% end if %>
							<INPUT type="hidden"  name=Discounted_Items_Skus_C value="Op|String|||||Discounted SKU's">
					
				<% small_help "Discounted SKU's" %>
			</tr>
			    
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Total Uses</B></td>
				<td width="60%" colspan=2 class="inputvalue">&nbsp;
						<input type="text" name="Coupon_Total_Usage" onKeyPress="return goodchars(event,'0123456789')"
							<% if op="edit" then %>
								value="<%= CouponTotalUsage %>"
							<% else %>
								value="10"
							<% end if %>				  
						size="5">
						<INPUT type="hidden"  name=Coupon_Total_Usage_C value="Re|Integer|0||||Total Uses">
					times
                         <% small_help "Total Uses" %></td>
			</tr>
    
				<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><B>Per Customer Limit</B></td>
				<td width="60%" colspan=2 class="inputvalue">&nbsp;
						<input type="text" name="Coupon_Customer_Usage" size="5" onKeyPress="return goodchars(event,'0123456789')"
							<% if op="edit" then %>
								value="<%= CouponCustomerUsage %>"
							<% else %>
								value="1"
							<% end if %>
						>
						<INPUT type="hidden"  name=Coupon_Customer_Usage_C value="Re|Integer|0||||Per Customer Limit">
					times
                         <% small_help "Per Customer Limit" %></td>
			</tr>
			
			<tr bgcolor='#FFFFFF'>
				<td width="40%" class="inputname"><b>Purchase Min</b></td>
				<td width="60%" class="inputvalue" colspan=2><%= Store_Currency %><input name="Total_From" size="10" onKeyPress="return goodchars(event,'0123456789.')"
							<% if op="edit" then %>
								value="<%= Total_From %>"
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
								value="<%= Total_To %>"
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
 frmvalidator.addValidation("Coupon_name","req","Please enter a name for this coupon.");
 frmvalidator.addValidation("Coupon_Id","req","Please enter a coupon id.");
 frmvalidator.addValidation("Coupon_Start_Date","date","Please enter a valid start date.");
 frmvalidator.addValidation("Coupon_End_Date","date","Please enter a valid end date.");
 frmvalidator.addValidation("Coupon_Total_Usage","req","Please enter the total number of uses allowed.");
 frmvalidator.addValidation("Coupon_Customer_Usage","req","Please enter the number of uses allowed per customer.");
 frmvalidator.addValidation("Coupon_Total_Usage","greaterthan=0","Please enter the total number of uses allowed.");
 frmvalidator.addValidation("Coupon_Customer_Usage","greaterthan=0","Please enter the number of uses allowed per customer.");
 frmvalidator.setAddnlValidationFunction("DoCustomValidation");

function DoCustomValidation()
{
  var frm = document.forms[0];
  if (document.forms[0].Coupon_Type[0].checked == true)
  {
    if (document.forms[0].Coupon_Fixed_Amount.value == "")
    {
    alert('Please enter a fixed coupon amount.');
    return false;
    }
    else
    {
    return true;
    }
  }
  else
  if (document.forms[0].Coupon_Type[1].checked == true)
  { 
    if (document.forms[0].Coupon_Percent_of_Sale.value == "")
    {
    alert('Please enter a percentage off.');
    return false;
    }
    else
    {
    return true;
    }
  }
  else
  {
    return true;
  }
}

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
