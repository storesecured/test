<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<!--#include file="include/countrycode_list.asp"-->
<!--#include file="include/location_list.asp"-->
<!--#include file="help/shipping_all_class.asp"-->


<%

if trim(Request.QueryString("op")) <> "" then
op = Request.QueryString("op")
else
op = "add" 
end if

showArea = trim(request.querystring("type"))

' EDIT : THEN LOAD CURRENT VALUES FROM THE DATABASE
' -----------------------------------------------------
if op = "edit" then
	
	sqlSelShipping="select * from Store_Shippers_class_all where " & _
						" Store_id="&Store_id & " and " & _
						"shipping_method_Id=" & Request.QueryString("Id")
   
	rs_store.open sqlSelShipping,conn_store,1,1
	Shippers_class=rs_store("Shippers_class")
	if showArea="" then
           showArea=Shippers_class
        end if
	ShippingMethodName=rs_store("Shipping_Method_Name")
	BaseFee=formatnumber(rs_store("Base_Fee"),2)
	
	WeightFee=formatnumber(rs_store("Weight_Fee"),2,0,0,0)

	MatrixLow=formatnumber(rs_store("Matrix_Low"),2,0,0,0)
	MatrixHigh=formatnumber(rs_store("Matrix_High"),2,0,0,0)

	Zip_Start=rs_store("Zip_Start")
	Zip_End=rs_store("Zip_End")
	Countries=rs_store("Countries")
	Ship_Location_Id=rs_store("Ship_Location_Id")
     Realtime_backup=rs_store("realtime_backup")
     if Realtime_backup then
     	checked_backup="checked"
     end if

	rs_store.close
end if
' -----------------------------------------------------


sFormAction = "update_records_action.asp"
sName = "Add_Class_1"
sCommonName="Shipping Method"
sCancel = "shipping_all_list.asp"
if request.querystring("Id") <>"" then
        sTitle = "Edit Shipping Method - "&ShippingMethodName
        sFullTitle = "Shipping > <a href=shipping_all_list.asp class=white>Methods</a> > Edit - "&ShippingMethodName
else
        sTitle = "Add Shipping Method"
        sFullTitle = "Shipping > <a href=shipping_all_list.asp class=white>Methods</a> > Add"
end if
thisRedirect = "shipping_all_class.asp"
sMenu = "shipping"
sQuestion_Path = "general/shipping.htm"

if op <> "" then 

   ' Titles as per display area
   '-------------------------------------
   if showArea = "1" then
   s1 = "selected"
   sTitle = "Flat Fee Shipping"
   elseif showArea = "2" then
   s2 = "selected"
   sTitle = "Flat Fee Plus Weight"
   lpounds="pounds"
   sOther="Individual item weights are configured from the item edit screen when using this shipping method."
   elseif showArea = "3" then
   s3 = "selected"
   sTitle = "Per Item"
   sOther="Individual item charges for shipping are configured from the item edit screen when using this shipping method."
   elseif showArea = "4" then
   s4 = "selected"
   sTitle = "Matrix Based on Total Order Amount"
   lcurrency=Store_Currency
   elseif showArea = "5" then
   s5 = "selected"
   sTitle = "Percentage of Total Order"
   lcurrency=Store_Currency
   elseif showArea = "6" then
   s6 = "selected"
   sTitle = "Matrix Based on Total Order Weight"
   lpounds="pounds"
   sOther="Individual item weights are configured from the item edit screen when using this shipping method."
   else
   s0 =  "selected"
   end if
   
   '-------------------------------------
   if op = "add" then
   op = ""
   end if

createHead thisRedirect
%>


<tr bgcolor='#FFFFFF'>
<td width="100%" colspan="3" height="23">
<input OnClick='JavaScript:self.location="Shipping_class_realtime.asp"' name="Shipping_Method_List" type="button" class="Buttons" value="Shipping Settings">
<input OnClick='JavaScript:self.location="location_manager.asp"' name="Shipping_Method_List" type="button" class="Buttons" value="Location List"></td>
</tr>

<input type="hidden" name="op" value="<%=op%>">
<input type="hidden" name="Shipping_Method_Id" value="<%=Request.QueryString("Id")%>">
<input type="hidden" name="Shipping_Form_Back_To" value="shipping_all_list.asp">

<% 
' ADD : Display drop down only in case of Add
' ---------------------------------------------------------------------------
%>
	<tr bgcolor='#FFFFFF'>
	<td width="20%" class="inputname"><B>Shipping Type</b></td>
	<td width="80%" class="inputvalue">
	<select id='ShippingClass' name="Shipping_Class" OnChange="pageRedirect()">
	<option value = "0" <%=s0%> >Select Shipping Type</option>
	<option value = "1" <%=s1%> >Flat Fee</option>
	<option value = "2" <%=s2%> >Flat fee + weight</option>
	<option value = "3" <%=s3%> >Per Item</option>
	<option value = "4" <%=s4%> >Total Order Matrix</option>
	<option value = "5" <%=s5%> >% of total order</option>
	<option value = "6" <%=s6%> >Total Weight Matrix</option>
	</select>
	<% small_help "Shipping Type" %></td>
	</tr>
<%
if instr(Shipping_Classes,showArea)=0 then %>
   <tr bgcolor='#FFFFFF'><TD bgcolor=red colspan=3><font color=white>IMPORTANT: The selected shipping type is currently disabled in
   your store setup.  Any shipping methods of this type that you add will not be shown to your customers unless you 
   enable this shipping type.  Use the Shipping Settings button above to enable/disable shipping types.</font></td></tr>
<% end if

select case showArea
' -----------------------------Shipping Class 1 : Flat Fee -------------------
	case "1"
	%>
	<input type="hidden" name="Add_What_Shipping_Class" value="1">
		   <% 
		   ' CALL : Sub Routine to call different form areas
		    showAllTop
			show13Center "Flat Fee"
			showAllBottom
		   %>
	<SCRIPT language="JavaScript">
	 var frmvalidator  = new Validator(0);
	 frmvalidator.addValidation("Shipping_Method_Name","req","Please enter a name for this shipping method.");
	 frmvalidator.addValidation("Base_Fee","req","Please enter a flat shipping amount.");
	 frmvalidator.addValidation("Countries","req","Please select the countries that the shipping method will apply to.");
	</script>
<%
' -----------------------------Shipping Class 1 : Flat Fee -------------------

' -----------------------------Shipping Class 2 : Flat Fee + Weight ----------
	case "2"
	%>
	<input type="hidden" name="Add_What_Shipping_Class" value="2">

			<% 
			' CALL : Sub Routine to call different form areas
			showAllTop
			show2456Center "2","Base Fee","Weight Between"
			showAllBottom
			%>
					

	<SCRIPT language="JavaScript">
	 var frmvalidator  = new Validator(0);
	 frmvalidator.addValidation("Shipping_Method_Name","req","Please enter a name for this shipping method.");
	 frmvalidator.addValidation("Base_Fee","req","Please enter a base fee.");
	 frmvalidator.addValidation("Weight_Fee","req","Please enter a weight fee.");
	 frmvalidator.addValidation("Matrix_Low","req","Please enter a low range.");
	 frmvalidator.addValidation("Matrix_High","req","Please enter a high range.");
	 frmvalidator.addValidation("Countries","req","Please select the countries that the shipping method will apply to.");
	</script>
<%
' -----------------------------Shipping Class 2 : Flat Fee + Weight ----------

' -----------------------------Shipping Class 3 : No Shipping (Per Item Only)-
	case "3"
	%>
	<input type="hidden" name="Add_What_Shipping_Class" value="3">

			<% 
		   ' CALL : Sub Routine to call different form areas
			showAllTop
			show13Center "Base Fee"
			showAllBottom
			%>
	<SCRIPT language="JavaScript">
	 var frmvalidator  = new Validator(0);
	 frmvalidator.addValidation("Shipping_Method_Name","req","Please enter a name for this shipping method.");
	 frmvalidator.addValidation("Countries","req","Please select the countries that the shipping method will apply to.");
	</script>

<%
' -----------------------------Shipping Class 3 : No Shipping (Per Item Only)-

' -----------------------------Shipping Class 4 :Total Order Matrix-----------
	case "4"
	%>
	<input type="hidden" name="Add_What_Shipping_Class" value="4">

		<% 
	   ' CALL : Sub Routine to call different form areas
		showAllTop
		show2456Center "4","Base Fee","Total Between"
		showAllBottom
		%>
	<SCRIPT language="JavaScript">
	 var frmvalidator  = new Validator(0);
	 frmvalidator.addValidation("Shipping_Method_Name","req","Please enter a name for this shipping method.");
	 frmvalidator.addValidation("Base_Fee","req","Please enter a base fee.");
	 frmvalidator.addValidation("Matrix_Low","req","Please enter a low range.");
	 frmvalidator.addValidation("Matrix_High","req","Please enter a high range.");
	 frmvalidator.addValidation("Countries","req","Please select the countries that the shipping method will apply to.");
	</script>

<%
' -----------------------------Shipping Class 4 :Total Order Matrix-----------


' -----------------------------Shipping Class 5 :% of Total Order  -----------

	case "5"
	%>
	<input type="hidden" name="Add_What_Shipping_Class" value="5">

					<% 
				   ' CALL : Sub Routine to call different form areas
					showAllTop
					show2456Center "5","Percentage of Order","Total Between"
					showAllBottom
					%>
	<SCRIPT language="JavaScript">
	 var frmvalidator  = new Validator(0);
	 frmvalidator.addValidation("Shipping_Method_Name","req","Please enter a name for this shipping method.");
	 frmvalidator.addValidation("Base_Fee","req","Please enter a base fee.");
	 frmvalidator.addValidation("Matrix_Low","req","Please enter a low range.");
	 frmvalidator.addValidation("Matrix_High","req","Please enter a high range.");
	 frmvalidator.addValidation("Countries","req","Please select the countries that the shipping method will apply to.");
	</script>

<%

' -----------------------------Shipping Class 5 :% of Total Order  -----------

' -----------------------------Shipping Class 6 :Total Weight Matrix ----------

	case "6"
	%>
	<input type="hidden" name="Add_What_Shipping_Class" value="6">

			<% 
		   ' CALL : Sub Routine to call different form areas
			showAllTop
			show2456Center "6","Base Fee","Weight Between"
			showAllBottom
			%>
	<SCRIPT language="JavaScript">
	var frmvalidator  = new Validator(0);
	frmvalidator.addValidation("Shipping_Method_Name","req","Please enter a name for this shipping method.");
	frmvalidator.addValidation("Base_Fee","req","Please enter a base fee.");
	frmvalidator.addValidation("Matrix_Low","req","Please enter a low range.");
	frmvalidator.addValidation("Matrix_High","req","Please enter a high range.");
	frmvalidator.addValidation("Countries","req","Please select the countries that the shipping method will apply to.");
	</script>

<%
' -----------------------------Shipping Class 6 :Total Weight Matrix ----------

case else 
	' --  First time no shipping is selected 
	' --- Do not show button
	createFoot thisRedirect,0

end select
' END OF SELECT-CASE STATEMENT


' SUBROUTINES to display different sections of the form
' TOP PART : 
'----------------------------------------------------------------
sub showAllTop
%>
	<tr bgcolor='#FFFFFF'>
	<td width="20%"class="inputname"><B>Name</b></td>
	<td width="80%" class="inputvalue">
	<input type="text" name="Shipping_Method_Name" size="30" maxlength=200 value="<%= ShippingMethodName %>" maxlength=200>*
	<input type="hidden" name="Shipping_Method_Name_C" value="Re|String|0|200|||Shipping Method">
	<% small_help "Name" %></td>
	</tr>
<%
end sub
'----------------------------------------------------------------

' CENTER PART : For shipping type 1,3
'----------------------------------------------------------------
sub show13Center(lFee)
%>
		<tr bgcolor='#FFFFFF'>
		<td width="20%"class="inputname"><B><%=lFee%></b></td>
		<td width="80%" class="inputvalue">
		<%= Store_Currency %><input type="text" name="Base_Fee" size="10" value="<%= BaseFee %>" onKeyPress="return goodchars(event,'-0123456789.')">*
		<input type="hidden" name="Base_Fee_C" value="Re|Integer|||||<%=lFee%>">
		<% small_help ""&lFee&"" %></td>
		</tr>
<%
end sub
'----------------------------------------------------------------

' CENTER PART : For shipping type 2,4,5,6
'----------------------------------------------------------------
sub show2456Center(shType,lFee,lBetween)
%>
		<tr bgcolor='#FFFFFF'>
		<td width="20%"class="inputname"><B><%=lFee%></b></td>
		<td width="80%" class="inputvalue">
		<% if shType<>"5" then %>
                <%= Store_Currency %>
                <% end if %>
                <input type="text" name="Base_Fee" size="10" value="<%= BaseFee %>" onKeyPress="return goodchars(event,'-0123456789.')">
                <% if shType="5" then %>
                %
                <% end if %>*
		<input type="hidden" name="Base_Fee_C" value="Re|Integer|||||<%=lFee%>">
		<% small_help ""&lFee&"" %></td>
		</tr>

		<% if shType = "2" then %>
		<tr bgcolor='#FFFFFF'>
		<td width="20%"class="inputname"><B>Weight Fee</b></td>
		<td width="80%" class="inputvalue">
		<%= Store_Currency %><input type="text" name="Weight_Fee" size="10" value="<%= WeightFee %>" onKeyPress="return goodchars(event,'-0123456789.')">*
		X pounds <input type="hidden" name="Weight_Fee_C" value="Re|Integer|||||Weight Fee">
		<% small_help "Weight Fee" %></td>
		</tr>
		<% end if %>

		<tr bgcolor='#FFFFFF'>
		<td width="20%"class="inputname"><B><%=lBetween%></b></td>
		<td width="80%" class="inputvalue">
		<%= lcurrency %><input type="text" name="Matrix_Low" size="10" value="<%= MatrixLow %>" onKeyPress="return goodchars(event,'-0123456789.')">*
		<%= lpounds %> <input type="hidden" name="Matrix_Low_C" value="Re|Integer|||||Matrix Low">
		and
		<%= lcurrency %><input type="text" name="Matrix_High" size="10" value="<%= MatrixHigh %>" onKeyPress="return goodchars(event,'-0123456789.')">*
		<%= lpounds %> <input type="hidden" name="Matrix_High_C" value="Re|Integer|||||Matrix High">
		<% small_help ""&lBetween&"" %></td>
		</tr>
<%
end sub
'----------------------------------------------------------------

' BOTTOM PART : Common for all shipping types
'---------------------------------------------------------------- 
sub showAllBottom
%>
	<tr bgcolor='#FFFFFF'>
	<td width="20%"class="inputname"><B>Zip Between</b></td>
	<td width="80%" class="inputvalue">
	<input type="text" name="Zip_Start" size="10" value="<%= Zip_Start %>" onKeyPress="return goodchars(event,'0123456789')">
	<input type="hidden" name="Zip_Start_C" value="Op|Integer|||||Zip Start">
	and
	<input type="text" name="Zip_End" size="10" value="<%= Zip_End %>" onKeyPress="return goodchars(event,'0123456789')">
	<input type="hidden" name="Zip_End_C" value="Op|Integer|||||Zip End">
	<% small_help "Zip Between" %></td>
	</tr>

	<tr bgcolor='#FFFFFF'>
	<td align="left" width="20%" height="20" class="inputname"><B>Countries</b></td>
	<td align="left" width="80%" height="20" class="inputvalue">
	<%= create_countrycode_list ("Countries",Countries,6,"All Countries:") %>
	*
	<input type="hidden" name="Countries_C" value="Re|String|||||Countries">
	<% small_help "Countries" %></td>
	</tr>

	<tr bgcolor='#FFFFFF'>
	<td align="left" width="20%" height="20" class="inputname"><B>Ship Location</b></td>
	<td align="left" width="80%" height="20" class="inputvalue" nowrap>
	<%= create_location_list ("Ship_Location_Id",Ship_Location_Id,1) %>
	*
	<% small_help "Ship_Location_Id" %></td>
	</tr>
	<%if sOther<>"" then %>
	<tr bgcolor='#FFFFFF'>
	<td colspan=3><%= sOther %></td>
	</tr>
        <% end if%>
     <tr bgcolor='#FFFFFF'>
	<td align="left" width="20%" height="20" class="inputname"><B>Only Show When No Realtime Rates Apply</b></td>
	<td align="left" width="80%" height="20" class="inputvalue" nowrap>
	<input class="image" name="Realtime_Backup" <%= checked_backup %> type="checkbox" value="-1">
	<% small_help "Realtime_Rate" %></td>
	</tr>
<% createFoot thisRedirect,1%>
<% end sub 
' BOTTOM PART END 
' ------------------------------------------------------------------------------------------------------
%>


<SCRIPT language="JavaScript">
// redirect function
function pageRedirect()
{
 var selVal = document.getElementById('ShippingClass').value
	 if (selVal != "0")
	 {
		window.location = "shipping_all_class.asp?op=<%= request.querystring("op") %>&id=<%= request.querystring("id")%>&type="+selVal
	 }
}
</script>

<% end if%>