
<!--#include file='Global_Settings.asp'-->
<!--#include file='pagedesign.asp'-->
<%

Template_Id=Request.Querystring("Id")
If Template_Id = "" then
	fn_error "Please select a template"
else

objHelpDict.Add "cancel","Stop current action."
objHelpDict.Add "cancel order","Cancel current order."
objHelpDict.Add "check out","Checkout shopping cart."
objHelpDict.Add "continue","Continue shopping after logging in."
objHelpDict.Add "continue shopping","Continue shopping before checking out."
objHelpDict.Add "delete","Delete a record"
objHelpDict.Add "login","Logs shopper into your store."
objHelpDict.Add "next","Browse to the next set of items."
objHelpDict.Add "no","Answer to question"
objHelpDict.Add "order","Place an item into customers shopping cart."
objHelpDict.Add "prev","Browse to the prev set of items"
objHelpDict.Add "process order","Process order after payment information is entered and shipping type is chosen."
objHelpDict.Add "reset","Reset customer information."
objHelpDict.Add "retrieve cart","Retrieve a saved cart."
objHelpDict.Add "save cart","Save shopping cart for later retrieval or wish list."
objHelpDict.Add "search","Process detailed search."
objHelpDict.Add "update","Update customer information."
objHelpDict.Add "update cart","Add or remove items from your cart."
objHelpDict.Add "view","View items placed in cart."
objHelpDict.Add "yes","Answer to question"

		sql_check="select template_name, template_active from store_design_template where store_id="&Store_Id&" and template_id="&Template_Id
		set rs_check=conn_store.execute(sql_check)

		'EOF check
		if rs_check.EOF then
		   fn_error "This template does not exist"
		else

				templ_active_check=rs_check("template_active")
				templ_name=rs_check("template_name")
				set rs_check=nothing


				sql_Select_Store_Activation = "Select * from store_design_template where Store_id = "&Store_id&" and Template_Id="&Template_Id
					rs_Store.open sql_Select_Store_Activation,conn_store,1,1
					Button_image_Login = Rs_store("Button_image_Login")
					Button_image_Continue = Rs_store("Button_image_Continue")
					Button_image_Reset = Rs_store("Button_image_Reset")
					Button_image_Yes = Rs_store("Button_image_Yes")
					Button_image_No = Rs_store("Button_image_No")
					Button_image_Order = Rs_store("Button_image_Order")
					Button_image_Search = Rs_store("Button_image_Search")
					Button_image_UpdateCart = Rs_store("Button_image_UpdateCart")
					Button_image_ContinueShopping = Rs_store("Button_image_ContinueShopping")
					Button_image_Checkout = Rs_store("Button_image_Checkout")
					Button_image_Cancel = Rs_store("Button_image_Cancel")
					Button_image_SaveCart = Rs_store("Button_image_SaveCart")
					Button_image_RetrieveCart = Rs_store("Button_image_RetrieveCart")
					Button_image_View = Rs_store("Button_image_View")
					Button_image_Update = Rs_store("Button_image_Update")
					Button_image_CancelOrder = Rs_store("Button_image_CancelOrder")
					Button_image_ProcessOrder = Rs_store("Button_image_ProcessOrder")
					Button_image_Next = Rs_store("Button_image_Next")
					Button_image_Prev = Rs_store("Button_image_Prev")
					Button_image_Delete = Rs_store("Button_image_Delete")
					Button_image_Up = Rs_store("Button_image_Up")
				rs_Store.Close

				sFormAction = "process_form.asp"
				sName = ""
				sFormName = "buttons"
				sTitle = "Edit Template Action Buttons - "&templ_name
				sFullTitle = "Design > <a href=template_list.asp class=white>Template</a> > <a href=layout_design.asp?Id="&Template_Id&" class=white>Edit</a> > Action Buttons - "&templ_name
				sCommonName = "Action Buttons"
  		                sCancel = "layout_design.asp?Id="&Template_Id
                                sSubmitName = "Update"
				thisRedirect = "buttons.asp"
				addPicker = 1
				sMenu="design"
				sQuestion_Path = "design/action_buttons.htm"
				createHead thisRedirect
				if Service_Type < 1	then %>
					<TR bgcolor='#FFFFFF'>
					<td colspan=2>
						This feature is not available at your current level of service.<BR><BR>
						PEARL Service or higher is required.
						<a href=billing.asp class=link>Click here to upgrade now.</a>
					</td></tr>
					<% createFoot thisRedirect, 0%>

				<% else %>
					<TR bgcolor='#FFFFFF'>
							 <td width="100%" colspan="3" height="11">
							 <input type="button" OnClick=JavaScript:self.location="template_list.asp" class="Buttons" value="Template List" name="Create_new_Page">&nbsp;<input type="button" OnClick=JavaScript:self.location="layout_design.asp?op=edit&Id=<%=Template_ID%>" class="Buttons" value="Template Detail" name="Create_new_Page"></td>
					</tr>
					<TR bgcolor='#FFFFFF'>
							<td colspan="3" align="center">
								<BR>
								<font color="#ff0000" face="verdana" size="1"><b>Click the Apply Template button to apply this design to your store.</b></font>

							</td>
						</tr>
				<!--Added for bringing the Preview and apply Template buttons-->
				<TR bgcolor='#FFFFFF'>
					<td colspan=3 align="center">
							<table cellspacing="5" cellpadding="0">
									<TR bgcolor='#FFFFFF'>
										<td>	
											<!--<a href="layout_settings.asp?Gen=1&template_id=<%=Template_Id%>">Apply Template</a>-->
											<input type="button" class="buttons" value="Apply Template" onClick="window.location='layout_settings.asp?Gen=1&template_id='+<%=Template_Id%>">							
										</td>
										<td>
											<!--<a href="layout_settings.asp?Gen=1&preview_id=1&template_id=<%=Template_Id%>">Preview Template</a>-->
											<input type="button" class="buttons" value="Preview Template" onClick="window.location='layout_settings.asp?Gen=1&preview_id=1&template_id='+<%=Template_Id%>">
										</td>
									</tr>
								</table>
					</td>
				</TR>
				<!--Added for bringing the Preview and apply Template buttons-->
					  <TR bgcolor='#FFFFFF'><TD colspan=3 class=instructions>Use this page to change the default grey buttons throughout your store.</td></tr>

					  <TR bgcolor='#FFFFFF'><td width='100%' height='25' colspan='3' class=instructions>Choose a picture to display in place of the default buttons or enter text to display on the button.  Leave blank to use default text.<br>
					  </td></tr><TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Login</B></td>
					  <td width='70%' class='inputvalue'><input type='text' name='Button_image_Login' value='<%= Button_image_Login %>' size='60'>
					  <INPUT type='hidden'	name=Button_image_Login_C value='Op|String|0|60|||Login' maxlength=60>
					  <a href="javascript:goImagePicker('Button_image_Login')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Login');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>

					  <% small_help "Login" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Continue</B></td>
					  <td width='70%' class='inputvalue'><input type='text' name='Button_image_Continue' value='<%= Button_image_Continue %>' size='60'>
					  <INPUT type='hidden'	name=Button_image_Continue_C value='Op|String|0|60|||Continue' maxlength=60>
					  <a href="javascript:goImagePicker('Button_image_Continue')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Continue');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Continue" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Reset</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Reset' value='<%= Button_image_Reset %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Reset_C value='Op|String|0|60|||Reset'>
					  <a href="javascript:goImagePicker('Button_image_Reset')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Reset');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Reset" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Yes</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Yes' value='<%= Button_image_Yes %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Yes_C value='Op|String|0|60|||Yes'>
					  <a href="javascript:goImagePicker('Button_image_Yes')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Yes');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Yes" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>No</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_No' value='<%= Button_image_No %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_No_C value='Op|String|0|60|||No'>
					  <a href="javascript:goImagePicker('Button_image_No')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_No');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "No" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Add to Cart</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Order' value='<%= Button_image_Order %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Order_C value='Op|String|0|60|||Order'>
					  <a href="javascript:goImagePicker('Button_image_Order')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Order');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Order" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Search</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Search' value='<%= Button_image_Search %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Search_C value='Op|String|0|60|||Search'>
					  <a href="javascript:goImagePicker('Button_image_Search')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Search');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Search" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Update Cart</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_UpdateCart' value='<%= Button_image_UpdateCart %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_UpdateCart_C value='Op|String|0|60|||Update Cart'>
					  <a href="javascript:goImagePicker('Button_image_UpdateCart')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_UpdateCart');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Update Cart" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Continue Shopping</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_ContinueShopping' value='<%= Button_image_ContinueShopping %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_ContinueShopping_C value='Op|String|0|60|||Continue Shopping'>
					  <a href="javascript:goImagePicker('Button_image_ContinueShopping')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_ContinueShopping');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Continue Shopping" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Check Out</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Checkout' value='<%= Button_image_Checkout %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Checkout_C value='Op|String|0|60|||Checkout'>
					  <a href="javascript:goImagePicker('Button_image_Checkout')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Checkout');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Check Out" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Cancel</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Cancel' value='<%= Button_image_Cancel %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Cancel_C value='Op|String|0|60|||Cancel'>
					  <a href="javascript:goImagePicker('Button_image_Cancel')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Cancel');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Cancel" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Save Cart</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_SaveCart' value='<%= Button_image_SaveCart %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_SaveCart_C value='Op|String|0|60|||Save Cart'>
					  <a href="javascript:goImagePicker('Button_image_SaveCart')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_SaveCart');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Save Cart" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Retrieve Cart</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_RetrieveCart' value='<%= Button_image_RetrieveCart %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_RetrieveCart_C value='Op|String|0|60|||Retrieve Cart'>
					  <a href="javascript:goImagePicker('Button_image_RetrieveCart')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_RetrieveCart');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Retrieve Cart" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>View</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_View' value='<%= Button_image_View %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_View_C value='Op|String|0|60|||View'>
					  <a href="javascript:goImagePicker('Button_image_View')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_View');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "View" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Update</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Update' value='<%= Button_image_Update %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Update_C value='Op|String|0|60|||Update'>
					  <a href="javascript:goImagePicker('Button_image_Update')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Update');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Update" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Cancel Order</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_CancelOrder' value='<%= Button_image_CancelOrder %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_CancelOrder_C value='Op|String|0|60|||CancelOrder'>
					  <a href="javascript:goImagePicker('Button_image_CancelOrder')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_CancelOrder');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Cancel Order" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Process Order</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_ProcessOrder' value='<%= Button_image_ProcessOrder %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_ProcessOrder_C value='Op|String|0|60|||ProcessOrder'>
					  <a href="javascript:goImagePicker('Button_image_ProcessOrder')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_ProcessOrder');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Process Order" %></td>
					  </tr>
					 <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Next</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Next' value='<%= Button_image_Next %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Next_C value='Op|String|0|60|||Next'>
					  <a href="javascript:goImagePicker('Button_image_Next')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Next');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Next" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Prev</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Prev' value='<%= Button_image_Prev %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Prev_C value='Op|String|0|60|||Prev'>
					  <a href="javascript:goImagePicker('Button_image_Prev')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Prev');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Prev" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Up</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Up' value='<%= Button_image_Up %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Up_C value='Op|String|0|60|||Up'>
					  <a href="javascript:goImagePicker('Button_image_Up')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Up');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Prev" %></td>
					  </tr>
					  <TR bgcolor='#FFFFFF'><td width='30%' class='inputname'><B>Delete</B></td><td width='70%' class='inputvalue'>
					  <input type='text' name='Button_image_Delete' value='<%= Button_image_Delete %>' size='60' maxlength=60>
					  <INPUT type='hidden'	name=Button_image_Delete_C value='Op|String|0|60|||Delete'>
					  <a href="javascript:goImagePicker('Button_image_Delete')"><img border='0' src='images/image.gif' width='23' height='22' alt='Image Picker'></a>
					  <a class="link" href="JavaScript:goFileUploader('Button_image_Delete');"><img border="0" src="images/folderup.gif" width="23" height="22" alt="Image Upload"></a>
				<% small_help "Delete" %></td>
					  </tr>

					<input type="hidden" name="Template_Id" value="<%= Template_Id%>">
					<% createFoot thisRedirect,1 %>
					<SCRIPT language="JavaScript">
					 var frmvalidator  = new Validator(0);

					</script>

				<% end if %>
		<% end if			' EOF condition%>
<%end if	 'Template_Id querstring condition%>

