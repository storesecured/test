<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->


<%

Template_Id=Request.Querystring("id")
if Template_Id = "" then
   fn_error "You must select at least one template"
else

		UpAlt_Text="Move Selected Object Up"
		DnAlt_Text="Move Selected Object Down"
		SaveAlt_Text="Save Object Positions"

		sql_check="select template_name, template_html from store_design_template where store_id="&Store_Id&" and template_id="&Template_Id
		rs_Store.open sql_check,conn_store,1,1
	    if not rs_store.eof then
	        templ_name=rs_store("template_name")
			template_html=rs_store("template_html")
	    end if
	    rs_store.close


        sMenu="design"
		sTitle = "Edit Template - "&templ_name
		sFullTitle = "Design > <a href=template_list.asp class=white>Templates</a> > Edit - "&templ_name
		thisRedirect = "template_list.asp"
		createHead thisRedirect
		%>
				</form>
		<%

			Template_Id=Request.Querystring("id")

			%>

			<script LANGUAGE="Javascript">
				function VerifyDelete(OID)
				{
					if(confirm("Are you sure you want to delete this object?"))
					document.location.href="manage_object.asp?Id="+OID+"&Del=1&tid=<%=Template_Id%>"
				}

			</script>


<script language="JavaScript"><!--
				var glob_arrtop=new Array();
				var glob_arrleft=new Array();
				var glob_arrright=new Array();
				var glob_arrbottom=new Array();
				var glob_arrcen_top=new Array();
				var glob_arrcen_bottom=new Array();


				var glob_names=new Array();

				glob_names["top"]="glob_arrtop"
				glob_names["left"]="glob_arrleft"
				glob_names["right"]="glob_arrright"
				glob_names["bottom"]="glob_arrbottom"
				glob_names["cen_top"]="glob_arrcen_top"
				glob_names["cen_bottom"]="glob_arrcen_bottom"


				function up(area,id)
				{
						if (id>0)
						{
							var i;
							temp=window.document.getElementById(area+id).innerHTML 
							window.document.getElementById(area+id).innerHTML=window.document.getElementById(area+(id-1)).innerHTML;
							window.document.getElementById(area+(id-1)).innerHTML=temp
							

							if (glob_names[area] == "glob_arrtop")
							{				
								t=glob_arrtop[id];
								glob_arrtop[id]=glob_arrtop[id-1];
								glob_arrtop[id-1]=t;
							}
							else if (glob_names[area] == "glob_arrleft")
							{				
								t=glob_arrleft[id];
								glob_arrleft[id]=glob_arrleft[id-1];
								glob_arrleft[id-1]=t;
							}
							
							else if (glob_names[area] == "glob_arrright")
							{				
								t=glob_arrright[id];
								glob_arrright[id]=glob_arrright[id-1];
								glob_arrright[id-1]=t;
							}
							
							else if (glob_names[area] == "glob_arrbottom")
							{					
								t=glob_arrbottom[id];
								glob_arrbottom[id]=glob_arrbottom[id-1];
								glob_arrbottom[id-1]=t;
							}
							
							else if (glob_names[area] == "glob_arrcen_top")
							{					
								t=glob_arrcen_top[id];
								glob_arrcen_top[id]=glob_arrcen_top[id-1];
								glob_arrcen_top[id-1]=t;
							}
							else if (glob_names[area] == "glob_arrcen_bottom")
							{					
								t=glob_arrcen_bottom[id];
								glob_arrcen_bottom[id]=glob_arrcen_bottom[id-1];
								glob_arrcen_bottom[id-1]=t;
							}

							sel(area,id-1)

						}
						else
							alert("you are already at the top of the list")


				}

				function dn(area,id)
				{
						if (id < eval(glob_names[area]).length-1)
						{
							temp=window.document.getElementById(area+id).innerHTML 
							window.document.getElementById(area+id).innerHTML=window.document.getElementById(area+(id+1)).innerHTML;
							window.document.getElementById(area+(id+1)).innerHTML=temp;

							if (glob_names[area] == "glob_arrtop")
							{				
								t=glob_arrtop[id];
								glob_arrtop[id]=glob_arrtop[id+1];
								glob_arrtop[id+1]=t;
							}
							else if (glob_names[area] == "glob_arrleft")
							{				
								t=glob_arrleft[id];
								glob_arrleft[id]=glob_arrleft[id+1];
								glob_arrleft[id+1]=t;
							}
							
							else if (glob_names[area] == "glob_arrright")
							{				
								t=glob_arrright[id];
								glob_arrright[id]=glob_arrright[id+1];
								glob_arrright[id+1]=t;
							}
							
							else if (glob_names[area] == "glob_arrbottom")
							{					
								t=glob_arrbottom[id];
								glob_arrbottom[id]=glob_arrbottom[id+1];
								glob_arrbottom[id+1]=t;
							}
							
							else if (glob_names[area] == "glob_arrcen_top")
							{					
								t=glob_arrcen_top[id];
								glob_arrcen_top[id]=glob_arrcen_top[id+1];
								glob_arrcen_top[id+1]=t;
							}
							else if (glob_names[area] == "glob_arrcen_bottom")
							{					
								t=glob_arrcen_bottom[id];
								glob_arrcen_bottom[id]=glob_arrcen_bottom[id+1];
								glob_arrcen_bottom[id+1]=t;
							}


							sel(area,id+1)
						}
						else
							alert("you are already at the bottom of the list")



				}


				var curr_selected=new Array()

				curr_selected["top"]=0
				curr_selected["left"]=0
				curr_selected["right"]=0
				curr_selected["bottom"]=0
				curr_selected["cen_top"]=0
				curr_selected["cen_bottom"]=0


				function sel(area,Id)
				{
					document.getElementById(area+curr_selected[area]).style.backgroundColor="#ffffff"
					curr_selected[area]=Id
					document.getElementById(area+Id).style.backgroundColor="#dddddd"

				}
				//--></script>


				<TR bgcolor='#FFFFFF'>
					<td width="100%" colspan="3" height="11">
					<input type="button" OnClick=JavaScript:self.location="template_list.asp" class="Buttons" value="Template List" name="Back_to_Templates">&nbsp;<input type="button" OnClick=JavaScript:self.location="custom_template.asp" class="Buttons" value="Add Template" name="Create_new_Template">&nbsp;<input type="button" OnClick=JavaScript:self.location="custom_template_copy.asp?Id=<%=Template_ID%>" class="Buttons" value="Copy Template" name="Copy_Template">
					</td>
				</tr>

			
				<TR bgcolor='#FFFFFF'>
					<td colspan="3" align="center">
					<BR>
					<font color="#ff0000" face="verdana" size="1"><b>Click the Apply Template button to apply this design to your store.</b></font>
					</td>
				</tr>


			
						
				<TR bgcolor='#FFFFFF'>
					<td width="100%" colspan="3" height="11" align="center">
					<input type="button" OnClick=JavaScript:self.location="designer_template.asp?Id=<%=Template_ID%>" class="Buttons" value="Header Footer Page" name="Header_Footer_Page">&nbsp;
					<input type="button" OnClick=JavaScript:self.location="bck_font.asp?Id=<%=Template_ID%>" class="Buttons" value="Color and Text" name="Create_new_Page">&nbsp;
					<input type="button" OnClick=JavaScript:self.location="buttons.asp?Id=<%=Template_ID%>" class="Buttons" value="Action Buttons" name="Action_Buttons_Page"><br />
					<input type="button" OnClick=JavaScript:self.location="custom_template.asp?op=edit&Id=<%=Template_ID%>" class="Buttons" value="General Info" name="Info_Page">
					<input type="button" OnClick=JavaScript:self.location="nav_layout.asp?Id=<%=Template_ID%>" class="Buttons" value="Advanced Nav Layout" name="Nav_Layout">
					</td>
				</tr>
				
						<!-- BUTTON ROWS -->

						<TR bgcolor='#FFFFFF'>
							<td colspan="3" align="center">
								<table cellspacing="5" cellpadding="0">
									<TR bgcolor='#FFFFFF'>
										<td>	
											<form name="frm_gendesign1" method="post" action="layout_settings.asp">
												<input type="hidden" name="Gen" value="1">
												<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
												<%If templ_active_check=0 then%>
													<input class="buttons" type="button" name="GenSubmit" value="Apply Template Now!" OnClick='if (confirm("Are you sure you want to apply this template to your store?")) JavaScript:document.frm_gendesign1.submit();' >
												<% else %>
													<input class="buttons" type="button" name="GenSubmit" value="Apply Template" OnClick='if (confirm("Are you sure you want to apply this template to your store?")) JavaScript:document.frm_gendesign1.submit();' >
												<% end if %> 
											</form>
										</td>
										<td>
											<form name="frm_gendesign2" method="post" action="layout_settings.asp">
												<input type="hidden" name="Gen" value="1">
												<input type="hidden" name="Preview_Id" id="Preview_Id" value="1">
												<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
												<input class="buttons" type="submit" name="GenSubmit" value="Preview Template">
											</form>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						
				
           <tr><td width="100%" colspan="3" align="center">
  <%          if instr(template_html,"_DES_OBJECTS_OBJ%")>0 then %>
				
				<table name=Main cellspacing=0 cellpadding=0 border=1 width="100%">

				<!-- top  starts -->
					<TR bgcolor='#FFFFFF'>
						<td valign="top">
							<table name=topMain cellspacing=5 cellpadding=0 border=0  height="100%"width="100%">

							<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0  width="100%">
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=1&tid=<%=Template_Id%>" style="color:#000000">Top</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=1&tid=<%=Template_Id%>" style="color:#000000"> Edit</a> <a href="manage_object.asp?area=1&op=add&tid=<%=Template_Id%>" style="color:#000000">Add</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
						
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
									
									<form name="frm_toporder" method="post" action="layout_settings.asp"> 

									  <div id=topbutton1 style="display:block">
							
										 <img src="images/up.gif" alt="<%=UpAlt_Text%>" onclick="up('top',curr_selected['top'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif" alt="<%=DnAlt_Text%>" onclick="dn('top',curr_selected['top'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"  onclick="document.frm_toporder.Left_Order.value=glob_arrtop;document.forms['frm_toporder'].submit()" style="cursor:pointer;cursor:hand">
											
											<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
											<input type="hidden" name="Left_Order" id="Left_Order" value="">
									 
									 </div>	

									</td>
								</tr>
							<!-- object 2  ends-->



									<%

									sql_obj="select * from store_design_objects where design_area=1 and store_id="&Store_Id&" and template_id="&Template_Id&" order by view_order"
									set rs_obj=conn_store.execute(sql_obj)
									objcount=-1
									sameLine=0
									while not (rs_obj.EOF)
									objcount=objcount+1
									
									%>

									<!-- object start -->
										<script language="javascript"> 
											glob_arrtop[<%=objcount%>]=<%=rs_Obj("Object_Id")%>;
										</script>
					
							
							<!-- object 2 starts -->
							<% if objcount<>0 and sameLine=1 then %>
							   <td valign="top">
                                                        <% else %>
								<TR bgcolor='#FFFFFF'><td valign="top"><table border=0 cellspacing=0 cellpadding=0><TR bgcolor='#FFFFFF'><td valign="top">
							<% end if %>


         <!-- div to show selection and move operation -->
									<div id="top<%=objcount%>"  style="background-color:#ffffff" onclick="sel('top',<%=objcount%>)">

										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <TR bgcolor='#FFFFFF'>
											<td valign="top">
											
											<% if rs_Obj("Object_Type")="Image" or rs_Obj("Object_Type")="Banner" then%>
												<img src="<% =Site_Name&"images/"%><%=rs_Obj("Image_Path")%>" alt="<%=rs_Obj("Image_Alt")%>" border="0"/>

											<% elseif rs_Obj("Object_Type")="HTML Text" then%>

												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=fn_replacePath(rs_Obj("HTML_Text"))%></font>
											<% else%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_Obj("Object_Type")%></font>
											<% end if%>

											</td>
										  </tr>
										  <tr > 
											<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?Id=<%=rs_Obj("Object_Id")%>&op=edit&tid=<%=Template_Id%>" style="color:#000000"> 
											  Edit</a> <a href="javascript:VerifyDelete('<%=rs_Obj("Object_Id")%>');"  style="color:#000000" >
											  Delete</a> <a href="javascript:onclick=document.getElementById('top<%=objcount%>').onclick()" style="color:#000000">Select</a></font></td>
										  </tr>
										  
										</table>
									
         </div>
         <!-- div ends -->




                                                                <% sameLine=1
                                				if rs_Obj("same_line") = 0 then
                                				   sameLine=0 %>
                                                                    </td></tr></table></td></tr>
                                                                <% else %>
                                                                    </td>
                                				<% end if %>
                                    
								<!-- object 2  ends-->
					
									
									<!-- object end -->
									<% 
									
									rs_obj.movenext
									Wend

									if sameLine<>0 then %>

                                                                           </tr></table></td></tr>
                                                                        <% end if %>

								<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">

										<div id=topbutton2 style="display:block">
											 <img src="images/up.gif" alt="<%=UpAlt_Text%>" onclick="up('top',curr_selected['top'])" style="cursor:pointer;cursor:hand">
											<img src="images/down.gif" alt="<%=DnAlt_Text%>" onclick="dn('top',curr_selected['top'])" style="cursor:pointer;cursor:hand">
											<img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"   onclick="document.frm_toporder.Left_Order.value=glob_arrtop;document.forms['frm_toporder'].submit()" style="cursor:pointer;cursor:hand">
										</div>


									</td>
								</tr>
								<!-- object 2  ends-->

								<script language="javascript">
									if (glob_arrtop.length <=0)
									{
										document.getElementById("topbutton1").style.display="none"
										document.getElementById("topbutton2").style.display="none"
									}
								</script>

								</form>



							</table>
						</td>
					</tr>
				<!-- top  ends -->


					<TR bgcolor='#FFFFFF'>
						<td>
							<table name=middleMain cellspacing=0 cellpadding=0 border=1 height="100%" width="100%">
								<TR bgcolor='#FFFFFF'>

									<!-- left  starts -->
									<td>
										<table name=leftMain cellspacing=5 cellpadding=0 border=0 >
											
					
							<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0 width="100%">
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=2&tid=<%=Template_Id%>" style="color:#000000">Left</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=2&tid=<%=Template_Id%>" style="color:#000000"> Edit</a> <a href="manage_object.asp?area=2&op=add&tid=<%=Template_Id%>" style="color:#000000">Add</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<form name="frm_leftorder" method="post" action="layout_settings.asp">

										<div id=leftbutton1 style="display:block">
											  <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('left',curr_selected['left'])" style="cursor:pointer;cursor:hand">
											<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('left',curr_selected['left'])" style="cursor:pointer;cursor:hand">
											<img src="images/savepos.gif"  alt="<%=SaveAlt_Text%>"  onclick="document.frm_leftorder.Left_Order.value=glob_arrleft;document.forms['frm_leftorder'].submit()" style="cursor:pointer;cursor:hand">

											<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
											<input type="hidden" name="Left_Order" id="Left_Order" value="">
										<div>

									</td>
								</tr>
							<!-- object 2  ends-->


									<%

									sql_obj="select * from store_design_objects where design_area=2 and store_id="&Store_Id&" and template_id="&Template_Id&" order by view_order"
									set rs_obj=conn_store.execute(sql_obj)
									objcount=-1
									sameLine=0
									while not (rs_obj.EOF)
									objcount=objcount+1	
									
									%>

								<!-- object start -->
										<script language="javascript"> 
											glob_arrleft[<%=objcount%>]=<%=rs_Obj("Object_Id")%>;
										</script>
					
							
							<!-- object 2 starts -->
								<% if objcount<>0 and sameLine=1 then %>
							   <td valign="top">
                                                        <% else %>
								<TR bgcolor='#FFFFFF'><td valign="top"><table border=0 cellspacing=0 cellpadding=0><TR bgcolor='#FFFFFF'><td valign="top">
							<% end if %>
									
									<!-- div to show selection and move operation -->
									<div id="left<%=objcount%>"  style="background-color:#ffffff" onclick="sel('left',<%=objcount%>)">
										
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <TR bgcolor='#FFFFFF'> 
											<td valign="top">
											
											<% if rs_Obj("Object_Type")="Image" or rs_Obj("Object_Type")="Banner" then%>
													<img src="<% =Site_Name&"images/"%><%=rs_Obj("Image_Path")%>" alt="<%=rs_Obj("Image_Alt")%>" border="0"/>

											<% elseif rs_Obj("Object_Type")="HTML Text" then%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=fn_replacePath(rs_Obj("HTML_Text"))%></font>
											<% else%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_Obj("Object_Type")%></font>
											<% end if%>

											</td>
										  </tr>
										  <TR> 
										  <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?Id=<%=rs_Obj("Object_Id")%>&op=edit&tid=<%=Template_Id%>" style="color:#000000"> 
											  Edit</a> <a href="javascript:VerifyDelete('<%=rs_Obj("Object_Id")%>');"  style="color:#000000" > 
											  Delete</a> <a href="javascript:onclick=document.getElementById('left<%=objcount%>').onclick()" style="color:#000000">Select</a></font></td>
										  </tr>
										</table>
									
									</div>
									<!-- div ends -->



									<% sameLine=1
                                				if rs_Obj("same_line") = 0 then
                                				   sameLine=0 %>
                                                                    </td></tr></table></td></tr>
                                                                <% else %>
                                                                    </td>
                                				<% end if %>

								<!-- object 2  ends-->

									
									<!-- object end -->
									<% 
									
									rs_obj.movenext
									Wend

									if sameLine<>0 then %>

                                                                           </tr></table></td></tr>
                                                                        <% end if %>
									

								<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">

										 <div id=leftbutton2 style="display:block">
										<img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('left',curr_selected['left'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('left',curr_selected['left'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"  onclick="document.frm_leftorder.Left_Order.value=glob_arrleft;document.forms['frm_leftorder'].submit()" style="cursor:pointer;cursor:hand">
										</div>
																			
									</td>
								</tr>
								<!-- object 2  ends-->

									<script language="javascript">
									if (glob_arrleft.length <=0)
									{
										document.getElementById("leftbutton1").style.display="none"
										document.getElementById("leftbutton2").style.display="none"
									}
								</script>

								</form>


							</table>
									</td>
									<!-- left ends -->

									<!-- center content  starts -->
									<td valign="top">
										<table name=centerMain cellspacing=0 cellpadding=0 border=1 height="100%" width="100%">
											<TR bgcolor='#FFFFFF'>
												<td valign="top">
													<table name=cen_topMain cellspacing=5 cellpadding=0 border=0 height=25 width="100%">
														
							


							<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0 height="100%" width="100%">
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=5&tid=<%=Template_Id%>" style="color:#000000">Center Top</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=5&tid=<%=Template_Id%>" style="color:#000000"> Edit</a> <a href="manage_object.asp?area=5&op=add&tid=<%=Template_Id%>" style="color:#000000">Add</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<form name="frm_cen_toporder" method="post" action="layout_settings.asp">

										 <div id=cen_topbutton1 style="display:block">
											  <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('cen_top',curr_selected['cen_top'])" style="cursor:pointer;cursor:hand">
											<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('cen_top',curr_selected['cen_top'])" style="cursor:pointer;cursor:hand">
											<img src="images/savepos.gif"  alt="<%=SaveAlt_Text%>" onclick="document.frm_cen_toporder.Left_Order.value=glob_arrcen_top;document.forms['frm_cen_toporder'].submit()" style="cursor:pointer;cursor:hand">

											<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
											<input type="hidden" name="Left_Order" id="Left_Order" value="">
										</div>

									</td>
								</tr>
							<!-- object 2  ends-->



										<%

									sql_obj="select * from store_design_objects where design_area=5 and store_id="&Store_Id&" and template_id="&Template_Id&" order by view_order"
									set rs_obj=conn_store.execute(sql_obj)
									objcount=-1
									sameLine=0
									while not (rs_obj.EOF)
									objcount=objcount+1	
									
									%>

									<!-- object start -->
										<script language="javascript"> 
											glob_arrcen_top[<%=objcount%>]=<%=rs_Obj("Object_Id")%>;
										</script>
					
							
							<!-- object 2 starts -->
								<% if objcount<>0 and sameLine=1 then %>
							   <td valign="top">
                                                        <% else %>
								<TR bgcolor='#FFFFFF'><td valign="top"><table border=0 cellspacing=0 cellpadding=0><TR bgcolor='#FFFFFF'><td valign="top">
							<% end if %>

									<!-- div to show selection and move operation -->
								<div id="cen_top<%=objcount%>"  style="background-color:#ffffff" onclick="sel('cen_top',<%=objcount%>)">
										
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <TR bgcolor='#FFFFFF'> 
											<td valign="top">
											
											<% if rs_Obj("Object_Type")="Image" then%>
												<img src="<% =Site_Name&"images/"%><%=rs_Obj("Image_Path")%>" alt="<%=rs_Obj("Image_Alt")%>" border="0"/>

											<% elseif rs_Obj("Object_Type")="HTML Text" then%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=fn_replacePath(rs_Obj("HTML_Text"))%></font>
											<% else%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_Obj("Object_Type")%></font>
											<% end if%>

											</td>
										  </tr>
										  <TR> 
										 <td>
												<font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?Id=<%=rs_Obj("Object_Id")%>&op=edit&tid=<%=Template_Id%>" style="color:#000000"> 
											  Edit</a> <a href="javascript:VerifyDelete('<%=rs_Obj("Object_Id")%>');"  style="color:#000000" > 
											  Delete</a> <a href="javascript:onclick=document.getElementById('cen_top<%=objcount%>').onclick()" style="color:#000000">Select</a></font></td>
										  </tr>
										</table>
									
									</div>
									<!-- div ends -->



									<% sameLine=1
                                				if rs_Obj("same_line") = 0 then
                                				   sameLine=0 %>
                                                                    </td></tr></table></td></tr>
                                                                <% else %>
                                                                    </td>
                                				<% end if %>

								<!-- object 2  ends-->
					
									
									<!-- object end -->
									<% 
									
									rs_obj.movenext
									Wend

									if sameLine<>0 then %>

                                                                           </tr></table></td></tr>
                                                                        <% end if %>
								<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">

										  <div id=cen_topbutton2 style="display:block">
										 <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('cen_top',curr_selected['cen_top'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('cen_top',curr_selected['cen_top'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif"  alt="<%=SaveAlt_Text%>"  onclick="document.frm_cen_toporder.Left_Order.value=glob_arrcen_top;document.forms['frm_cen_toporder'].submit()" style="cursor:pointer;cursor:hand">
									</div>
																			
									</td>
								</tr>
								<!-- object 2  ends-->

								<script language="javascript">
									if (glob_arrcen_top.length <=0)
									{
										document.getElementById("cen_topbutton1").style.display="none"
										document.getElementById("cen_topbutton2").style.display="none"
									}
								</script>

								</form>


													</table>

												</td>
											</tr>

											<TR bgcolor='#FFFFFF'>
												<td valign="top" height="80%">
							
													<table name=cen_cenMain cellspacing=5 cellpadding=0 border=0 height="100%" width="100%">
														<TR bgcolor='#FFFFFF'>
															<td valign="top">





				<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0>
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=7&tid=<%=Template_Id%>" style="color:#000000">Center</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=7&tid=<%=Template_Id%>" style="color:#000000"> Edit</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										  <input class="buttons" type="button" name="Submit3" value="Page Manager" onclick="document.location.href='page_manager.asp'">
									</td>
								</tr>
							<!-- object 2  ends-->


															</td>
														</tr>
													</table>

												</td>
											</tr>

											<TR bgcolor='#FFFFFF'>
												<td valign="bottom">
							
													<table name=cen_botMain cellspacing=5 cellpadding=0 border=0  height=25 width="100%">
														
							

				<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0 height="100%" width="100%">
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=6&tid=<%=Template_Id%>" style="color:#000000">Center Bottom</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=6&tid=<%=Template_Id%>" style="color:#000000"> Edit</a> <a href="manage_object.asp?area=6&op=add&tid=<%=Template_Id%>" style="color:#000000">Add</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<form name="frm_cen_bottomorder" method="post" action="layout_settings.asp">

									  <div id=cen_bottombutton1 style="display:block">
									
										 <img src="images/up.gif" alt="<%=UpAlt_Text%>" onclick="up('cen_bottom',curr_selected['cen_bottom'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('cen_bottom',curr_selected['cen_bottom'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"  onclick="document.frm_cen_bottomorder.Left_Order.value=glob_arrcen_bottom;document.forms['frm_cen_bottomorder'].submit()" style="cursor:pointer;cursor:hand">

										<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
										<input type="hidden" name="Left_Order" id="Left_Order" value="">

									</div>

									</td>
								</tr>
							<!-- object 2  ends-->


									<%

									sql_obj="select * from store_design_objects where design_area=6 and store_id="&Store_Id&" and template_id="&Template_Id&" order by view_order"
									set rs_obj=conn_store.execute(sql_obj)
									objcount=-1
									sameLine=0
									while not (rs_obj.EOF)
									objcount=objcount+1	
									
									%>

									<!-- object start -->
										<script language="javascript"> 
											glob_arrcen_bottom[<%=objcount%>]=<%=rs_Obj("Object_Id")%>;
										</script>
									


					<!-- object 2 starts -->
								<% if objcount<>0 and sameLine=1 then %>
							   <td valign="top">
                                                        <% else %>
								<TR bgcolor='#FFFFFF'><td valign="top"><table border=0 cellspacing=0 cellpadding=0><TR bgcolor='#FFFFFF'><td valign="top">
							<% end if %>
									
									<!-- div to show selection and move operation -->
								<div id="cen_bottom<%=objcount%>"  style="background-color:#ffffff" onclick="sel('cen_bottom',<%=objcount%>)">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <TR bgcolor='#FFFFFF'> 
											<td valign="top">
											
											<% if rs_Obj("Object_Type")="Image" then%>
												<img src="<% =Site_Name&"images/"%><%=rs_Obj("Image_Path")%>" alt="<%=rs_Obj("Image_Alt")%>" border="0"/>

											<% elseif rs_Obj("Object_Type")="HTML Text" then%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=fn_replacePath(rs_Obj("HTML_Text"))%></font>
											<% else%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_Obj("Object_Type")%></font>
											<% end if%>

											</td>
										  </tr>
										  <TR> 
											<td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?Id=<%=rs_Obj("Object_Id")%>&op=edit&tid=<%=Template_Id%>" style="color:#000000"> 
											  Edit</a> <a href="javascript:VerifyDelete('<%=rs_Obj("Object_Id")%>');"  style="color:#000000" > 
											  Delete</a> <a href="javascript:onclick=document.getElementById('cen_bottom<%=objcount%>').onclick()" style="color:#000000">Select</a></font></td>
										  </tr>
										</table>
									
									</div>
									<!-- div ends -->



									<% sameLine=1
                                				if rs_Obj("same_line") = 0 then
                                				   sameLine=0 %>
                                                                    </td></tr></table></td></tr>
                                                                <% else %>
                                                                    </td>
                                				<% end if %>

								<!-- object 2  ends-->
					

									<!-- object end -->
									<% 
									
									rs_obj.movenext
									Wend

									if sameLine<>0 then %>

                                                                           </tr></table></td></tr>
                                                                        <% end if %>

								<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">

									 <div id=cen_bottombutton2 style="display:block">
										 <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('cen_bottom',curr_selected['cen_bottom'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('cen_bottom',curr_selected['cen_bottom'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"  onclick="document.frm_cen_bottomorder.Left_Order.value=glob_arrcen_bottom;document.forms['frm_cen_bottomorder'].submit()" style="cursor:pointer;cursor:hand">
									 </div>
																			
									</td>
								</tr>
								<!-- object 2  ends-->

							
								<script language="javascript">
									if (glob_arrcen_bottom.length <=0)
									{
										document.getElementById("cen_bottombutton1").style.display="none"
										document.getElementById("cen_bottombutton2").style.display="none"
									}
								</script>
								</form>



													</table>
										
												</td>
											</tr>
					
										</table>

									</td>
									<!-- center content  starts -->

									<!-- right  starts -->
									<td>
										<table name=rightMain cellspacing=5 cellpadding=0 border=0>


				<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0 width="100%">
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=3&tid=<%=Template_Id%>" style="color:#000000">Right</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=3&tid=<%=Template_Id%>" style="color:#000000"> Edit</a> <a href="manage_object.asp?area=3&op=add&tid=<%=Template_Id%>" style="color:#000000">Add</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
									<form name="frm_rightorder" method="post" action="layout_settings.asp">           

										<div id=rightbutton1 style="display:block">
											 <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('right',curr_selected['right'])" style="cursor:pointer;cursor:hand">
											<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('right',curr_selected['right'])" style="cursor:pointer;cursor:hand">
											<img src="images/savepos.gif"  alt="<%=SaveAlt_Text%>"  onclick="document.frm_rightorder.Left_Order.value=glob_arrright;document.forms['frm_rightorder'].submit()" style="cursor:pointer;cursor:hand">
										
										<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
										<input type="hidden" name="Left_Order" id="Left_Order" value="">
										</div>	

									</td>
								</tr>
							<!-- object 2  ends-->





							
									<%

									sql_obj="select * from store_design_objects where design_area=3 and store_id="&Store_Id&" and template_id="&Template_Id&" order by view_order"
									set rs_obj=conn_store.execute(sql_obj)
									objcount=-1
									sameLine=0
									while not (rs_obj.EOF)
									objcount=objcount+1	
									
									%>

									<!-- object start -->
										<script language="javascript"> 
											glob_arrright[<%=objcount%>]=<%=rs_Obj("Object_Id")%>;
										</script>
					
							
							<!-- object 2 starts -->
								<% if objcount<>0 and sameLine=1 then %>
							   <td valign="top">
                                                        <% else %>
								<TR bgcolor='#FFFFFF'><td valign="top"><table border=0 cellspacing=0 cellpadding=0><TR bgcolor='#FFFFFF'><td valign="top">
							<% end if %>
									
									<!-- div to show selection and move operation -->
									<div id="right<%=objcount%>"  style="background-color:#ffffff" onclick="sel('right',<%=objcount%>)">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <TR bgcolor='#FFFFFF'> 
											<td valign="top">
											
											<% if rs_Obj("Object_Type")="Image" then%>
												<img src="<% =Site_Name&"images/"%><%=rs_Obj("Image_Path")%>" alt="<%=rs_Obj("Image_Alt")%>" border="0"/>

											<% elseif rs_Obj("Object_Type")="HTML Text" then%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=fn_replacePath(rs_Obj("HTML_Text"))%></font>
											<% else%>
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_Obj("Object_Type")%></font>
											<% end if%>

											</td>
										  </tr>
										  <TR> 
											 <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?Id=<%=rs_Obj("Object_Id")%>&op=edit&tid=<%=Template_Id%>" style="color:#000000"> 
											  Edit</a> <a href="javascript:VerifyDelete('<%=rs_Obj("Object_Id")%>');"  style="color:#000000" > 
											  Delete</a> <a href="javascript:onclick=document.getElementById('right<%=objcount%>').onclick()" style="color:#000000">Select</a></font></td>
										  </tr>
										</table>
									
									</div>
									<!-- div ends -->



									<% sameLine=1
                                				if rs_Obj("same_line") = 0 then
                                				   sameLine=0 %>
                                                                    </td></tr></table></td></tr>
                                                                <% else %>
                                                                    </td>
                                				<% end if %>

								<!-- object 2  ends-->
					

									<!-- object end -->
									<% 
									
									rs_obj.movenext
									Wend

									if sameLine<>0 then %>

                                                                           </tr></table></td></tr>
                                                                        <% end if %>

								<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">

									 <div id=rightbutton2 style="display:block">
									  <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('right',curr_selected['right'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('right',curr_selected['right'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif"  alt="<%=SaveAlt_Text%>"  onclick="document.frm_rightorder.Left_Order.value=glob_arrright;document.forms['frm_rightorder'].submit()" style="cursor:pointer;cursor:hand">


									  </div>
																			
									</td>
								</tr>
								<!-- object 2  ends-->

										<script language="javascript">
									if (glob_arrright.length <=0)
									{
										document.getElementById("rightbutton1").style.display="none"
										document.getElementById("rightbutton2").style.display="none"
									}
									</script>
								</script>

								</form>










										</table>

									</td>
								<!-- right ends -->
								</tr>
		
							</table>
						</td>
					</tr>
		 
				


<!-- bottom starts -->
					<TR bgcolor='#FFFFFF'>
						<td>
							<table name=bottomMain cellspacing=5 cellpadding=0 border=0 height="100%" width="100%">
								
							
							<!-- object 2 -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<!-- actual  object code -->
					
											<table name=obj2 cellspacing=0 cellpadding=0 border=0 height="100%" width="100%">
												<TR bgcolor='#FFFFFF'>
													<td valign="top">
														<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?ar=4&tid=<%=Template_Id%>" style="color:#000000">Bottom</a></font>
													</td>
												</tr>
											  <TR bgcolor='#FFFFFF'> 
												   <td valign="top">
														<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
															<a href="manage_object.asp?ar=4&tid=<%=Template_Id%>" style="color:#000000"> Edit</a> <a href="manage_object.asp?area=4&op=add&tid=<%=Template_Id%>" style="color:#000000">Add</a></font>
													</td>
											  </tr>

											</table>

										<!-- actual object code -->
									</td>
								</tr>
							<!-- object 2  ends-->

							<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">
										<form name="frm_bottomorder" method="post" action="layout_settings.asp">

										<div id=bottombutton1 style="display:block">
											  <img src="images/up.gif" alt="<%=UpAlt_Text%>" onclick="up('bottom',curr_selected['bottom'])" style="cursor:pointer;cursor:hand">
											 <img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('bottom',curr_selected['bottom'])" style="cursor:pointer;cursor:hand">
											 <img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"  onclick="document.frm_bottomorder.Left_Order.value=glob_arrbottom;document.forms['frm_bottomorder'].submit()" style="cursor:pointer;cursor:hand">

											<input type="hidden" name="Template_Id" id="Template_Id" value="<%=Template_Id%>">
											<input type="hidden" name="Left_Order" id="Left_Order" value="">
										</div>

									</td>
								</tr>
							<!-- object 2  ends-->



									<%

									sql_obj="select * from store_design_objects where design_area=4 and store_id="&Store_Id&" and template_id="&Template_Id&" order by view_order"
									set rs_obj=conn_store.execute(sql_obj)
									objcount=-1
									sameLine
									while not (rs_obj.EOF)
									objcount=objcount+1	
									
									%>

									<!-- object start -->
										<script language="javascript"> 
											glob_arrbottom[<%=objcount%>]=<%=rs_Obj("Object_Id")%>;
										</script>
							
							
							<!-- object 2 starts -->
								<% if objcount<>0 and sameLine=1 then %>
							   <td valign="top">
                                                        <% else %>
								<TR bgcolor='#FFFFFF'><td valign="top"><table border=0 cellspacing=0 cellpadding=0><TR bgcolor='#FFFFFF'><td valign="top">
							<% end if %>
									
									<!-- div to show selection and move operation -->
								<div id="bottom<%=objcount%>"  style="background-color:#ffffff" onclick="sel('bottom',<%=objcount%>)">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
										  <TR bgcolor='#FFFFFF'> 
											<td valign="top">
											
											<% if rs_Obj("Object_Type")="Image" then%>
												<img src="<% =Site_Name&"images/"%><%=rs_Obj("Image_Path")%>" alt="<%=rs_Obj("Image_Alt")%>" border="0"/>

											<% elseif rs_Obj("Object_Type")="HTML Text" then%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=fn_replacePath(rs_Obj("HTML_Text"))%></font>
											<% else%>	
												<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_Obj("Object_Type")%></font>
											<% end if%>

											</td>
										  </tr>
										  <TR> 
											  <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="manage_object.asp?Id=<%=rs_Obj("Object_Id")%>&op=edit&tid=<%=Template_Id%>" style="color:#000000"> 
											  Edit</a> <a href="javascript:VerifyDelete('<%=rs_Obj("Object_Id")%>');"  style="color:#000000" > 
											  Delete</a> <a href="javascript:onclick=document.getElementById('bottom<%=objcount%>').onclick()" style="color:#000000">Select</a></font></td>
										  </tr>
										</table>
									
									</div>
									<!-- div ends -->



									<% sameLine=1
                                				if rs_Obj("same_line") = 0 then
                                				   sameLine=0 %>
                                                                    </td></tr></table></td></tr>
                                                                <% else %>
                                                                    </td>
                                				<% end if %>

								<!-- object 2  ends-->
					

									<!-- object end -->
									<% 
									
									rs_obj.movenext
									Wend

									if sameLine<>0 then %>

                                                                           </tr></table></td></tr>
                                                                        <% end if %>

								<!-- object 2 starts -->
								<TR bgcolor='#FFFFFF'>
									<td valign="top">


										<div id=bottombutton2 style="display:block">
										 <img src="images/up.gif"  alt="<%=UpAlt_Text%>" onclick="up('bottom',curr_selected['bottom'])" style="cursor:pointer;cursor:hand">
										<img src="images/down.gif"  alt="<%=DnAlt_Text%>" onclick="dn('bottom',curr_selected['bottom'])" style="cursor:pointer;cursor:hand">
										<img src="images/savepos.gif" alt="<%=SaveAlt_Text%>"  onclick="document.frm_bottomorder.Left_Order.value=glob_arrbottom;document.forms['frm_bottomorder'].submit()" style="cursor:pointer;cursor:hand">

									</div>
																			
									</td>
								</tr>
								<!-- object 2  ends-->

								<script language="javascript">
									if (glob_arrbottom.length <=0)
									{
										document.getElementById("bottombutton1").style.display="none"
										document.getElementById("bottombutton2").style.display="none"
									}
								</script>
								</form>

							
							</table>

						</td>
					</tr>
				<!-- bottom  ends -->
				</table>
			   </td></tr>
				<% else %>
				<table><TR bgcolor='#FFFFFF'><td>The currently selected template does not have the objects necessary to use the new template design interface (ie, you are probably using an older template or a custom designed template).  If you wish to use the new design interface please choose the add new template button above.</td></tr></table>
				<% end if %>


		<%
			createFoot thisRedirect, 0
end if	 ' template id main querstring check 


function fn_replacePath (sText)
    
         sText = replace(sText,"src=""images/images_"&Store_Id&"/","src="""&Site_Name&"images/")
         sText = replace(sText,"src=""images/","src="""&Site_Name&"images/")
         fn_replacePath=sText
end function
%>

