<%
sNeedEditor=1
function create_editor_button (sTextFieldName,sStartValue,sCustomTags)
	if show_editor_cookie=2 then %>
	       <input class=buttons type="button" value="Use Popup Editor" name="Editor" OnClick="javascript:goHtmlEditor('<%= sTextFieldName %>')">
        <% else%>
	       <input class=buttons type="button" value="Use Popup Editor" name="Editor" OnClick="javascript:goNewHtmlEditor('<%= sTextFieldName %>')">

        <%
	end if
        response.write "<textarea rows='12' name='"&sTextFieldName&"' cols='83' id='"&sTextFieldName&"'>"&sStartValue&"</textarea>"

end function

function create_editor_full_button (sTextFieldName,sStartValue,sCustomTags)
	if show_editor_cookie=2 then %>
	       <input class=buttons type="button" value="Use Popup Editor" name="Editor" OnClick="javascript:goHtmlEditor('<%= sTextFieldName %>')">
        <% else%>
	       <input class=buttons type="button" value="Use Popup Editor" name="Editor" OnClick="javascript:goNewHtmlEditor_full('<%= sTextFieldName %>')">

        <%
	end if
        response.write "<textarea rows='12' name='"&sTextFieldName&"' cols='83' id='"&sTextFieldName&"'>"&sStartValue&"</textarea>"

end function

function create_editor (sTextFieldName,sStartValue,sCustomTags)
	if show_editor_cookie=2 then %>
	       <input class=buttons type="button" value="Use Popup Editor" name="Editor" OnClick="javascript:goHtmlEditor('<%= sTextFieldName %>')">
        <% elseif show_editor_cookie=0 then %>
	       <input class=buttons type="button" value="Use Popup Editor" name="Editor" OnClick="javascript:goNewHtmlEditor('<%= sTextFieldName %>')">

        <%
	end if
        response.write "<textarea rows='12' name='"&sTextFieldName&"' cols='83' id='"&sTextFieldName&"'>"&sStartValue&"</textarea>"

        if show_editor_cookie=1 then
        %>
	<script>
			var oEdit<%= sTextFieldName %> = new InnovaEditor("oEdit<%= sTextFieldName %>");
			oEdit<%= sTextFieldName %>.width="100%";//You can also use %, for example: oEdit<%= sTextFieldName %>.width="100%"
			oEdit<%= sTextFieldName %>.height=300;

			oEdit<%= sTextFieldName %>.features=["FullScreen","|","Search","Cut","Copy","Paste","PasteWord","|","Undo","Redo","|",
					  "ForeColor","BackColor","|","Hyperlink","InternalLink","Bookmark","|",
					  "Image","Flash","Media","|","Table","Guidelines","Absolute","|","Numbering","Bullets","|",
					  "Characters","Line","BRK",
					  "StyleAndFormatting","|","TextFormatting","ParagraphFormatting","ListFormatting","BoxFormatting","CssTest",
                                          "FontName","FontSize","CustomTag","|",
					  "Bold","Italic","Underline","|","Strikethrough","Subscript","Superscript","|",
					  "Indent","Outdent","|","JustifyLeft","JustifyCenter","JustifyRight","JustifyFull","|","RemoveFormat","ClearAll","HTMLSource"];

			oEdit<%= sTextFieldName %>.arrStyle = [["BODY",false,"","background:#FFFFFF;font-family:Verdana,Arial,Helvetica;font-size:x-small;"],
						[".ScreenText",true,"Screen Text","font-family:Tahoma;"],
						[".ImportantWords",true,"Important Words","font-weight:bold;"],
						[".Highlight",true,"Highlight","font-family:Arial;color:red;"]];

			oEdit<%= sTextFieldName %>.cmdAssetManager = "modalDialogShow('/Editor/assetmanager/assetmanager.asp',640,465)"; //Command to open the Asset Manager add-on.
  			oEdit<%= sTextFieldName %>.cmdInternalLink = "window.open('editor_internal.asp','mywindow','resizable=0,width=450,height=400')"; //Command to open your custom link lookup page.

			<% if sCustomTags="" then
				sCustomTags="[""None Available"",""""]"
			end if %>
			oEdit<%= sTextFieldName %>.arrCustomTag=[<%= sCustomTags %>];//Define custom tag selection
			oEdit<%= sTextFieldName %>.customColors=["#ff4500","#ffa500","#808000","#4682b4","#1e90ff","#9400d3","#ff1493","#a9a9a9"];//predefined custom colors
			oEdit<%= sTextFieldName %>.publishingPath="<%= Site_Name %>";
			oEdit<%= sTextFieldName %>.REPLACE("<%= sTextFieldName %>");


		    if (oUtil.setEdit) oUtil.setEdit("oEdit<%= sTextFieldName %>");


		</script>
	<%
	end if
end function %>