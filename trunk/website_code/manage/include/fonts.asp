
<%
function font_selector(sSelected,sName,sType) %>
<select name=<%= sName %> onChange="changeFont_<%= sName %>()">
<option value='<%= sSelected %>' selected><%= sSelected %></option>
<option value='Arial'>Arial</option>
<option value='Arial Black'>Arial Black</option>
<option value='Courier New'>Courier New</option>
<option value='Impact'>Impact</option>
<option value='Tahoma'>Tahoma</option>
<option value='Times New Roman'>Times New Roman</option>
<option value='Verdana'>Verdana</option>
</select><br>
<% sSelected = Replace(sSelected," ","_")
sSelected = Replace(sSelected,",","")
sSelected = Replace(sSelected,".","") %>
<img src="images/spacer.gif" name=font_thumbnail_<%= sName %>>
<script language="Javascript"/>
function changeFont_<%= sName %>() {
           var font = document.EditImage.<%= sName %>.options[document.EditImage.<%= sName %>.selectedIndex].value
           font = fontImage_<%= sName %>(font)
           document.font_thumbnail_<%= sName %>.src = "images/fonts/"+font+".jpg";
}
function fontImage_<%= sName %>(sFont) {
           var font = document.EditImage.<%= sName %>.options[document.EditImage.<%= sName %>.selectedIndex].value
           var arrTmp = font.split(" ");
           if (arrTmp.length > 1) font = arrTmp.join("_");
           var arrTmp = font.split(",");
           if (arrTmp.length > 1) font = arrTmp.join("");
           var arrTmp = font.split(".");
           return font;
}
</script>
<% end function %>
