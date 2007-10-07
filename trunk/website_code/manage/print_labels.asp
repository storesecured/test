<!--#include file="Global_Settings.asp"-->

<html><head><style>td {COLOR: 000000;TEXT-DECORATION: none;font-size : 8pt;font-family :Verdana;}
td.to {COLOR: 000000;TEXT-DECORATION: none;font-size : 9pt;font-family :Verdana;}</style></head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<%

EditOrders = request.querystring("Orders")
if EditOrders = "" then
   fn_error "You must select at least one order"
else
StartCell = request.querystring("Start_Cell")

sql_select = "select * from Store_Shipping_Labels WHERE store_id = "&Store_Id
rs_Store.open sql_select,conn_store,1,1
  Page_Width = rs_Store("Page_Width")
  Page_Height = rs_Store("Page_Height")
  Address_Width = rs_Store("Address_Width")
  Address_Height = rs_Store("Address_Height")
  Spacing_Width = rs_Store("Spacing_Width")
  Spacing_Height = rs_Store("Spacing_Height")
  Top_Margin = rs_Store("Top_Margin")
  Bottom_Margin = rs_Store("Bottom_Margin")
  Left_Margin = rs_Store("Left_Margin")
  Right_Margin = rs_Store("Right_Margin")
  Rows = rs_Store("Total_Rows")
  Cols = rs_Store("Total_Cols")
  Image = rs_Store("Image_Name")
  Image_Pos = rs_Store("Image_Pos")

rs_Store.close

Image = "<img src='"&Site_Name&"images/"&Image&"'>"

if not(isNumeric(StartCell)) then
	StartCell = 0	
elseif cint(StartCell) > cint(Rows * Cols) then
   StartCell = 0
else
   StartCell = cint(StartCell)
end if


iConversion = 90
iPageWidth = Page_Width * iConversion
iPageHeight = Page_Height * iConversion
iAddressWidth = Address_Width * iConversion
iAddressHeight = Address_Height * iConversion
iSpacingWidth = Spacing_Width * iConversion
iSpacingHeight = Spacing_Height * iConversion

iTopMargin = Top_Margin * iConversion
iBottomMargin = Bottom_Margin * iConversion
iLeftMargin = Left_Margin * iConversion
iRightMargin = Right_Margin * iConversion
iRows = Rows
iCols = Cols

num_rows=0
iCell = 1



From_Address = Store_Name&"<BR>"&Store_Address1
if Store_Address2 <> "" then
  From_Address = From_Address & "<BR>" & Store_Address2
end if
From_Address = From_Address & "<BR>" & Store_City & ", " & Store_State & " " & Store_Zip & "<BR>" & Store_Country

iTotalColumns = iCols + iCols + 1

sql_select = "select * from Store_purchases WHERE (OID in("&EditOrders&")) AND store_id = "&Store_Id
rs_Store.open sql_select,conn_store,1,1
if rs_Store.eof and rs_Store.bof then
   response.write "Could not find any records to create labels for."
else

  response.write "<table width="&iPageWidth - iLeftMargin - iRightMargin&" border=0 cellspacing=0 cellpadding=0 valign=top>"
  Do while not rs_Store.eof
     num_cols=0
     if num_rows = 0 then
        response.write "<tr><td><img src=images/spacer.gif height="&iTopMargin&" width="&iLeftMargin&"></td>"
     elseif num_rows = iRows+1 then
        num_rows = 0
        response.write "<tr><td><img src=images/spacer.gif height="&iTopMargin&" width="&iLeftMargin&"></td>"
     else
        response.write "<tr><td><img src=images/spacer.gif height="&iSpacingHeight&" width="&iLeftMargin&"></td>"
     end if

     Do while  num_cols<> iCols and not rs_store.eof
        if num_rows = 0 then
           response.write "<td><img src=images/spacer.gif width="&iAddressWidth&" height=1></td>"
        else
           if iCell >= StartCell then
             To_Address = rs_Store("ShipFirstName") & " " & rs_store("ShipLastName")
             if rs_Store("ShipCompany") <> "" then
               To_Address = To_Address & "<BR>"& rs_Store("ShipCompany")
             end if
             To_Address = To_Address & "<BR>"& rs_Store("ShipAddress1")
             if rs_Store("ShipAddress2") <> "" then
               To_Address = To_Address & "<BR>" & rs_Store("ShipAddress2")
             end if
             To_Address = To_Address &"<BR>"&rs_Store("ShipCity") & ", "&rs_Store("ShipState") & " " & rs_Store("ShipZip")
             if Store_Country<>rs_Store("ShipCountry") then
                To_Address=To_Address&"<BR>"&rs_Store("ShipCountry")
             end if
             response.write "<td height="&iAddressHeight&" width="&iAddressWidth&"><table height="&iAddressHeight&" width="&iAddressWidth&">"
             if Image_Pos = 1 then
               response.write "<tr><td valign=top colspan=2>"&Image&"</td>"
             else
               response.write "<tr><td valign=top><B>From</b></td><td valign=top>"&From_Address&"</td>"
             end if
             if Image_Pos = 2 then
               response.write "<td align=right valign=top>"&Image&"</td>"
             else
               response.write "<td align=right valign=top>Order ID#"&rs_store("oid")&"</td>"
             end if
             if Image_Pos = 3 then
               response.write "<tr><td></td><td valign=top colspan=2>"&Image&"</td></tr></table></td>"
             else
               'response.write "<tr><td></td><td valign=top align=right><b>Ship To</b></td><td valign=top class=to>"&To_Address&"</td></tr></table></td>"
               response.write "<tr><td colspan=3><table align=right><tr><td valign=top align=right><b>Ship To</b></td><td valign=top class=to>"&To_Address&"</td></tr></table></td></tr></table></td>"

             end if
             rs_store.movenext
           else
             response.write "<td></td>"
           end if
           iCell = iCell + 1
        end if

        if num_cols <> iCols and num_rows <> 0 then
           response.write "<td><img src=images/spacer.gif width="&iSpacingWidth&" height="&iAddressHeight&"></td>"
        elseif num_rows = 0 then
           response.write "<td><img src=images/spacer.gif width="&iSpacingWidth&" height=1></td>"
        else
           response.write "<td><img src=images/spacer.gif width="&iRightMargin&" height=1></td></tr>"
        end if
        num_cols = num_cols + 1

     Loop
     'movenext would be here
     if num_rows = iRows then
        response.write "<tr><td colspan="&iTotalColumns&"><img src=images/spacer.gif width=1 height="&iBottomMargin&"></td></tr>"
     else
        response.write "<tr><td colspan="&iTotalColumns&"><img src=images/spacer.gif width=1 height="&iSpacingHeight&"></tr></tr>"
     end if
     num_rows = num_rows + 1

  Loop

  response.write "</table>"

end if
rs_Store.close
end if
%>

