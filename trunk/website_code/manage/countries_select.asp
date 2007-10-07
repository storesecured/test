<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%
sInstructions="Choose which countries should be shown to your customers for registration/shipment."
sTextHelp="shipping/shipping-countries.doc"

'Select Countries

sFormAction = "countries_select.asp"
sFormName = "Select_countries"
sTitle = "Allowable Countries"
sFullTitle = "Shipping > Allowable Countries"

thisRedirect = "countries_select.asp"
sMenu = "shipping"
sQuestion_Path = "advanced/allowable_countries.htm"
createHead thisRedirect
%>


<SCRIPT LANGUAGE="JavaScript">
var sortitems=1;
var strTemp="";

		function fnValidate(tbox) 
		{

		for(i = 0; i < tbox.options.length; i++) {
		strTemp= strTemp  + tbox.options[i].text + ":";
		}

document.forms[0].hidCountry.value = strTemp;


		var ans = document.getElementById("SltCountries").checked;

  	 if(ans == false)
		 {

				if (eval(tbox.options.length) == 0) 
				{
					alert("Please select country for shipping.");
				}
				else 
				document.forms[0].action="countries_select.asp?mode=S";
				document.forms[0].submit();
		  }		
			else 
				document.forms[0].action="countries_select.asp?mode=A";
				document.forms[0].submit();
	}
	
	 
	 function move(fbox,tbox)
   {

      for(var i=0; i<fbox.options.length; i++)
      {
         if(fbox.options[i].selected && fbox.options[i].value != "")
         {
            var no = new Option();
            no.value = fbox.options[i].value;
            no.text = fbox.options[i].text;
            tbox.options[tbox.options.length] = no;
            fbox.options[i].value = "";
            fbox.options[i].text = "";
         }
      }

      ShiftUp(fbox);

      if (sortitems) SortD(tbox);
   }

   function ShiftUp(box)
   {

      for(var i=0; i<box.options.length; i++)
      {
         if(box.options[i].value == "")
         {
            for(var j=i; j<box.options.length-1; j++)
            {
               box.options[j].value = box.options[j+1].value;
               box.options[j].text = box.options[j+1].text;
            }

            var ln = i;
            break;
         }
      }

      if(ln < box.options.length)
      {
         box.options.length -= 1;
         ShiftUp(box);
      }
   }

   function SortD(box)
   {
      var temp_opts = new Array();
      var temp = new Object();
      for(var i=0; i<box.options.length; i++)
      {
         temp_opts[i] = box.options[i];
      }

      for(var x=0; x<temp_opts.length-1; x++)
      {
         for(var y=(x+1); y<temp_opts.length; y++)
         {
            if(temp_opts[x].text > temp_opts[y].text)
            {
               temp = temp_opts[x].text;
               temp_opts[x].text = temp_opts[y].text;
               temp_opts[y].text = temp;
            }
         }
      }

      for(var i=0; i<box.options.length; i++)
      {
         box.options[i].value = temp_opts[i].value;
         box.options[i].text = temp_opts[i].text;
      }
   }



</script>


<%

Function fixQuotes(sTxt)
	fixQuotes = Replace(sTxt,"""","&#034;")
	fixQuotes = Replace(fixQuotes,"'","&#039;")
End Function



sMode = request("mode")


	If sMode = "S" then 	

StrCountry = nullifyQ(request("hidCountry"))
sql_country=replace(strcountry,":","','")
StrCountryCode=""
sql_select="select country_id from sys_countries where country in ('"&sql_country&"') order by country_id"

set myfields=server.createobject("scripting.dictionary")
Call DataGetrows(conn_store,sql_select,mydata,myfields,noRecords)
if noRecords = 0 then
        FOR rowcounter= 0 TO myfields("rowcount")
                if StrCountryCode="" then
                        StrCountryCode=mydata(myfields("country_id"),rowcounter)
                else
	               StrCountryCode=StrCountryCode&","&mydata(myfields("country_id"),rowcounter)
	        end if
        next
end if



		sql_update = "exec wsp_settings_countries " & store_id & ",'" & StrCountryCode & "'"
		conn_store.execute sql_update
        response.redirect "countries_select.asp"
	End if

	If sMode = "A" then
		sql_update = "exec wsp_settings_countries "  & store_id & ",'ALL';"
		conn_store.execute sql_update


        response.redirect "countries_select.asp"
	End if

	SelctedCountries=Countries_Selected	

%>


            			<TR bgcolor='#FFFFFF'>
				<td align="left" class="inputname" colspan=4>
					<input class="image" type="radio" value="All" id="SltCountries"   name="SltCountries"

								<% if SelctedCountries = "ALL" then %>
									checked
								<% end if %>
						>All Countries
						
			     </td>

			</tr>



				<TR bgcolor='#FFFFFF'>
					<td valign="top" nowrap><input class="image" type="radio" value="part" id="SltCountries" name="SltCountries"
						
								<% if SelctedCountries <> "ALL" then %>
									checked
								<% end if %>						
						>Selected Countries
						</td><td><b>Select countries</b><br>
						<select multiple size="20" name="list1" style="width:200">
				<% 

                                sql_country = "SELECT Country,Country_id FROM Sys_Countries where country_code is not null ORDER BY Country;"

				set myfields=server.createobject("scripting.dictionary")
				Call DataGetrows(conn_store,sql_country,mydata,myfields,noRecords)

				if noRecords = 0 then
					FOR rowcounter= 0 TO myfields("rowcount")
					sCountryId=mydata(myfields("country_id"),rowcounter)
					sCountry=mydata(myfields("country"),rowcounter)
					if Not(Is_In_Collection(SelctedCountries,sCountryId,",")) then
                                                response.write "<Option value='"&sCountry&"'>"&sCountry&"</option>"
					end if
                                        Next
				End If %>
			</select>
					</td>
		  
					<td width="16%" align="center">
		  			<input type="button" onClick="move(list2,list1)" value="<<">
						<input type="button" onClick="move(list1,list2)" value=">>">
					</td>


		  
					<td width="42%" valign="top" align="center"><b>Selected countries</b><br>


						<select multiple size="20" name="list2" style="width:200">

									<%
									if SelctedCountries <> "ALL" then

									if noRecords = 0 then
                                        					FOR rowcounter= 0 TO myfields("rowcount")
                                        					sCountryId=mydata(myfields("country_id"),rowcounter)
                                        					sCountry=mydata(myfields("country"),rowcounter)
                                        					if Is_In_Collection(SelctedCountries,sCountryId,",") then
                                                                                        response.write "<Option value='"&sCountry&"'>"&sCountry&"</option>"
                                        					end if
                                                                                Next
                                        				End If 
                                                                        
                                                                        End if %>

				    </select>

							</td>
				</tr>
<input type="hidden" name = "hidCountry" id = "hidCountry">

				<TR bgcolor='#FFFFFF'>
					<td colspan=4 align=center>
						<input type="button" onClick="fnValidate(list2)" value="Save">
					</td>
				</tr>


<% createFoot thisRedirect, 0%>
