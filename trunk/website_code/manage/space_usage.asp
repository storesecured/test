<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->

<%

dim sizeUsed,availableMB,total_size

sFormName = "space"
sTitle = "Space Usage"
sFullTitle = "My Account > Space Usage"
thisRedirect = "space_usage.asp"
sMenu = "account"
CreateHead thisRedirect

dim usize

if Service_Type = 0 then
	available = 5000
elseif Service_Type = 3 then
	available = 50000
elseif Service_Type = 5 then
	available = 100000
elseif Service_Type = 7 then
	available = 250000
elseif Service_Type = 9 or Service_Type = 10 then
	available = 500000
else
	available = 10000000
end if


if Additional_Storage > 0 then
	available = available + (100000 * Additional_Storage)
end if

usize = SizeUsage()
total_size = total_size/1000
available = available/1000

 
%>
	<TR bgcolor='#FFFFFF'>
		 <td width="40%" height="23" class="inputname"><B>Total Space</B></td>
		 <td width="60%" height="23" class="inputvalue"><%=formatNumber(available,2,-1,0,0)%> MB</td>
	</tr>
	<TR bgcolor='#FFFFFF'>
		 <td width="40%" height="23" class="inputname"><B>Used Space</B></td>
		 <td width="60%" height="23" class="inputvalue"><%=formatNumber(total_size,2,-1,0,0)%> MB</td>
	</tr>
	
	<TR bgcolor='#FFFFFF'>
		 <td width="40%" height="23" class="inputname"><B>Used Space Percentage</B></td>
		 <td width="60%" height="23" class="inputvalue"><%=usize%>%</td>
	</tr>
	 

<% createFoot thisRedirect,0 %>

