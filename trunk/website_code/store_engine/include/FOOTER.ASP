<script language="JavaScript" type="text/javascript">
var submitcount = 0;
function submitForm()
{
<%
if sSubmitName<>"" then %>
if (submitcount == 0) {
	//sumbit form 
	<%
  FormArray=split(sFormName,",")
  SubmitArray=split(sSubmitName,",")
	For i=0 to ubound(FormArray) %>
  	try
  	{
        document.forms["<%= FormArray(i) %>"].<%= SubmitArray(i) %>.value="Processing...";
        document.forms["<%= FormArray(i) %>"].<%= SubmitArray(i) %>.disabled=true;
    }
    catch(e){}
  <% Next %>
	submitcount ++;

	return true;
	}
else
	{
		alert("The transaction has already been submitted please wait.\n\nIf you are receiving this error after using the back button to return to the payment page, please hit F5 to refresh the screen.");
		return false;
	}


<% else %>
   return true;
<% end if %>
}
</script>

<%

response.write Store_Footer
Set rs_Store = Nothing

conn_store.close
set conn_store = nothing

%>


