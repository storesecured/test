<%
if Start_Row= "" then
	Start_Row=1
end if
if End_Row= "" then
	End_Row=1
end if
%>
 <script language="JavaScript" type="text/javascript">
	function goReload(formName, controlType, controlName, itemID, theUrl,fName){



			for (i=0;i<document.forms(formName)(controlName).length;i++)
				if (controlType=="radio")
            {
            if (document.forms(formName)(controlName)(i).checked)
					selVal = document.forms(formName)(controlName)(i).value;
				}
				else
				{
				   if (document.forms(formName)(controlName)(i).selected)
					selVal = document.forms(formName)(controlName)(i).value;
				}
		if (theUrl != '') {
         theUrl = theUrl + "&";
      }
		theUrl = fName+"?"+theUrl+controlName+"="+selVal+"#ITEM_"+itemID;
		self.location = theUrl;}

	function goReloadAttr(formName, controlType, controlName, itemID, theUrl,fName){

		if(controlType=="dropdown")
		{
			for (i=0;i<document.forms(formName)(controlName).length;i++)
				if (document.forms(formName)(controlName)(i).selected)
					selVal = document.forms(formName)(controlName)(i).value;
		}
		else if(controlType=="radio")
		{		
			for (i=0;i<document.forms(formName)(controlName).length;i++)
				if (document.forms(formName)(controlName)(i).checked)
					selVal = document.forms(formName)(controlName)(i).value;
		}


		selVals = '';
		for (i=0;i<selVal.length;i++)
			if (selVal.charAt(i)=='|')
				i = selVal.length;
			else
				selVals = selVals + selVal.charAt(i);

      if (theUrl != '') {
         theUrl = theUrl + "&";
      }
		theUrl = fName+"?"+theUrl+"repost=1&start="+<%= Start_Row_Items %>;

		if (document.forms(formName)("u_d_1")) {
			u_d_1 = document.forms(formName)("u_d_1").value;
			if (u_d_1 != "")
				theUrl = theUrl+"&u_d_1_"+itemID+"="+u_d_1;
		}
		if (document.forms(formName)("u_d_2")) {
			u_d_2 = document.forms(formName)("u_d_2").value;
			if (u_d_2 != "")
				theUrl = theUrl+"&u_d_2_"+itemID+"="+u_d_2;
		}
		if (document.forms(formName)("u_d_3")) {
			u_d_3 = document.forms(formName)("u_d_3").value;
			if (u_d_3 != "")
				theUrl = theUrl+"&u_d_3_"+itemID+"="+u_d_3;
		}
		if (document.forms(formName)("u_d_4")) {
			u_d_4 = document.forms(formName)("u_d_4").value;
			if (u_d_4 != "")
				theUrl = theUrl+"&u_d_4_"+itemID+"="+u_d_4;
		}
		if (document.forms(formName)("u_d_5")) {
			u_d_5 = document.forms(formName)("u_d_5").value;
			if (u_d_5 != "")
				theUrl = theUrl+"&u_d_5_"+itemID+"="+u_d_5;
		}

		theUrl = theUrl+"&"+controlName+"_"+itemID+"="+selVals+"#ITEM_"+itemID
		self.location = theUrl;}
</script>
