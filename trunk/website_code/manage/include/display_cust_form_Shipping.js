function validateState1()
 {
 	if(window.document.getElementById("Country1").value=="United States")
 	    {
 	        window.document.getElementById("UnitedStatesDiv1").style.display="block"
 	        window.document.getElementById("CanadaDiv1").style.display="none"
 	        window.document.getElementById("optDiv1").style.display="none" 	       
 	        window.document.getElementById("State_UnitedStates1").options[1].selected=true
 	        window.document.getElementById("State1").value = window.document.getElementById("State_UnitedStates1").value
 	    }
 	else if(window.document.getElementById("Country1").value=="Canada")
 	    {
 	        window.document.getElementById("UnitedStatesDiv1").style.display="none"
 	        window.document.getElementById("CanadaDiv1").style.display="block"
 	        window.document.getElementById("optDiv1").style.display="none" 
 	        window.document.getElementById("State_Canada1").options[1].selected=true
 	        window.document.getElementById("State1").value = window.document.getElementById("State_Canada1").value	    	    
 	    }
 	else
 	   {
 	        window.document.getElementById("UnitedStatesDiv1").style.display="none"
 	        window.document.getElementById("CanadaDiv1").style.display="none"
 	        window.document.getElementById("optDiv1").style.display="block"
 	        window.document.getElementById("txtStore_State1").value=""	    
 	    }
 	    
 }	
 
 function Validations1()
 	{
    var str="";
    if((((window.document.getElementById("Country1").value=="United States") || (window.document.getElementById("Country1").value=="Canada")) && window.document.getElementById("State1").value=="") || window.document.getElementById("First_name").value=="" || window.document.getElementById("Last_name").value=="" || window.document.getElementById("Address1").value=="" || window.document.getElementById("City").value=="" || window.document.getElementById("Zip").value=="" || window.document.getElementById("Phone").value=="" || window.document.getElementById("Email").value=="")
    {
		    if(((window.document.getElementById("Country1").value=="United States") || (window.document.getElementById("Country1").value=="Canada")) && window.document.getElementById("State1").value=="")  
		    {
			    str = str + "Please Enter STATE \n";		    			
		    }
    		
		    if(window.document.getElementById("First_name").value=="")
		    {
                str = str + "Please Enter First Name \n";
		    }
		    if(window.document.getElementById("Last_name").value=="")
		    {
                str = str + "Please Enter Last Name \n";
		    }
		    if(window.document.getElementById("Address1").value=="")
		    {
                str = str + "Please Enter Address1 \n";
		    }
		    if(window.document.getElementById("City").value=="")
		    {
                str = str + "Please Enter City \n";
		    }
		    if(window.document.getElementById("Zip").value=="")
		    {
                str = str + "Please Enter Zip Code \n";
		    }
		    if(window.document.getElementById("Phone").value=="")
		    {
                str = str + "Please Enter Phone \n";
		    }
		    if(window.document.getElementById("Email").value=="")
		    {
                str = str + "Please Enter Email Address \n";
		    }
		
	alert(str);
	return false;
	}
	
	else
	{
		return true;
	}	
}    
 
 
 function storeStateData1()
    {
    if(window.document.getElementById("Country1").value=="United States")
 	    {
            window.document.getElementById("State1").value = window.document.getElementById("State_UnitedStates1").value
        }   
    else if(window.document.getElementById("Country1").value=="Canada")
 	    {
            window.document.getElementById("State1").value = window.document.getElementById("State_Canada1").value	    
        }
    else
        {    
            window.document.getElementById("State1").value = window.document.getElementById("txtStore_State1").value	    
        }
    }    