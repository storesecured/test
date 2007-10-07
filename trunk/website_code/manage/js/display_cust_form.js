function validateState()
 {
 	if(window.document.getElementById("Country").value=="United States")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="block"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="none" 	       
 	        window.document.getElementById("State_UnitedStates").options[0].selected=true
 	        window.document.getElementById("State").value = window.document.getElementById("State_UnitedStates").value
 	    }
 	else if(window.document.getElementById("Country").value=="Canada")
 	    {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="block"
 	        window.document.getElementById("optDiv").style.display="none" 
 	        window.document.getElementById("State_Canada").options[0].selected=true
 	        window.document.getElementById("State").value = window.document.getElementById("State_Canada").value	    	    
 	    }
 	else
 	   {
 	        window.document.getElementById("UnitedStatesDiv").style.display="none"
 	        window.document.getElementById("CanadaDiv").style.display="none"
 	        window.document.getElementById("optDiv").style.display="block"
 	        window.document.getElementById("txtStore_State").value=""
			window.document.getElementById("State").value=""
 	    }
 	    
 }	
 
 function Validations()
 	{
    var My_str="";
    if((((window.document.getElementById("Country").value=="United States") || (window.document.getElementById("Country").value=="Canada")) && window.document.getElementById("State").value=="") || window.document.getElementById("First_Name").value=="" || window.document.getElementById("Last_Name").value=="" || window.document.getElementById("Address1").value=="" || window.document.getElementById("City").value=="" || window.document.getElementById("Zip").value=="" || window.document.getElementById("Phone").value=="" || window.document.getElementById("Email").value=="" || window.document.getElementById("Country").value=="")
    {
		    if(((window.document.getElementById("Country").value=="United States") || (window.document.getElementById("Country").value=="Canada")) && window.document.getElementById("State").value=="")  
		    {
			    My_str = My_str + "Please Enter STATE \n";		    			
		    }
    		
		    if(window.document.getElementById("First_Name").value=="")
		    {
                My_str = My_str + "Please Enter First Name \n";
		    }
		    if(window.document.getElementById("Last_Name").value=="")
		    {
                My_str = My_str + "Please Enter Last Name \n";
		    }
		    if(window.document.getElementById("Address1").value=="")
		    {
                My_str = My_str + "Please Enter Address1 \n";
		    }
		    if(window.document.getElementById("City").value=="")
		    {
                My_str = My_str + "Please Enter City \n";
		    }
		    if(window.document.getElementById("Zip").value=="")
		    {
                My_str = My_str + "Please Enter Zip Code \n";
		    }
		    if(window.document.getElementById("Phone").value=="")
		    {
                My_str = My_str + "Please Enter Phone \n";
		    }
		    if(window.document.getElementById("Email").value=="")
		    {
                My_str = My_str + "Please Enter Email Address \n";
		    }
			if(window.document.getElementById("Country").value == "")
			{
				My_str = My_str + "Please Select Country \n";
			}
		
	alert(My_str);
	return false;
	}
	
	else
	{
		return true;
	}	
}  
   
 function storeStateData()
    {
    if(window.document.getElementById("Country").value=="United States")
 	    {
            window.document.getElementById("State").value = window.document.getElementById("State_UnitedStates").value
        }   
    else if(window.document.getElementById("Country").value=="Canada")
 	    {
            window.document.getElementById("State").value = window.document.getElementById("State_Canada").value	    
        }
    else
        {    
            window.document.getElementById("State").value = window.document.getElementById("txtStore_State").value	    
        }
    }    	    
 	