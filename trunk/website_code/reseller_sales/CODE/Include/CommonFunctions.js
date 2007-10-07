//FUNCTIONS

var whitespace =" \t\n\r ";

// function for checking empty filed

function isEmpty(str)
{  
	return ((str == null) || (str.length == 0))
}

// end of function
	
// function to check the whitespace

function isWhitespace(str)
{    var i;
	 var flag

 	  // Is s empty?
 	  if (isEmpty(str)) return true;		
 	   // Search through string's characters one by one
 	   // until we find a non-whitespace character.
 	   // When we do, return false; if we don't, return true.
 	   for (i = 0; i < str.length; i++)
 	   {   
 	       // Check that current character isn't whitespace.
 	       var c = str.charAt(i);

		   if (whitespace.indexOf(c) == -1)
		   		return false
 	   }	
 	   // All characters are whitespace.
		    return true;
}

// end of function



// function to check all the entered values are characters only

function isAllCharacters(objValue)
{
		var characters="' -abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ./"
		var tmp
		var lTag
		lTag = 0
		temp = (objValue.length)
		for (var i=0;i<temp;i++)
		{
			tmp=objValue.substring(i,i+1)
			if (characters.indexOf(tmp)==-1)
			{
				lTag = 1
			}
		}
		if(lTag == 1)
			return false
		else
			return true
}
	
// end of function


// function for triming

function trim(pstrString)
{
	var intLoop=0;

	for(intLoop=0; intLoop<pstrString.length; )
	{
		if(pstrString.charAt(intLoop)==" ")
			pstrString=pstrString.substring(intLoop+1, pstrString.length);
		else
			break;
	}

	for(intLoop=pstrString.length-1; intLoop>=0; intLoop=pstrString.length-1)
	{
		if(pstrString.charAt(intLoop)==" ")
			pstrString=pstrString.substring(0,intLoop);
		else
			break;
	}
	return pstrString;
}

// function to check all the entered values are characters only

function isAlphaNumeric(objValue)
{	
			var characters="'-abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 "
			var tmp
			var lTag
			lTag = 0
			temp = (objValue.length)
			for (var i=0;i<temp;i++)
			{
				tmp=objValue.substring(i,i+1)
				if (characters.indexOf(tmp)==-1)
				{
					lTag = 1
				}
			}
			if(lTag == 1)
				return false
			else
				return true
	
}
	
// end of function

// function for checking if inputed value is a valid number
// @param objValue - string 

function isAllNumeric(objValue)
{	
			lTempLength = objValue.length
			lTempCounter = 0 
			lTempString = trim(objValue)
			flag = false
			
			do
			{
			if(lTempString.charAt(lTempCounter) == " " || lTempString.lastIndexOf('.')!= lTempString.indexOf('.'))
			{
				flag = false
				break
			}
			else if(lTempString.charAt(lTempCounter) > 0 || lTempString.charAt(lTempCounter) < 9 || lTempString.charAt(lTempCounter) == ".")
				flag = true
			else
				{
					flag = false
					break
				}
				lTempCounter = lTempCounter + 1
			}

			
			while(lTempCounter <= lTempLength)			
			
			if(flag == true)
				return true
			else
				return false
}
		
// end of function		


	
//-----------This is a functn for email validation added by Peupy------------

function IsEmail(InString)
{
   var left,right;
   
   if(InString.length==0) return(false);
   for(Count=0;Count<InString.length;Count++)
   {
       TempChar = InString.substring(Count,Count + 1);
       if("1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.@_-".indexOf(TempChar,0)==-1) return(false); 
   }

   if(InString.indexOf('@')< 1) return(false);//if it has no @ or @ at the beginning
   
   if(InString.lastIndexOf('@')!= InString.indexOf('@')) return(false);// if it has more than one @
   
   left = InString.substring(0,InString.indexOf('@'));
   right = InString.substring(InString.indexOf('@') + 1,InString.length);
   
   if((!isDotExpression(left,0))||(!isDotExpression(right,1))) return(false);
   return(true);

}//end of function



function isDotExpression(InString,NeedsDot)
{
   var dots,index,tmpNeedDot;
   dots=0;
   
   for(index=0;index<InString.length;index++)
   {
      if(InString.substring(index,index+1)==".")
      {
         if((index==0)||(index==InString.length-1)) return(false);
         dots ++;
         
         if(dots>1)tmpNeedDot=1;
         else tmpNeedDot=0;

         if(!isDotExpression(InString.substring(0,index),tmpNeedDot)) return(false);
      }      
   }
   if((NeedsDot==1)&&(dots<1)) return (false);
   if(InString.length < dots * 2+1) return (false);
   return (true);
}

// function to check valid phone no

function isPhone(objValue)
{	
			var characters="-+0123456789() "
			var tmp
			var lTag
			lTag = 0
			temp = (objValue.length)
			for (var i=0;i<temp;i++)
			{
				tmp=objValue.substring(i,i+1)
				if (characters.indexOf(tmp)==-1)
				{
					lTag = 1
				}
			}
			if(lTag == 1)
				return false
			else
				return true
	
}
	
// end of function
	//-----------------------------------------------------