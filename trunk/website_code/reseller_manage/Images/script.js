	function checkBrowser(){
	this.ver=navigator.appVersion
	this.dom=document.getElementById?1:0
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.ns5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie5 || this.ie4 || this.ns4 || this.ns5)
	return this
}
bw=new checkBrowser()
//With nested layers for netscape, this function hides the layer if it's visible and visa versa
function showHideAuto(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	if(obj.visibility=='visible' || obj.visibility=='show' || obj.visibility=='block') obj.visibility='hidden'
	else obj.visibility='visible'
}
function showHideForm(div,field,nest){
  if (document.forms[0].elements(field).checked)
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	  obj.display='block'
  }
  	else
  {
  	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	  obj.display='none'
  }
}

//Shows the div
function show(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0;
	obj.visibility='visible'
}
//Hides the div
function hide(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	obj.visibility='hidden'
}
/*
  -------------------------------------------------------------------------
	                    JavaScript Form Validator 
                                Version 2.0.2
	Copyright 2003 JavaScript-coder.com. All rights reserved.
	You use this script in your Web pages, provided these opening credit
    lines are kept intact.
	The Form validation script is distributed free from JavaScript-Coder.com

	You may please add a link to JavaScript-Coder.com, 
	making it easy for others to find this script.
	Checkout the Give a link and Get a link page:
	http://www.javascript-coder.com/links/how-to-link.php

    You may not reprint or redistribute this code without permission from 
    JavaScript-Coder.com.
	
	JavaScript Coder
	It precisely codes what you imagine!
	Grab your copy here:
		http://www.javascript-coder.com/
    -------------------------------------------------------------------------  
*/
function Validator(frmname)
{
  this.formobj=document.forms[frmname];
	if(!this.formobj)
	{
	  alert("BUG: couldnot get Form object "+frmname);
		return;
	}
	if(this.formobj.onsubmit)
	{
	 this.formobj.old_onsubmit = this.formobj.onsubmit;
	 this.formobj.onsubmit=null;
	}
	else
	{
	 this.formobj.old_onsubmit = null;
	}
	this.formobj.onsubmit=form_submit_handler;
	this.addValidation = add_validation;
	this.setAddnlValidationFunction=set_addnl_vfunction;
	this.clearAllValidations = clear_all_validations;
}
function set_addnl_vfunction(functionname)
{
  this.formobj.addnlvalidation = functionname;
}
function clear_all_validations()
{
	for(var itr=0;itr < this.formobj.elements.length;itr++)
	{
		this.formobj.elements[itr].validationset = null;
	}
}
function form_submit_handler()
{
	for(var itr=0;itr < this.elements.length;itr++)
	{
		if(this.elements[itr].validationset &&
	   !this.elements[itr].validationset.validate())
		{
		  return false;
		}
	}
	if(this.addnlvalidation)
	{
	  str =" var ret = "+this.addnlvalidation+"()";
	  eval(str);
    if(!ret) return ret;
	}
	if (submitForm()) {
     return true;
  }
  else
  {
     return false;	
  }

}

function add_validation(itemname,descriptor,errstr)
{
  if(!this.formobj)
	{
	  alert("BUG: the form object is not set properly");
		return;
	}//if
	var itemobj = this.formobj[itemname];
  if(!itemobj)
	{
	  alert("BUG: Couldnot get the input object named: "+itemname);
		return;
	}
	if(!itemobj.validationset)
	{
	  itemobj.validationset = new ValidationSet(itemobj);
	}
  itemobj.validationset.add(descriptor,errstr);
}
function ValidationDesc(inputitem,desc,error)
{
  this.desc=desc;
	this.error=error;
	this.itemobj = inputitem;
	this.validate=vdesc_validate;
}
function vdesc_validate()
{
 if(!V2validateData(this.desc,this.itemobj,this.error))
 {
    this.itemobj.focus();
		return false;
 }
 return true;
}
function ValidationSet(inputitem)
{
    this.vSet=new Array();
	this.add= add_validationdesc;
	this.validate= vset_validate;
	this.itemobj = inputitem;
}
function add_validationdesc(desc,error)
{
  this.vSet[this.vSet.length]= 
	  new ValidationDesc(this.itemobj,desc,error);
}
function vset_validate()
{
   for(var itr=0;itr<this.vSet.length;itr++)
	 {
	   if(!this.vSet[itr].validate())
		 {
		   return false;
		 }
	 }
	 return true;
}
function validateEmailv2(email)
{
// a very simple email validation checking. 
// you can add more complex email checking if it helps 
    if(email.length <= 0)
	{
	  return true;
	}
    var splitted = email.match("^(.+)@(.+)$");
    if(splitted == null) return false;
    if(splitted[1] != null )
    {
      var regexp_user=/^\"?[\w-_\.]*\"?$/;
      if(splitted[1].match(regexp_user) == null) return false;
    }
    if(splitted[2] != null)
    {
      var regexp_domain=/^[\w-\.]*\.[A-Za-z]{2,4}$/;
      if(splitted[2].match(regexp_domain) == null) 
      {
	    var regexp_ip =/^\[\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\]$/;
	    if(splitted[2].match(regexp_ip) == null) return false;
      }// if
      return true;
    }
return false;
}
function V2validateData(strValidateStr,objValue,strError) 
{ 
    var epos = strValidateStr.search("="); 
    var  command  = ""; 
    var  cmdvalue = ""; 
    if(epos >= 0) 
    { 
     command  = strValidateStr.substring(0,epos); 
     cmdvalue = strValidateStr.substr(epos+1); 
    } 
    else 
    { 
     command = strValidateStr; 
    } 
    switch(command) 
    { 
        case "req": 
        case "required": 
         { 
           if(eval(objValue.value.length) == 0) 
           { 
              if(!strError || strError.length ==0) 
              { 
                strError = objValue.name + " : Required Field"; 
              }//if 
              alert(strError);
              objValue.focus
              return false; 
           }//if 
           break;             
         }//case required 
        case "maxlength": 
        case "maxlen": 
          { 
             if(eval(objValue.value.length) >  eval(cmdvalue)) 
             { 
               if(!strError || strError.length ==0) 
               { 
                 strError = objValue.name + " : "+cmdvalue+" characters maximum "; 
               }//if 
               alert(strError + "\n[Current length = " + objValue.value.length + " ]");
               objValue.focus
               return false; 
             }//if 
             break; 
          }//case maxlen 
        case "minlength": 
        case "minlen": 
           { 
             if(eval(objValue.value.length) <  eval(cmdvalue)) 
             { 
               if(!strError || strError.length ==0) 
               { 
                 strError = objValue.name + " : " + cmdvalue + " characters minimum  "; 
               }//if               
               alert(strError + "\n[Current length = " + objValue.value.length + " ]");
               objValue.focus
               return false;                 
             }//if 
             break; 
            }//case minlen 
        case "alnum": 
        case "alphanumeric": 
           { 
              var charpos = objValue.value.search("[^A-Za-z0-9]"); 
              if(objValue.value.length > 0 &&  charpos >= 0)
              { 
               if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": Only alpha-numeric characters allowed "; 
                }//if 
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]");
                objValue.focus
                return false; 
              }//if 
              break; 
           }//case alphanumeric 
        case "num": 
        case "numeric": 
           { 
              var charpos = objValue.value.search("[^0-9]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
                if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": Only digits allowed "; 
                }//if               
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]");
                objValue.focus
                return false; 
              }//if 
              break;               
           }//numeric 
        case "alphabetic": 
        case "alpha": 
           { 
              var charpos = objValue.value.search("[^A-Za-z]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
                  if(!strError || strError.length ==0)
                { 
                  strError = objValue.name+": Only alphabetic characters allowed "; 
                }//if                             
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]"); 
                objValue.focus
                return false; 
              }//if 
              break; 
           }//alpha

		case "alnumhyphen":
			{
              var charpos = objValue.value.search("[^A-Za-z0-9\-_]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
                  if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": characters allowed are A-Z,a-z,0-9,- and _"; 
                }//if                             
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]"); 
                objValue.focus
                return false; 
              }//if 			
			break;
			}
        case "email": 
          { 
               if(!validateEmailv2(objValue.value)) 
               { 
                 if(!strError || strError.length ==0) 
                 { 
                    strError = objValue.name+": Enter a valid Email address "; 
                 }//if                                               
                 alert(strError); 
                 objValue.focus
                 return false; 
               }//if 
           break; 
          }//case email 
        case "date":
          { 
               if(!isValidDate(objValue.value))
               { 
                 if(!strError || strError.length ==0) 
                 { 
                    strError = objValue.name+": Enter a valid date";
                 }//if                                               
                 alert(strError); 
                 objValue.focus
                 return false; 
               }//if 
           break; 
          }//case email
        case "lt": 
        case "lessthan": 
         { 
            if(isNaN(objValue.value)) 
            { 
              alert(objValue.name+": Should be a number "); 
              return false; 
            }//if 
            if(eval(objValue.value) >=  eval(cmdvalue))
            { 
              if(!strError || strError.length ==0) 
              { 
                strError = objValue.name + " : value should be less than "+ cmdvalue; 
              }//if               
              alert(strError);
              objValue.focus
              return false;                 
             }//if             
            break; 
         }//case lessthan 
        case "gt": 
        case "greaterthan": 
         { 
            if(isNaN(objValue.value)) 
            { 
              alert(objValue.name+": Should be a number "); 
              return false; 
            }//if 
             if(eval(objValue.value) <=  eval(cmdvalue)) 
             { 
               if(!strError || strError.length ==0) 
               { 
                 strError = objValue.name + " : value should be greater than "+ cmdvalue; 
               }//if               
               alert(strError); 
               objValue.focus
               return false;                 
             }//if             
            break; 
         }//case greaterthan 
        case "regexp": 
         { 
		 	if(objValue.value.length > 0)
			{
	            if(!objValue.value.match(cmdvalue)) 
	            { 
	              if(!strError || strError.length ==0) 
	              { 
	                strError = objValue.name+": Invalid characters found "; 
	              }//if                                                               
	              alert(strError); 
	              objValue.focus
	              return false;                   
	            }//if 
			}
           break; 
         }//case regexp 
        case "dontselect": 
         { 
            if(objValue.selectedIndex == null) 
            { 
              alert("BUG: dontselect command for non-select Item"); 
              return false; 
            } 
            if(objValue.selectedIndex == eval(cmdvalue)) 
            { 
             if(!strError || strError.length ==0) 
              { 
              strError = objValue.name+": Please Select one option "; 
              }//if                                                               
              alert(strError);
              objValue.focus
              return false;                                   
             } 
             break; 
         }//case dontselect 
    }//switch 
    return true; 
}
/*
	Copyright 2003 JavaScript-coder.com. All rights reserved.
*/

function alertUser(){
	 alert("Your session has been inactive for 15 minutes.\nTo protect your personal information all inactive sessions will expire after 20 minutes.\nAny changes you have made will be lost if they are not saved within the next 5 minutes!");
}
function goodchars(e, goods)
{
  var key, keychar;
  key = getkey(e);
  if (key == null) return true;
  // get character
  keychar = String.fromCharCode(key);
  keychar = keychar.toLowerCase();
  goods = goods.toLowerCase();
  // check goodkeys
  if (goods.indexOf(keychar) != -1)
  return true;
  // control keys
  if ( key==null || key==0 || key==8 || key==9 || key==13 || key==27 )
  return true;
  // else return false
  return false;
}
function getkey(e)
{
  if (window.event)
  return window.event.keyCode;
  else if (e)
  return e.which;
  else
  return null;
}
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);
// otherwise, update 'characters left' counter
else 
countfield.value = maxlimit - field.value.length;
}

	function goColorPicker(resCtrl) {
		var w = 230
		var h = 280
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="Color_Picker.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'colorPicker','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,'+winprops);}

	function goImagePicker(resCtrl) {
		var w = 400
		var h = 480
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="Image_Picker.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'imagePicker','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,'+winprops);}

  function goFileUploader(resCtrl) {
		var w = 380
		var h = 170
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="file_uploader.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'fileUploader','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=on,resizable=no,'+winprops);}

	function goFilePicker(resCtrl) {
		var w = 520
		var h = 380
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="File_Picker.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'filePicker','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}

	function goItemPicker(resCtrl) {
		var w = 600;
		var h = 400;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="Item_Picker.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'itemPicker','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,'+winprops);}

	function goShowPicture(picture) {
		var w = 650;
		var h = 480;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="show_image.asp?returnArg="+picture;
		reWin = window.open(theUrl,'picker','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,'+winprops);}

	function goHtmlEditor(resCtrl,sType) {
		var w = 770
		var h = 550
		if (screen.width>770)
			w = 770
		if (screen.height>550)
			h = 550
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="html_editor.asp?returnArg="+resCtrl+"&sType="+sType;
		reWin = window.open(theUrl,'htmlEditor','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,'+winprops);}

	function goSend(resCtrl) {
		var w = 470;
		var h = 275;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="contactus.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'article','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}



	function setResults(ctrl,val){
		document.forms[0][ctrl].value = val;}

	function getResults(ctrl){
		return document.forms[0][ctrl].value;}

	function goCardCode() {
		var w = 300
		var h = 435
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="cvcard.asp";
		reWin = window.open(theUrl,'cardcode','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}

	 function goCheck() {
		var w = 450
		var h = 300
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="check.asp";
		reWin = window.open(theUrl,'check','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

//\//////////////////////////////////////////////////////////////////////////////////
//\  overLIB 3.51  --  This notice must remain untouched at all times.
//\  Copyright Erik Bosrup 1998-2002. All rights reserved.
//\
//\  By Erik Bosrup (erik@bosrup.com).	Last modified 2002-11-01.
//\  Portions by Dan Steinman (dansteinman.com). Additions by other people are
//\  listed on the overLIB homepage.
//\
//\  Get the latest version at http://www.bosrup.com/web/overlib/
//\
//\  This script is published under an open source license. Please read the license
//\  agreement online at: http://www.bosrup.com/web/overlib/license.html
//\  If you have questions regarding the license please contact erik@bosrup.com.
//\
//\  This script library was originally created for personal use. By request it has
//\  later been made public. This is free software. Do not sell this as your own
//\  work, or remove this copyright notice. For full details on copying or changing
//\  this script please read the license agreement at the link above.
//\
//\  Please give credit on sites that use overLIB and submit changes of the script
//\  so other people can use them as well. This script is free to use, don't abuse.
//\//////////////////////////////////////////////////////////////////////////////////
//\mini


////////////////////////////////////////////////////////////////////////////////////
// CONSTANTS
// Don't touch these. :)
////////////////////////////////////////////////////////////////////////////////////
var INARRAY		=	1;
var CAPARRAY		=	2;
var STICKY		=	3;
var BACKGROUND		=	4;
var NOCLOSE		=	5;
var CAPTION		=	6;
var LEFT		=	7;
var RIGHT		=	8;
var CENTER		=	9;
var OFFSETX		=	10;
var OFFSETY		=	11;
var FGCOLOR		=	12;
var BGCOLOR		=	13;
var TEXTCOLOR		=	14;
var CAPCOLOR		=	15;
var CLOSECOLOR		=	16;
var WIDTH		=	17;
var BORDER		=	18;
var STATUS		=	19;
var AUTOSTATUS		=	20;
var AUTOSTATUSCAP	=	21;
var HEIGHT		=	22;
var CLOSETEXT		=	23;
var SNAPX		=	24;
var SNAPY		=	25;
var FIXX		=	26;
var FIXY		=	27;
var FGBACKGROUND	=	28;
var BGBACKGROUND	=	29;
var PADX		=	30; // PADX2 out
var PADY		=	31; // PADY2 out
var FULLHTML		=	34;
var ABOVE		=	35;
var BELOW		=	36;
var CAPICON		=	37;
var TEXTFONT		=	38;
var CAPTIONFONT 	=	39;
var CLOSEFONT		=	40;
var TEXTSIZE		=	41;
var CAPTIONSIZE 	=	42;
var CLOSESIZE		=	43;
var FRAME		=	44;
var TIMEOUT		=	45;
var FUNCTION		=	46;
var DELAY		=	47;
var HAUTO		=	48;
var VAUTO		=	49;
var CLOSECLICK		=	50;
var CSSOFF		=	51;
var CSSSTYLE		=	52;
var CSSCLASS		=	53;
var FGCLASS		=	54;
var BGCLASS		=	55;
var TEXTFONTCLASS	=	56;
var CAPTIONFONTCLASS	=	57;
var CLOSEFONTCLASS	=	58;
var PADUNIT		=	59;
var HEIGHTUNIT		=	60;
var WIDTHUNIT		=	61;
var TEXTSIZEUNIT	=	62;
var TEXTDECORATION	=	63;
var TEXTSTYLE		=	64;
var TEXTWEIGHT		=	65;
var CAPTIONSIZEUNIT	=	66;
var CAPTIONDECORATION	=	67;
var CAPTIONSTYLE	=	68;
var CAPTIONWEIGHT	=	69;
var CLOSESIZEUNIT	=	70;
var CLOSEDECORATION	=	71;
var CLOSESTYLE		=	72;
var CLOSEWEIGHT 	=	73;


////////////////////////////////////////////////////////////////////////////////////
// DEFAULT CONFIGURATION
// You don't have to change anything here if you don't want to. All of this can be
// changed on your html page or through an overLIB call.
////////////////////////////////////////////////////////////////////////////////////

// Main background color (the large area)
// Usually a bright color (white, yellow etc)
if (typeof ol_fgcolor == 'undefined') { var ol_fgcolor = "#eefadc";}

// Border color and color of caption
// Usually a dark color (black, brown etc)
if (typeof ol_bgcolor == 'undefined') { var ol_bgcolor = "#eefadc";}
	
// Text color
// Usually a dark color
if (typeof ol_textcolor == 'undefined') { var ol_textcolor = "#000000";}
	
// Color of the caption text
// Usually a bright color
if (typeof ol_capcolor == 'undefined') { var ol_capcolor = "#FFFFFF";}
	
// Color of "Close" when using Sticky
// Usually a semi-bright color
if (typeof ol_closecolor == 'undefined') { var ol_closecolor = "#000000";}

// Font face for the main text
if (typeof ol_textfont == 'undefined') { var ol_textfont = "Verdana,Arial,Helvetica";}

// Font face for the caption
if (typeof ol_captionfont == 'undefined') { var ol_captionfont = "Verdana,Arial,Helvetica";}

// Font face for the close text
if (typeof ol_closefont == 'undefined') { var ol_closefont = "Verdana,Arial,Helvetica";}

// Font size for the main text
// When using CSS this will be very small.
if (typeof ol_textsize == 'undefined') { var ol_textsize = "1";}

// Font size for the caption
// When using CSS this will be very small.
if (typeof ol_captionsize == 'undefined') { var ol_captionsize = "1";}

// Font size for the close text
// When using CSS this will be very small.
if (typeof ol_closesize == 'undefined') { var ol_closesize = "1";}

// Width of the popups in pixels
// 100-300 pixels is typical
if (typeof ol_width == 'undefined') { var ol_width = "200";}

// How thick the ol_border should be in pixels
// 1-3 pixels is typical
if (typeof ol_border == 'undefined') { var ol_border = "1";}

// How many pixels to the right/left of the cursor to show the popup
// Values between 3 and 12 are best
if (typeof ol_offsetx == 'undefined') { var ol_offsetx = 10;}
	
// How many pixels to the below the cursor to show the popup
// Values between 3 and 12 are best
if (typeof ol_offsety == 'undefined') { var ol_offsety = 10;}

// Default text for popups
// Should you forget to pass something to overLIB this will be displayed.
if (typeof ol_text == 'undefined') { var ol_text = "Default Text"; }

// Default caption
// You should leave this blank or you will have problems making non caps popups.
if (typeof ol_cap == 'undefined') { var ol_cap = ""; }

// Decides if sticky popups are default.
// 0 for non, 1 for stickies.
if (typeof ol_sticky == 'undefined') { var ol_sticky = 0; }

// Default background image. Better left empty unless you always want one.
if (typeof ol_background == 'undefined') { var ol_background = ""; }

// Text for the closing sticky popups.
// Normal is "Close".
if (typeof ol_close == 'undefined') { var ol_close = "Close"; }

// Default vertical alignment for popups.
// It's best to leave RIGHT here. Other options are LEFT and CENTER.
if (typeof ol_hpos == 'undefined') { var ol_hpos = RIGHT; }

// Default status bar text when a popup is invoked.
if (typeof ol_status == 'undefined') { var ol_status = ""; }

// If the status bar automatically should load either text or caption.
// 0=nothing, 1=text, 2=caption
if (typeof ol_autostatus == 'undefined') { var ol_autostatus = 1; }

// Default height for popup. Often best left alone.
if (typeof ol_height == 'undefined') { var ol_height = 40; }

// Horizontal grid spacing that popups will snap to.
// 0 makes no grid, anything else will cause a snap to that grid spacing.
if (typeof ol_snapx == 'undefined') { var ol_snapx = 0; }

// Vertical grid spacing that popups will snap to.
// 0 makes no grid, andthing else will cause a snap to that grid spacing.
if (typeof ol_snapy == 'undefined') { var ol_snapy = 0; }

// Sets the popups horizontal position to a fixed column.
// Anything above -1 will cause fixed position.
if (typeof ol_fixx == 'undefined') { var ol_fixx = -1; }

// Sets the popups vertical position to a fixed row.
// Anything above -1 will cause fixed position.
if (typeof ol_fixy == 'undefined') { var ol_fixy = -1; }

// Background image for the popups inside.
if (typeof ol_fgbackground == 'undefined') { var ol_fgbackground = ""; }

// Background image for the popups frame.
if (typeof ol_bgbackground == 'undefined') { var ol_bgbackground = ""; }

// How much horizontal left padding text should get by default when BACKGROUND is used.
if (typeof ol_padxl == 'undefined') { var ol_padxl = 1; }

// How much horizontal right padding text should get by default when BACKGROUND is used.
if (typeof ol_padxr == 'undefined') { var ol_padxr = 1; }

// How much vertical top padding text should get by default when BACKGROUND is used.
if (typeof ol_padyt == 'undefined') { var ol_padyt = 1; }

// How much vertical bottom padding text should get by default when BACKGROUND is used.
if (typeof ol_padyb == 'undefined') { var ol_padyb = 1; }

// If the user by default must supply all html for complete popup control.
// Set to 1 to activate, 0 otherwise.
if (typeof ol_fullhtml == 'undefined') { var ol_fullhtml = 0; }

// Default vertical position of the popup. Default should normally be BELOW.
// ABOVE only works when HEIGHT is defined.
if (typeof ol_vpos == 'undefined') { var ol_vpos = BELOW; }

// Default height of popup to use when placing the popup above the cursor.
if (typeof ol_aboveheight == 'undefined') { var ol_aboveheight = 0; }

// Default icon to place next to the popups caption.
if (typeof ol_capicon == 'undefined') { var ol_capicon = ""; }

// Default frame. We default to current frame if there is no frame defined.
if (typeof ol_frame == 'undefined') { var ol_frame = self; }

// Default timeout. By default there is no timeout.
if (typeof ol_timeout == 'undefined') { var ol_timeout = 0; }

// Default javascript funktion. By default there is none.
if (typeof ol_function == 'undefined') { var ol_function = null; }

// Default timeout. By default there is no timeout.
if (typeof ol_delay == 'undefined') { var ol_delay = 0; }

// If overLIB should decide the horizontal placement.
if (typeof ol_hauto == 'undefined') { var ol_hauto = 0; }

// If overLIB should decide the vertical placement.
if (typeof ol_vauto == 'undefined') { var ol_vauto = 0; }



// If the user has to click to close stickies.
if (typeof ol_closeclick == 'undefined') { var ol_closeclick = 0; }

// This variable determines if you want to use CSS or inline definitions.
// CSSOFF=no CSS    CSSSTYLE=use CSS inline styles    CSSCLASS=use classes
if (typeof ol_css == 'undefined') { var ol_css = CSSOFF; }

// Main background class (eqv of fgcolor)
// This is only used if CSS is set to use classes (ol_css = CSSCLASS)
if (typeof ol_fgclass == 'undefined') { var ol_fgclass = ""; }

// Frame background class (eqv of bgcolor)
// This is only used if CSS is set to use classes (ol_css = CSSCLASS)
if (typeof ol_bgclass == 'undefined') { var ol_bgclass = ""; }

// Main font class
// This is only used if CSS is set to use classes (ol_css = CSSCLASS)
if (typeof ol_textfontclass == 'undefined') { var ol_textfontclass = ""; }

// Caption font class
// This is only used if CSS is set to use classes (ol_css = CSSCLASS)
if (typeof ol_captionfontclass == 'undefined') { var ol_captionfontclass = ""; }

// Close font class
// This is only used if CSS is set to use classes (ol_css = CSSCLASS)
if (typeof ol_closefontclass == 'undefined') { var ol_closefontclass = ""; }

// Unit to be used for the text padding above
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
// Options include "px", "%", "in", "cm"
if (typeof ol_padunit == 'undefined') { var ol_padunit = "px";}

// Unit to be used for height of popup
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
// Options include "px", "%", "in", "cm"
if (typeof ol_heightunit == 'undefined') { var ol_heightunit = "px";}

// Unit to be used for width of popup
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
// Options include "px", "%", "in", "cm"
if (typeof ol_widthunit == 'undefined') { var ol_widthunit = "px";}

// Font size unit for the main text
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_textsizeunit == 'undefined') { var ol_textsizeunit = "px";}

// Decoration of the main text ("none", "underline", "line-through" or "blink")
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_textdecoration == 'undefined') { var ol_textdecoration = "none";}

// Font style of the main text ("normal" or "italic")
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_textstyle == 'undefined') { var ol_textstyle = "normal";}

// Font weight of the main text ("normal", "bold", "bolder", "lighter", ect.)
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_textweight == 'undefined') { var ol_textweight = "normal";}

// Font size unit for the caption
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_captionsizeunit == 'undefined') { var ol_captionsizeunit = "px";}

// Decoration of the caption ("none", "underline", "line-through" or "blink")
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_captiondecoration == 'undefined') { var ol_captiondecoration = "none";}

// Font style of the caption ("normal" or "italic")
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_captionstyle == 'undefined') { var ol_captionstyle = "normal";}

// Font weight of the caption ("normal", "bold", "bolder", "lighter", ect.)
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_captionweight == 'undefined') { var ol_captionweight = "bold";}

// Font size unit for the close text
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_closesizeunit == 'undefined') { var ol_closesizeunit = "px";}

// Decoration of the close text ("none", "underline", "line-through" or "blink")
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_closedecoration == 'undefined') { var ol_closedecoration = "none";}

// Font style of the close text ("normal" or "italic")
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_closestyle == 'undefined') { var ol_closestyle = "normal";}

// Font weight of the close text ("normal", "bold", "bolder", "lighter", ect.)
// Only used if CSS inline styles are being used (ol_css = CSSSTYLE)
if (typeof ol_closeweight == 'undefined') { var ol_closeweight = "normal";}



////////////////////////////////////////////////////////////////////////////////////
// ARRAY CONFIGURATION
// You don't have to change anything here if you don't want to. The following
// arrays can be filled with text and html if you don't wish to pass it from
// your html page.
////////////////////////////////////////////////////////////////////////////////////

// Array with texts.
if (typeof ol_texts == 'undefined') { var ol_texts = new Array("Text 0", "Text 1"); }

// Array with captions.
if (typeof ol_caps == 'undefined') { var ol_caps = new Array("Caption 0", "Caption 1"); }


////////////////////////////////////////////////////////////////////////////////////
// END CONFIGURATION
// Don't change anything below this line, all configuration is above.
////////////////////////////////////////////////////////////////////////////////////







////////////////////////////////////////////////////////////////////////////////////
// INIT
////////////////////////////////////////////////////////////////////////////////////

// Runtime variables init. Used for runtime only, don't change, not for config!
var o3_text = "";
var o3_cap = "";
var o3_sticky = 0;
var o3_background = "";
var o3_close = "Close";
var o3_hpos = RIGHT;
var o3_offsetx = 2;
var o3_offsety = 2;
var o3_fgcolor = "";
var o3_bgcolor = "";
var o3_textcolor = "";
var o3_capcolor = "";
var o3_closecolor = "";
var o3_width = 100;
var o3_border = 1;
var o3_status = "";
var o3_autostatus = 0;
var o3_height = -1;
var o3_snapx = 0;
var o3_snapy = 0;
var o3_fixx = -1;
var o3_fixy = -1;
var o3_fgbackground = "";
var o3_bgbackground = "";
var o3_padxl = 0;
var o3_padxr = 0;
var o3_padyt = 0;
var o3_padyb = 0;
var o3_fullhtml = 0;
var o3_vpos = BELOW;
var o3_aboveheight = 0;
var o3_capicon = "";
var o3_textfont = "Verdana,Arial,Helvetica";
var o3_captionfont = "Verdana,Arial,Helvetica";
var o3_closefont = "Verdana,Arial,Helvetica";
var o3_textsize = "1";
var o3_captionsize = "1";
var o3_closesize = "1";
var o3_frame = self;
var o3_timeout = 0;
var o3_timerid = 0;
var o3_allowmove = 0;
var o3_function = null; 
var o3_delay = 0;
var o3_delayid = 0;
var o3_hauto = 0;
var o3_vauto = 0;
var o3_closeclick = 0;

var o3_css = CSSOFF;
var o3_fgclass = "";
var o3_bgclass = "";
var o3_textfontclass = "";
var o3_captionfontclass = "";
var o3_closefontclass = "";
var o3_padunit = "px";
var o3_heightunit = "px";
var o3_widthunit = "px";
var o3_textsizeunit = "px";
var o3_textdecoration = "";
var o3_textstyle = "";
var o3_textweight = "";
var o3_captionsizeunit = "px";
var o3_captiondecoration = "";
var o3_captionstyle = "";
var o3_captionweight = "";
var o3_closesizeunit = "px";
var o3_closedecoration = "";
var o3_closestyle = "";
var o3_closeweight = "";



// Display state variables
var o3_x = 0;
var o3_y = 0;
var o3_allow = 0;
var o3_showingsticky = 0;
var o3_removecounter = 0;

// Our layer
var over = null;
var fnRef;

// Decide browser version
var ns4 = (navigator.appName == 'Netscape' && parseInt(navigator.appVersion) == 4);
var ns6 = (document.getElementById)? true:false;
var ie4 = (document.all)? true:false;
if (ie4) var docRoot = 'document.body';
var ie5 = false;
if (ns4) {
	var oW = window.innerWidth;
	var oH = window.innerHeight;
	window.onresize = function () {if (oW!=window.innerWidth||oH!=window.innerHeight) location.reload();}
}


// Microsoft Stupidity Check(tm).
if (ie4) {
	if ((navigator.userAgent.indexOf('MSIE 5') > 0) || (navigator.userAgent.indexOf('MSIE 6') > 0)) {
		if(document.compatMode && document.compatMode == 'CSS1Compat') docRoot = 'document.documentElement';
		ie5 = true;
	}
	if (ns6) {
		ns6 = false;
	}
}


// Capture events, alt. diffuses the overlib function.
if ( (ns4) || (ie4) || (ns6)) {
	document.onmousemove = mouseMove
	if (ns4) document.captureEvents(Event.MOUSEMOVE)
} else {
	overlib = no_overlib;
	nd = no_overlib;
	ver3fix = true;
}


// Fake function for 3.0 users.
function no_overlib() {
	return ver3fix;
}



////////////////////////////////////////////////////////////////////////////////////
// PUBLIC FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////


// overlib(arg0, ..., argN)
// Loads parameters into global runtime variables.
function overlib() {
	
	// Load defaults to runtime.
	o3_text = ol_text;
	o3_cap = ol_cap;
	o3_sticky = ol_sticky;
	o3_background = ol_background;
	o3_close = ol_close;
	o3_hpos = ol_hpos;
	o3_offsetx = ol_offsetx;
	o3_offsety = ol_offsety;
	o3_fgcolor = ol_fgcolor;
	o3_bgcolor = ol_bgcolor;
	o3_textcolor = ol_textcolor;
	o3_capcolor = ol_capcolor;
	o3_closecolor = ol_closecolor;
	o3_width = ol_width;
	o3_border = ol_border;
	o3_status = ol_status;
	o3_autostatus = ol_autostatus;
	o3_height = ol_height;
	o3_snapx = ol_snapx;
	o3_snapy = ol_snapy;
	o3_fixx = ol_fixx;
	o3_fixy = ol_fixy;
	o3_fgbackground = ol_fgbackground;
	o3_bgbackground = ol_bgbackground;
	o3_padxl = ol_padxl;
	o3_padxr = ol_padxr;
	o3_padyt = ol_padyt;
	o3_padyb = ol_padyb;
	o3_fullhtml = ol_fullhtml;
	o3_vpos = ol_vpos;
	o3_aboveheight = ol_aboveheight;
	o3_capicon = ol_capicon;
	o3_textfont = ol_textfont;
	o3_captionfont = ol_captionfont;
	o3_closefont = ol_closefont;
	o3_textsize = ol_textsize;
	o3_captionsize = ol_captionsize;
	o3_closesize = ol_closesize;
	o3_timeout = ol_timeout;
	o3_function = ol_function;
	o3_delay = ol_delay;
	o3_hauto = ol_hauto;
	o3_vauto = ol_vauto;
	o3_closeclick = ol_closeclick;
	
	o3_css = ol_css;
	o3_fgclass = ol_fgclass;
	o3_bgclass = ol_bgclass;
	o3_textfontclass = ol_textfontclass;
	o3_captionfontclass = ol_captionfontclass;
	o3_closefontclass = ol_closefontclass;
	o3_padunit = ol_padunit;
	o3_heightunit = ol_heightunit;
	o3_widthunit = ol_widthunit;
	o3_textsizeunit = ol_textsizeunit;
	o3_textdecoration = ol_textdecoration;
	o3_textstyle = ol_textstyle;
	o3_textweight = ol_textweight;
	o3_captionsizeunit = ol_captionsizeunit;
	o3_captiondecoration = ol_captiondecoration;
	o3_captionstyle = ol_captionstyle;
	o3_captionweight = ol_captionweight;
	o3_closesizeunit = ol_closesizeunit;
	o3_closedecoration = ol_closedecoration;
	o3_closestyle = ol_closestyle;
	o3_closeweight = ol_closeweight;
	fnRef = '';
	

	// Special for frame support, over must be reset...
	if ( (ns4) || (ie4) || (ns6) ) {
		if (over) cClick();
		o3_frame = ol_frame;
		if (ns4) over = o3_frame.document.overDiv
		if (ie4) over = o3_frame.overDiv.style
		if (ns6) over = o3_frame.document.getElementById("overDiv");
	}
	
	
	// What the next argument is expected to be.
	var parsemode = -1, udf, v = null;
	
	var ar = arguments;
	udf = (!ar.length ? 1 : 0);

	for (i = 0; i < ar.length; i++) {

		if (parsemode < 0) {
			// Arg is maintext, unless its a PARAMETER
			if (typeof ar[i] == 'number') {
				udf = (ar[i] == FUNCTION ? 0 : 1);
				i--;
			} else {
				o3_text = ar[i];
			}

			parsemode = 0;
		} else {
			// Note: NS4 doesn't like switch cases with vars.
			if (ar[i] == INARRAY) { udf = 0; o3_text = ol_texts[ar[++i]]; continue; }
			if (ar[i] == CAPARRAY) { o3_cap = ol_caps[ar[++i]]; continue; }
			if (ar[i] == STICKY) { o3_sticky = 1; continue; }
			if (ar[i] == BACKGROUND) { o3_background = ar[++i]; continue; }
			if (ar[i] == NOCLOSE) { o3_close = ""; continue; }
			if (ar[i] == CAPTION) { o3_cap = ar[++i]; continue; }
			if (ar[i] == CENTER || ar[i] == LEFT || ar[i] == RIGHT) { o3_hpos = ar[i]; continue; }
			if (ar[i] == OFFSETX) { o3_offsetx = ar[++i]; continue; }
			if (ar[i] == OFFSETY) { o3_offsety = ar[++i]; continue; }
			if (ar[i] == FGCOLOR) { o3_fgcolor = ar[++i]; continue; }
			if (ar[i] == BGCOLOR) { o3_bgcolor = ar[++i]; continue; }
			if (ar[i] == TEXTCOLOR) { o3_textcolor = ar[++i]; continue; }
			if (ar[i] == CAPCOLOR) { o3_capcolor = ar[++i]; continue; }
			if (ar[i] == CLOSECOLOR) { o3_closecolor = ar[++i]; continue; }
			if (ar[i] == WIDTH) { o3_width = ar[++i]; continue; }
			if (ar[i] == BORDER) { o3_border = ar[++i]; continue; }
			if (ar[i] == STATUS) { o3_status = ar[++i]; continue; }
			if (ar[i] == AUTOSTATUS) { o3_autostatus = (o3_autostatus == 1) ? 0 : 1; continue; }
			if (ar[i] == AUTOSTATUSCAP) { o3_autostatus = (o3_autostatus == 2) ? 0 : 2; continue; }
			if (ar[i] == HEIGHT) { o3_height = ar[++i]; o3_aboveheight = ar[i]; continue; } // Same param again.
			if (ar[i] == CLOSETEXT) { o3_close = ar[++i]; continue; }
			if (ar[i] == SNAPX) { o3_snapx = ar[++i]; continue; }
			if (ar[i] == SNAPY) { o3_snapy = ar[++i]; continue; }
			if (ar[i] == FIXX) { o3_fixx = ar[++i]; continue; }
			if (ar[i] == FIXY) { o3_fixy = ar[++i]; continue; }
			if (ar[i] == FGBACKGROUND) { o3_fgbackground = ar[++i]; continue; }
			if (ar[i] == BGBACKGROUND) { o3_bgbackground = ar[++i]; continue; }
			if (ar[i] == PADX) { o3_padxl = ar[++i]; o3_padxr = ar[++i]; continue; }
			if (ar[i] == PADY) { o3_padyt = ar[++i]; o3_padyb = ar[++i]; continue; }
			if (ar[i] == FULLHTML) { o3_fullhtml = 1; continue; }
			if (ar[i] == BELOW || ar[i] == ABOVE) { o3_vpos = ar[i]; continue; }
			if (ar[i] == CAPICON) { o3_capicon = ar[++i]; continue; }
			if (ar[i] == TEXTFONT) { o3_textfont = ar[++i]; continue; }
			if (ar[i] == CAPTIONFONT) { o3_captionfont = ar[++i]; continue; }
			if (ar[i] == CLOSEFONT) { o3_closefont = ar[++i]; continue; }
			if (ar[i] == TEXTSIZE) { o3_textsize = ar[++i]; continue; }
			if (ar[i] == CAPTIONSIZE) { o3_captionsize = ar[++i]; continue; }
			if (ar[i] == CLOSESIZE) { o3_closesize = ar[++i]; continue; }
			if (ar[i] == FRAME) { opt_FRAME(ar[++i]); continue; }
			if (ar[i] == TIMEOUT) { o3_timeout = ar[++i]; continue; }
			if (ar[i] == FUNCTION) { udf = 0; if (typeof ar[i+1] != 'number') v = ar[++i]; opt_FUNCTION(v); continue; } 
			if (ar[i] == DELAY) { o3_delay = ar[++i]; continue; }
			if (ar[i] == HAUTO) { o3_hauto = (o3_hauto == 0) ? 1 : 0; continue; }
			if (ar[i] == VAUTO) { o3_vauto = (o3_vauto == 0) ? 1 : 0; continue; }
			if (ar[i] == CLOSECLICK) { o3_closeclick = (o3_closeclick == 0) ? 1 : 0; continue; }
			if (ar[i] == CSSOFF) { o3_css = ar[i]; continue; }
			if (ar[i] == CSSSTYLE) { o3_css = ar[i]; continue; }
			if (ar[i] == CSSCLASS) { o3_css = ar[i]; continue; }
			if (ar[i] == FGCLASS) { o3_fgclass = ar[++i]; continue; }
			if (ar[i] == BGCLASS) { o3_bgclass = ar[++i]; continue; }
			if (ar[i] == TEXTFONTCLASS) { o3_textfontclass = ar[++i]; continue; }
			if (ar[i] == CAPTIONFONTCLASS) { o3_captionfontclass = ar[++i]; continue; }
			if (ar[i] == CLOSEFONTCLASS) { o3_closefontclass = ar[++i]; continue; }
			if (ar[i] == PADUNIT) { o3_padunit = ar[++i]; continue; }
			if (ar[i] == HEIGHTUNIT) { o3_heightunit = ar[++i]; continue; }
			if (ar[i] == WIDTHUNIT) { o3_widthunit = ar[++i]; continue; }
			if (ar[i] == TEXTSIZEUNIT) { o3_textsizeunit = ar[++i]; continue; }
			if (ar[i] == TEXTDECORATION) { o3_textdecoration = ar[++i]; continue; }
			if (ar[i] == TEXTSTYLE) { o3_textstyle = ar[++i]; continue; }
			if (ar[i] == TEXTWEIGHT) { o3_textweight = ar[++i]; continue; }
			if (ar[i] == CAPTIONSIZEUNIT) { o3_captionsizeunit = ar[++i]; continue; }
			if (ar[i] == CAPTIONDECORATION) { o3_captiondecoration = ar[++i]; continue; }
			if (ar[i] == CAPTIONSTYLE) { o3_captionstyle = ar[++i]; continue; }
			if (ar[i] == CAPTIONWEIGHT) { o3_captionweight = ar[++i]; continue; }
			if (ar[i] == CLOSESIZEUNIT) { o3_closesizeunit = ar[++i]; continue; }
			if (ar[i] == CLOSEDECORATION) { o3_closedecoration = ar[++i]; continue; }
			if (ar[i] == CLOSESTYLE) { o3_closestyle = ar[++i]; continue; }
			if (ar[i] == CLOSEWEIGHT) { o3_closeweight = ar[++i]; continue; }
		}
	}
	if (udf && o3_function) o3_text = o3_function();

	if (o3_delay == 0) {
		return overlib351();
	} else {
		o3_delayid = setTimeout("overlib351()", o3_delay);
		return false;
	}
}



// Clears popups if appropriate
function nd() {
	if ( o3_removecounter >= 1 ) { o3_showingsticky = 0 };
	if ( (ns4) || (ie4) || (ns6) ) {
		if ( o3_showingsticky == 0 ) {
			o3_allowmove = 0;
			if (over != null) hideObject(over);
		} else {
			o3_removecounter++;
		}
	}
	
	return true;
}







////////////////////////////////////////////////////////////////////////////////////
// OVERLIB 3.51 FUNCTION
////////////////////////////////////////////////////////////////////////////////////


// This function decides what it is we want to display and how we want it done.
function overlib351() {

	// Make layer content
	var layerhtml;

	if (o3_background != "" || o3_fullhtml) {
		// Use background instead of box.
		layerhtml = ol_content_background(o3_text, o3_background, o3_fullhtml);
	} else {
		// They want a popup box.

		// Prepare popup background
		if (o3_fgbackground != "" && o3_css == CSSOFF) {
			o3_fgbackground = "BACKGROUND=\""+o3_fgbackground+"\"";
		}
		if (o3_bgbackground != "" && o3_css == CSSOFF) {
			o3_bgbackground = "BACKGROUND=\""+o3_bgbackground+"\"";
		}

		// Prepare popup colors
		if (o3_fgcolor != "" && o3_css == CSSOFF) {
			o3_fgcolor = "BGCOLOR=\""+o3_fgcolor+"\"";
		}
		if (o3_bgcolor != "" && o3_css == CSSOFF) {
			o3_bgcolor = "BGCOLOR=\""+o3_bgcolor+"\"";
		}

		// Prepare popup height
		if (o3_height > 0 && o3_css == CSSOFF) {
			o3_height = "HEIGHT=" + o3_height;
		} else {
			o3_height = "";
		}

		// Decide which kinda box.
		if (o3_cap == "") {
			// Plain
			layerhtml = ol_content_simple(o3_text);
		} else {
			// With caption
			if (o3_sticky) {
				// Show close text
				layerhtml = ol_content_caption(o3_text, o3_cap, o3_close);
			} else {
				// No close text
				layerhtml = ol_content_caption(o3_text, o3_cap, "");
			}
		}
	}
	
	// We want it to stick!
	if (o3_sticky) {
		if (o3_timerid > 0) {
			clearTimeout(o3_timerid);
			o3_timerid = 0;
		}
		o3_showingsticky = 1;
		o3_removecounter = 0;
	}
	
	// Write layer
	layerWrite(layerhtml);
	
	// Prepare status bar
	if (o3_autostatus > 0) {
		o3_status = o3_text;
		if (o3_autostatus > 1) {
			o3_status = o3_cap;
		}
	}

	// When placing the layer the first time, even stickies may be moved.
	o3_allowmove = 0;

	// Initiate a timer for timeout
	if (o3_timeout > 0) {
		if (o3_timerid > 0) clearTimeout(o3_timerid);
		o3_timerid = setTimeout("cClick()", o3_timeout);
	}

	// Show layer
	disp(o3_status);

	// Stickies should stay where they are. 
	if (o3_sticky) o3_allowmove = 0;

	return (o3_status != '');
}



////////////////////////////////////////////////////////////////////////////////////
// LAYER GENERATION FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////

// Makes simple table without caption
function ol_content_simple(text) {
	if (o3_css == CSSCLASS) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING="+o3_border+" CELLSPACING=0 class=\""+o3_bgclass+"\"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 class=\""+o3_fgclass+"\"><TR><TD VALIGN=TOP><FONT class=\""+o3_textfontclass+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>";
	if (o3_css == CSSSTYLE) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING="+o3_border+" CELLSPACING=0 style=\"background-color: "+o3_bgcolor+"; height: "+o3_height+o3_heightunit+";\"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 style=\"color: "+o3_fgcolor+"; background-color: "+o3_fgcolor+"; height: "+o3_height+o3_heightunit+";\"><TR><TD VALIGN=TOP><FONT style=\"font-family: "+o3_textfont+"; color: "+o3_textcolor+"; font-size: "+o3_textsize+o3_textsizeunit+"; text-decoration: "+o3_textdecoration+"; font-weight: "+o3_textweight+"; font-style:"+o3_textstyle+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>";
	if (o3_css == CSSOFF) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING="+o3_border+" CELLSPACING=0 "+o3_bgcolor+" "+o3_height+"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 "+o3_fgcolor+" "+o3_fgbackground+" "+o3_height+"><TR><TD VALIGN=TOP><FONT FACE=\""+o3_textfont+"\" COLOR=\""+o3_textcolor+"\" SIZE=\""+o3_textsize+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>";

	set_background("");
	return txt;
}




// Makes table with caption and optional close link
function ol_content_caption(text, title, close) {
	closing = "";
	closeevent = "onMouseOver";

	if (o3_closeclick == 1) closeevent = "onClick";
	if (o3_capicon != "") o3_capicon = "<IMG SRC=\""+o3_capicon+"\"> ";

	if (close != "") {
		if (o3_css == CSSCLASS) closing = "<TD ALIGN=RIGHT><A HREF=\"javascript:return "+fnRef+"cClick();\" "+closeevent+"=\"return " + fnRef + "cClick();\" class=\""+o3_closefontclass+"\">"+close+"</A></TD>";
		if (o3_css == CSSSTYLE) closing = "<TD ALIGN=RIGHT><A HREF=\"javascript:return "+fnRef+"cClick();\" "+closeevent+"=\"return " + fnRef + "cClick();\" style=\"color: "+o3_closecolor+"; font-family: "+o3_closefont+"; font-size: "+o3_closesize+o3_closesizeunit+"; text-decoration: "+o3_closedecoration+"; font-weight: "+o3_closeweight+"; font-style:"+o3_closestyle+";\">"+close+"</A></TD>";
		if (o3_css == CSSOFF) closing = "<TD ALIGN=RIGHT><A HREF=\"javascript:return "+fnRef+"cClick();\" "+closeevent+"=\"return " + fnRef + "cClick();\"><FONT COLOR=\""+o3_closecolor+"\" FACE=\""+o3_closefont+"\" SIZE=\""+o3_closesize+"\">"+close+"</FONT></A></TD>";
	}

	if (o3_css == CSSCLASS) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING="+o3_border+" CELLSPACING=0 class=\""+o3_bgclass+"\"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD><FONT class=\""+o3_captionfontclass+"\">"+o3_capicon+title+"</FONT></TD>"+closing+"</TR></TABLE><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 class=\""+o3_fgclass+"\"><TR><TD VALIGN=TOP><FONT class=\""+o3_textfontclass+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>";
	if (o3_css == CSSSTYLE) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING="+o3_border+" CELLSPACING=0 style=\"background-color: "+o3_bgcolor+"; background-image: url("+o3_bgbackground+"); height: "+o3_height+o3_heightunit+";\"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD><FONT style=\"font-family: "+o3_captionfont+"; color: "+o3_capcolor+"; font-size: "+o3_captionsize+o3_captionsizeunit+"; font-weight: "+o3_captionweight+"; font-style: "+o3_captionstyle+"; text-decoration: " + o3_captiondecoration + ";\">"+o3_capicon+title+"</FONT></TD>"+closing+"</TR></TABLE><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 style=\"color: "+o3_fgcolor+"; background-color: "+o3_fgcolor+"; height: "+o3_height+o3_heightunit+";\"><TR><TD VALIGN=TOP><FONT style=\"font-family: "+o3_textfont+"; color: "+o3_textcolor+"; font-size: "+o3_textsize+o3_textsizeunit+"; text-decoration: "+o3_textdecoration+"; font-weight: "+o3_textweight+"; font-style:"+o3_textstyle+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>";
	if (o3_css == CSSOFF) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING="+o3_border+" CELLSPACING=0 "+o3_bgcolor+" "+o3_bgbackground+" "+o3_height+"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD><B><FONT COLOR=\""+o3_capcolor+"\" FACE=\""+o3_captionfont+"\" SIZE=\""+o3_captionsize+"\">"+o3_capicon+title+"</FONT></B></TD>"+closing+"</TR></TABLE><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 "+o3_fgcolor+" "+o3_fgbackground+" "+o3_height+"><TR><TD VALIGN=TOP><FONT COLOR=\""+o3_textcolor+"\" FACE=\""+o3_textfont+"\" SIZE=\""+o3_textsize+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>";

	set_background("");
	return txt;
}

// Sets the background picture, padding and lots more. :)
function ol_content_background(text, picture, hasfullhtml) {
	var txt;
	if (hasfullhtml) {
		txt = text;
	} else {
		var pU, hU, wU;
		pU = (o3_padunit == '%' ? '%' : '');
		hU = (o3_heightunit == '%' ? '%' : '');
		wU = (o3_widthunit == '%' ? '%' : '');

		if (o3_css == CSSCLASS) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING=0 CELLSPACING=0 HEIGHT="+o3_height+"><TR><TD COLSPAN=3 HEIGHT="+o3_padyt+"></TD></TR><TR><TD WIDTH="+o3_padxl+"></TD><TD VALIGN=TOP WIDTH="+(o3_width-o3_padxl-o3_padxr)+"><FONT class=\""+o3_textfontclass+"\">"+text+"</FONT></TD><TD WIDTH="+o3_padxr+"></TD></TR><TR><TD COLSPAN=3 HEIGHT="+o3_padyb+"></TD></TR></TABLE>";
		if (o3_css == CSSSTYLE) txt = "<TABLE WIDTH="+o3_width+wU+" BORDER=0 CELLPADDING=0 CELLSPACING=0 HEIGHT="+o3_height+hU+"><TR><TD COLSPAN=3 HEIGHT="+o3_padyt+pU+"></TD></TR><TR><TD WIDTH="+o3_padxl+pU+"></TD><TD VALIGN=TOP WIDTH="+(o3_width-o3_padxl-o3_padxr)+pU+"><FONT style=\"font-family: "+o3_textfont+"; color: "+o3_textcolor+"; font-size: "+o3_textsize+o3_textsizeunit+";\">"+text+"</FONT></TD><TD WIDTH="+o3_padxr+pU+"></TD></TR><TR><TD COLSPAN=3 HEIGHT="+o3_padyb+pU+"></TD></TR></TABLE>";
		if (o3_css == CSSOFF) txt = "<TABLE WIDTH="+o3_width+" BORDER=0 CELLPADDING=0 CELLSPACING=0 HEIGHT="+o3_height+"><TR><TD COLSPAN=3 HEIGHT="+o3_padyt+"></TD></TR><TR><TD WIDTH="+o3_padxl+"></TD><TD VALIGN=TOP WIDTH="+(o3_width-o3_padxl-o3_padxr)+"><FONT FACE=\""+o3_textfont+"\" COLOR=\""+o3_textcolor+"\" SIZE=\""+o3_textsize+"\">"+text+"</FONT></TD><TD WIDTH="+o3_padxr+"></TD></TR><TR><TD COLSPAN=3 HEIGHT="+o3_padyb+"></TD></TR></TABLE>";
	}
	set_background(picture);
	return txt;
}

// Loads a picture into the div.
function set_background(pic) {
	if (pic == "") {
		if (ns4) over.background.src = null;
		if (ie4) over.backgroundImage = "none";
		if (ns6) over.style.backgroundImage = "none";
	} else {
		if (ns4) {
			over.background.src = pic;
		} else if (ie4) {
			over.backgroundImage = "url("+pic+")";
		} else if (ns6) {
			over.style.backgroundImage = "url("+pic+")";
		}
	}
}



////////////////////////////////////////////////////////////////////////////////////
// HANDLING FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////


// Displays the popup
function disp(statustext) {
	if ( (ns4) || (ie4) || (ns6) ) {
		if (o3_allowmove == 0)	{
			placeLayer();
			showObject(over);
			o3_allowmove = 1;
		}
	}

	if (statustext != "") {
		self.status = statustext;
	}
}

// Decides where we want the popup.
function placeLayer() {
	var placeX, placeY;
	
	// HORIZONTAL PLACEMENT
	if (o3_fixx > -1) {
		// Fixed position
		placeX = o3_fixx;
	} else {
		winoffset = (ie4) ? eval('o3_frame.'+docRoot+'.scrollLeft') : o3_frame.pageXOffset;
		if (ie4) iwidth = eval('o3_frame.'+docRoot+'.clientWidth');
		if (ns4 || ns6) iwidth = o3_frame.innerWidth;
		
		// If HAUTO, decide what to use.
		if (o3_hauto == 1) {
			if ( (o3_x - winoffset) > ((eval(iwidth)) / 2)) {
				o3_hpos = LEFT;
			} else {
				o3_hpos = RIGHT;
			}
		}
		
		// From mouse
		if (o3_hpos == CENTER) { // Center
			placeX = o3_x+o3_offsetx-(o3_width/2);
			if (placeX < winoffset) placeX = winoffset;
		}
		if (o3_hpos == RIGHT) { // Right
			placeX = o3_x+o3_offsetx;
			if ( (eval(placeX) + eval(o3_width)) > (winoffset + iwidth) ) {
				placeX = iwidth + winoffset - o3_width;
				if (placeX < 0) placeX = 0;
			}
		}
		if (o3_hpos == LEFT) { // Left
			placeX = o3_x-o3_offsetx-o3_width;
			if (placeX < winoffset) placeX = winoffset;
		}
	
		// Snapping!
		if (o3_snapx > 1) {
			var snapping = placeX % o3_snapx;
			if (o3_hpos == LEFT) {
				placeX = placeX - (o3_snapx + snapping);
			} else {
				// CENTER and RIGHT
				placeX = placeX + (o3_snapx - snapping);
			}
			if (placeX < winoffset) placeX = winoffset;
		}
	}

	
	
	// VERTICAL PLACEMENT
	if (o3_fixy > -1) {
		// Fixed position
		placeY = o3_fixy;
	} else {
		scrolloffset = (ie4) ? eval('o3_frame.'+docRoot+'.scrollTop') : o3_frame.pageYOffset;

		// If VAUTO, decide what to use.
		if (o3_vauto == 1) {
			if (ie4) iheight = eval('o3_frame.'+docRoot+'.clientHeight');
			if (ns4 || ns6) iheight = o3_frame.innerHeight;

			iheight = (eval(iheight)) / 2;
			if ( (o3_y - scrolloffset) > iheight) {
				o3_vpos = ABOVE;
			} else {
				o3_vpos = BELOW;
			}
		}


		// From mouse
		if (o3_vpos == ABOVE) {
			if (o3_aboveheight == 0) {
				var divref = (ie4) ? o3_frame.document.all['overDiv'] : over;
				o3_aboveheight = (ns4) ? divref.clip.height : divref.offsetHeight;
			}

			placeY = o3_y - (o3_aboveheight + o3_offsety);
			if (placeY < scrolloffset) placeY = scrolloffset;
		} else {
			// BELOW
			placeY = o3_y + o3_offsety;
		}

		// Snapping!
		if (o3_snapy > 1) {
			var snapping = placeY % o3_snapy;
			
			if (o3_aboveheight > 0 && o3_vpos == ABOVE) {
				placeY = placeY - (o3_snapy + snapping);
			} else {
				placeY = placeY + (o3_snapy - snapping);
			}
			
			if (placeY < scrolloffset) placeY = scrolloffset;
		}
	}


	// Actually move the object.	
	repositionTo(over, placeX, placeY);
}


// Moves the layer
function mouseMove(e) {
	if ( (ns4) || (ns6) ) {o3_x=e.pageX; o3_y=e.pageY;}
	if (ie4) {o3_x=event.x; o3_y=event.y;}
	if (ie5) {o3_x=eval('event.x+o3_frame.'+docRoot+'.scrollLeft'); o3_y=eval('event.y+o3_frame.'+docRoot+'.scrollTop');}
	
	if (o3_allowmove == 1) {
		placeLayer();
	}
}

// The Close onMouseOver function for stickies
function cClick() {
	hideObject(over);
	o3_showingsticky = 0;
	
	return false;
}


// Makes sure target frame has overLIB
function compatibleframe(frameid) {
	if (ns4) {
		if (typeof frameid.document.overDiv =='undefined') return false;
	} else if (ie4) {
		if (typeof frameid.document.all["overDiv"] =='undefined') return false;
	} else if (ns6) {
		if (frameid.document.getElementById('overDiv') == null) return false;
	}

	return true;
}



////////////////////////////////////////////////////////////////////////////////////
// LAYER FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////


// Writes to a layer
function layerWrite(txt) {
	txt += "\n";
	
	if (ns4) {
		var lyr = o3_frame.document.overDiv.document
		lyr.write(txt)
		lyr.close()
	} else if (ie4) {
		o3_frame.document.all["overDiv"].innerHTML = txt
	} else if (ns6) {
		range = o3_frame.document.createRange();
		range.setStartBefore(over);
		domfrag = range.createContextualFragment(txt);
		while (over.hasChildNodes()) {
			over.removeChild(over.lastChild);
		}
		over.appendChild(domfrag);
	}
}

// Make an object visible
function showObject(obj) {
	if (ns4) obj.visibility = "show";
	else if (ie4) obj.visibility = "visible";
	else if (ns6) obj.style.visibility = "visible";
}

// Hides an object
function hideObject(obj) {
	if (ns4) obj.visibility = "hide";
	else if (ie4) obj.visibility = "hidden";
	else if (ns6) obj.style.visibility = "hidden";

	if (o3_timerid > 0) clearTimeout(o3_timerid);
	if (o3_delayid > 0) clearTimeout(o3_delayid);
	o3_timerid = 0;
	o3_delayid = 0;
	self.status = "";
}

// Move a layer
function repositionTo(obj,xL,yL) {
	if ( (ns4) || (ie4) ) {
		obj.left = (ie4 ? xL + 'px' : xL);
		obj.top = (ie4 ? yL + 'px' : yL);
	} else if (ns6) {
		obj.style.left = xL + "px";
		obj.style.top = yL+ "px";
	}
}

function getFrameRef(thisFrame, ofrm) {
	var retVal = '';
	for (var i=0; i<thisFrame.length; i++) {
		if (thisFrame[i].length > 0) { 
			retVal = getFrameRef(thisFrame[i],ofrm);
			if (retVal == '') continue;
		} else if (thisFrame[i] != ofrm) continue;
		
		retVal = '['+i+']' + retVal;
		break;
	}
	
	return retVal;
}




////////////////////////////////////////////////////////////////////////////////////
// PARSER FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////


// Defines which frame we should point to.
function opt_FRAME(frm) {
	o3_frame = compatibleframe(frm) ? frm : ol_frame;

	if (o3_frame != ol_frame) {
		var tFrm = getFrameRef(top.frames, o3_frame);
		var sFrm = getFrameRef(top.frames, ol_frame);

		if (sFrm.length == tFrm.length) { 
			l = tFrm.lastIndexOf('['); 
			if (l) {
				while(sFrm.substring(0,l) != tFrm.substring(0,l)) l = tFrm.lastIndexOf('[',l-1);
				tFrm = tFrm.substr(l);
				sFrm = sFrm.substr(l);
			}
		}
			
		var cnt = 0, p = '', str = tFrm;
			
		while((k = str.lastIndexOf('[')) != -1) {
			cnt++;
			str = str.substring(0,k);
		}

		for (var i=0; i<cnt; i++) p = p + 'parent.';
		fnRef = p + 'frames' + sFrm + '.';
	}

	if ( (ns4) || (ie4 || (ns6)) ) {
		if (ns4) over = o3_frame.document.overDiv;
		if (ie4) over = o3_frame.overDiv.style;
		if (ns6) over = o3_frame.document.getElementById("overDiv");
	}

	return 0;
}

// Calls an external function
function opt_FUNCTION(callme) {
	o3_text = (callme ? callme() : (o3_function ? o3_function() : 'No Function'));
	return 0;
}




//end (For internal purposes.)
////////////////////////////////////////////////////////////////////////////////////
// OVERLIB 2 COMPATABILITY FUNCTIONS
// If you aren't upgrading you can remove the below section.
////////////////////////////////////////////////////////////////////////////////////

// Converts old 0=left, 1=right and 2=center into constants.
function vpos_convert(d) {
	if (d == 0) {
		d = LEFT;
	} else {
		if (d == 1) {
			d = RIGHT;
		} else {
			d = CENTER;
		}
	}
	
	return d;
}

// Simple popup
function dts(d,text) {
	o3_hpos = vpos_convert(d);
	overlib(text, o3_hpos, CAPTION, "");
}

// Caption popup
function dtc(d,text, title) {
	o3_hpos = vpos_convert(d);
	overlib(text, CAPTION, title, o3_hpos);
}

// Sticky
function stc(d,text, title) {
	o3_hpos = vpos_convert(d);
	overlib(text, CAPTION, title, o3_hpos, STICKY);
}

// Simple popup right
function drs(text) {
	dts(1,text);
}

// Caption popup right
function drc(text, title) {
	dtc(1,text,title);
}

// Sticky caption right
function src(text,title) {
	stc(1,text,title);
}

// Simple popup left
function dls(text) {
	dts(0,text);
}

// Caption popup left
function dlc(text, title) {
	dtc(0,text,title);
}

// Sticky caption left
function slc(text,title) {
	stc(0,text,title);
}

// Simple popup center
function dcs(text) {
	dts(2,text);
}

// Caption popup center
function dcc(text, title) {
	dtc(2,text,title);
}

// Sticky caption center
function scc(text,title) {
	stc(2,text,title);
}

function isValidDate(dateStr) {
// Checks for the following valid date formats:
// MM/DD/YY   MM/DD/YYYY   MM-DD-YY   MM-DD-YYYY
// Also separates date into month, day, and year variables

var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{2}|\d{4})$/;

// To require a 4 digit year entry, use this line instead:
// var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{4})$/;

var matchArray = dateStr.match(datePat); // is the format ok?
if (matchArray == null) {
return false;
}
month = matchArray[1]; // parse date into variables
day = matchArray[3];
year = matchArray[4];
if (month < 1 || month > 12) { // check month range
return false;
}
if (day < 1 || day > 31) {
return false;
}
if ((month==4 || month==6 || month==9 || month==11) && day==31) {
return false
}
if (month == 2) { // check for february 29th
var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
if (day>29 || (day==29 && !isleap)) {
return false;
   }
}
return true;  // date is valid
}

