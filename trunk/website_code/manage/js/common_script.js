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
function showHideForm1(div,field,nest){
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


function showHideForm(div,field,nest){

  // (document.forms[0].elements(field).checked)
  // changed to eval("document.forms[0]."+field+".checked")

  e = eval("document.forms[0]."+field+".checked")
  
  if (e)
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

        this.old_onsubmit();
	if (submitForm()) 
	{
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
	 alert("Your session has been inactive for 40 minutes.\nTo protect your personal information all inactive sessions will expire after 45 minutes.\nAny changes you have made will be lost if they are not saved within the next 5 minutes!");
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

	function goImagePicker(resCtrl) {
		var w = 500
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

function goSupportFileUploader(resCtrl) {
		var w = 380
		var h = 170
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="supportfile_uploader.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'supportfileUploader','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=on,resizable=yes,'+winprops);}


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

	function goItemsSelect(resCtrl) {
		var w = 600;
		var h = 500;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="items_select.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'itemsSelect','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,'+winprops);}

        function goShowPicture(picture) {
		var w = 650;
		var h = 480;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="show_image.asp?returnArg="+picture;
		reWin = window.open(theUrl,'picker','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,'+winprops);}

	
	function goArticle(resCtrl) {
		var w = 650;
		var h = 480;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="article.asp?TOPIC="+resCtrl;
		reWin = window.open(theUrl,'article','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}

	function goSend(resCtrl) {
		var w = 470;
		var h = 275;
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl="contactus.asp?returnArg="+resCtrl;
		reWin = window.open(theUrl,'article','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,'+winprops);}



	function setResults(ctrl,val){
		document.forms[0][ctrl].value = val;
		document.forms[0][ctrl].focus();
		}

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






/**
 * mm_menu 20MAR2002 Version 6.0
 * Andy Finnell, March 2002
 * Copyright (c) 2000-2002 Macromedia, Inc.
 *
 * based on menu.js
 * by gary smith, July 1997
 * Copyright (c) 1997-1999 Netscape Communications Corp.
 *
 * Netscape grants you a royalty free license to use or modify this
 * software provided that this copyright notice appears on all copies.
 * This software is provided "AS IS," without a warranty of any kind.
 */
function Menu(label, mw, mh, fnt, fs, fclr, fhclr, bg, bgh, halgn, valgn, pad, space, to, sx, sy, srel, opq, vert, idt, aw, ah) 
{
	this.version = "020320 [Menu; mm_menu.js]";
	this.type = "Menu";
	this.menuWidth = mw;
	this.menuItemHeight = mh;
	this.fontSize = fs;
	this.fontWeight = "plain";
	this.fontFamily = fnt;
	this.fontColor = fclr;
	this.fontColorHilite = fhclr;
	this.bgColor = "#555555";
	this.menuBorder = 1;
	this.menuBgOpaque=opq;
	this.menuItemBorder = 1;
	this.menuItemIndent = idt;
	this.menuItemBgColor = bg;
	this.menuItemVAlign = valgn;
	this.menuItemHAlign = halgn;
	this.menuItemPadding = pad;
	this.menuItemSpacing = space;
	this.menuLiteBgColor = "#ffffff";
	this.menuBorderBgColor = "#777777";
	this.menuHiliteBgColor = bgh;
	this.menuContainerBgColor = "#cccccc";
	this.childMenuIcon = "arrows.gif";
	this.submenuXOffset = sx;
	this.submenuYOffset = sy;
	this.submenuRelativeToItem = srel;
	this.vertical = vert;
	this.items = new Array();
	this.actions = new Array();
	this.childMenus = new Array();
	this.hideOnMouseOut = true;
	this.hideTimeout = to;
	this.addMenuItem = addMenuItem;
	this.writeMenus = writeMenus;
	this.MM_showMenu = MM_showMenu;
	this.onMenuItemOver = onMenuItemOver;
	this.onMenuItemAction = onMenuItemAction;
	this.hideMenu = hideMenu;
	this.hideChildMenu = hideChildMenu;
	if (!window.menus) window.menus = new Array();
	this.label = " " + label;
	window.menus[this.label] = this;
	window.menus[window.menus.length] = this;
	if (!window.activeMenus) window.activeMenus = new Array();
}

function addMenuItem(label, action) {
	this.items[this.items.length] = label;
	this.actions[this.actions.length] = action;
}

function FIND(item) {
	if( window.mmIsOpera ) return(document.getElementById(item));
	if (document.all) return(document.all[item]);
	if (document.getElementById) return(document.getElementById(item));
	return(false);
}

function writeMenus(container) {
	if (window.triedToWriteMenus) return;
	var agt = navigator.userAgent.toLowerCase();
	window.mmIsOpera = agt.indexOf("opera") != -1;
	if (!container && document.layers) {
		window.delayWriteMenus = this.writeMenus;
		var timer = setTimeout('delayWriteMenus()', 500);
		container = new Layer(100);
		clearTimeout(timer);
	} else if (document.all || document.hasChildNodes || window.mmIsOpera) {
		document.writeln('<span id="menuContainer"></span>');
		container = FIND("menuContainer");
	}

	window.mmHideMenuTimer = null;
	if (!container) return;	
	window.triedToWriteMenus = true; 
	container.isContainer = true;
	container.menus = new Array();
	for (var i=0; i<window.menus.length; i++) 
		container.menus[i] = window.menus[i];
	window.menus.length = 0;
	var countMenus = 0;
	var countItems = 0;
	var top = 0;
	var content = '';
	var lrs = false;
	var theStat = "";
	var tsc = 0;
	if (document.layers) lrs = true;
	for (var i=0; i<container.menus.length; i++, countMenus++) {
		var menu = container.menus[i];
		if (menu.bgImageUp || !menu.menuBgOpaque) {
			menu.menuBorder = 0;
			menu.menuItemBorder = 0;
		}
		if (lrs) {
			var menuLayer = new Layer(100, container);
			var lite = new Layer(100, menuLayer);
			lite.top = menu.menuBorder;
			lite.left = menu.menuBorder;
			var body = new Layer(100, lite);
			body.top = menu.menuBorder;
			body.left = menu.menuBorder;
		} else {
			content += ''+
			'<div id="menuLayer'+ countMenus +'" style="position:absolute;z-index:1;left:10px;top:'+ (i * 100) +'px;visibility:hidden;color:' +  menu.menuBorderBgColor + ';">\n'+
			'  <div id="menuLite'+ countMenus +'" style="position:absolute;z-index:1;left:'+ menu.menuBorder +'px;top:'+ menu.menuBorder +'px;visibility:hide;" onmouseout="mouseoutMenu();">\n'+
			'	 <div id="menuFg'+ countMenus +'" style="position:absolute;left:'+ menu.menuBorder +'px;top:'+ menu.menuBorder +'px;visibility:hide;">\n'+
			'';
		}
		var x=i;
		for (var i=0; i<menu.items.length; i++) {
			var item = menu.items[i];
			var childMenu = false;
			var defaultHeight = menu.fontSize+2*menu.menuItemPadding;
			if (item.label) {
				item = item.label;
				childMenu = true;
			}
			menu.menuItemHeight = menu.menuItemHeight || defaultHeight;
			var itemProps = '';
			if( menu.fontFamily != '' ) itemProps += 'font-family:' + menu.fontFamily +';';
			itemProps += 'font-weight:' + menu.fontWeight + ';fontSize:' + menu.fontSize + 'px;';
			if (menu.fontStyle) itemProps += 'font-style:' + menu.fontStyle + ';';
			if (document.all || window.mmIsOpera) 
				itemProps += 'font-size:' + menu.fontSize + 'px;" onmouseover="onMenuItemOver(null,this);" onclick="onMenuItemAction(null,this);';
			else if (!document.layers) {
				itemProps += 'font-size:' + menu.fontSize + 'px;';
			}
			var l;
			if (lrs) {
				var lw = menu.menuWidth;
				if( menu.menuItemHAlign == 'right' ) lw -= menu.menuItemPadding;
				l = new Layer(lw,body);
			}
			var itemLeft = 0;
			var itemTop = i*menu.menuItemHeight;
			if( !menu.vertical ) {
				itemLeft = i*menu.menuWidth;
				itemTop = 0;
			}
			var dTag = '<div id="menuItem'+ countItems +'" style="position:absolute;left:' + itemLeft + 'px;top:'+ itemTop +'px;'+ itemProps +'">';
			var dClose = '</div>'
			if (menu.bgImageUp) dTag = '<div id="menuItem'+ countItems +'" style="background:url('+menu.bgImageUp+');position:absolute;left:' + itemLeft + 'px;top:'+ itemTop +'px;'+ itemProps +'">';

			var left = 0, top = 0, right = 0, bottom = 0;
			left = 1 + menu.menuItemPadding + menu.menuItemIndent;
			right = left + menu.menuWidth - 2*menu.menuItemPadding - menu.menuItemIndent;
			if( menu.menuItemVAlign == 'top' ) top = menu.menuItemPadding;
			if( menu.menuItemVAlign == 'bottom' ) top = menu.menuItemHeight-menu.fontSize-1-menu.menuItemPadding;
			if( menu.menuItemVAlign == 'middle' ) top = ((menu.menuItemHeight/2)-(menu.fontSize/2)-1);
			bottom = menu.menuItemHeight - 2*menu.menuItemPadding;
			var textProps = 'position:absolute;left:' + left + 'px;top:' + top + 'px;';
			if (lrs) {
				textProps +=itemProps + 'right:' + right + ';bottom:' + bottom + ';';
				dTag = "";
				dClose = "";
			}
			
			if(document.all && !window.mmIsOpera) {
				item = '<div align="' + menu.menuItemHAlign + '">' + item + '</div>';
			} else if (lrs) {
				item = '<div style="text-align:' + menu.menuItemHAlign + ';">' + item + '</div>';
			} else {
				var hitem = null;
				if( menu.menuItemHAlign != 'left' ) {
					if(window.mmIsOpera) {
						var operaWidth = menu.menuItemHAlign == 'center' ? -(menu.menuWidth-2*menu.menuItemPadding) : (menu.menuWidth-6*menu.menuItemPadding);
						hitem = '<div id="menuItemHilite' + countItems + 'Shim" style="position:absolute;top:1px;left:' + menu.menuItemPadding + 'px;width:' + operaWidth + 'px;text-align:' 
							+ menu.menuItemHAlign + ';visibility:visible;">' + item + '</div>';
						item = '<div id="menuItemText' + countItems + 'Shim" style="position:absolute;top:1px;left:' + menu.menuItemPadding + 'px;width:' + operaWidth + 'px;text-align:' 
							+ menu.menuItemHAlign + ';visibility:visible;">' + item + '</div>';
					} else {
						hitem = '<div id="menuItemHilite' + countItems + 'Shim" style="position:absolute;top:1px;left:1px;right:-' + (left+menu.menuWidth-3*menu.menuItemPadding) + 'px;text-align:' 
							+ menu.menuItemHAlign + ';visibility:visible;">' + item + '</div>';
						item = '<div id="menuItemText' + countItems + 'Shim" style="position:absolute;top:1px;left:1px;right:-' + (left+menu.menuWidth-3*menu.menuItemPadding) + 'px;text-align:' 
							+ menu.menuItemHAlign + ';visibility:visible;">' + item + '</div>';
					}
				} else hitem = null;
			}
			if(document.all && !window.mmIsOpera) item = '<div id="menuItemShim' + countItems + '" style="position:absolute;left:0px;top:0px;">' + item + '</div>';
			var dText	= '<div id="menuItemText'+ countItems +'" style="' + textProps + 'color:'+ menu.fontColor +';">'+ item +'&nbsp</div>\n'
						+ '<div id="menuItemHilite'+ countItems +'" style="' + textProps + 'color:'+ menu.fontColorHilite +';visibility:hidden;">' 
						+ (hitem||item) +'&nbsp</div>';
			if (childMenu) content += ( dTag + dText + '<div id="childMenu'+ countItems +'" style="position:absolute;left:0px;top:3px;"><img src="'+ menu.childMenuIcon +'"></div>\n' + dClose);
			else content += ( dTag + dText + dClose);
			if (lrs) {
				l.document.open("text/html");
				l.document.writeln(content);
				l.document.close();	
				content = '';
				theStat += "-";
				tsc++;
				if (tsc > 50) {
					tsc = 0;
					theStat = "";
				}
				status = theStat;
			}
			countItems++;  
		}
		if (lrs) {
			var focusItem = new Layer(100, body);
			focusItem.visiblity="hidden";
			focusItem.document.open("text/html");
			focusItem.document.writeln("&nbsp;");
			focusItem.document.close();	
		} else {
		  content += '	  <div id="focusItem'+ countMenus +'" style="position:absolute;left:0px;top:0px;visibility:hide;" onclick="onMenuItemAction(null,this);">&nbsp;</div>\n';
		  content += '   </div>\n  </div>\n</div>\n';
		}
		i=x;
	}
	if (document.layers) {		
		container.clip.width = window.innerWidth;
		container.clip.height = window.innerHeight;
		container.onmouseout = mouseoutMenu;
		container.menuContainerBgColor = this.menuContainerBgColor;
		for (var i=0; i<container.document.layers.length; i++) {
			proto = container.menus[i];
			var menu = container.document.layers[i];
			container.menus[i].menuLayer = menu;
			container.menus[i].menuLayer.Menu = container.menus[i];
			container.menus[i].menuLayer.Menu.container = container;
			var body = menu.document.layers[0].document.layers[0];
			body.clip.width = proto.menuWidth || body.clip.width;
			body.clip.height = proto.menuHeight || body.clip.height;
			for (var n=0; n<body.document.layers.length-1; n++) {
				var l = body.document.layers[n];
				l.Menu = container.menus[i];
				l.menuHiliteBgColor = proto.menuHiliteBgColor;
				l.document.bgColor = proto.menuItemBgColor;
				l.saveColor = proto.menuItemBgColor;
				l.onmouseover = proto.onMenuItemOver;
				l.onclick = proto.onMenuItemAction;
				l.mmaction = container.menus[i].actions[n];
				l.focusItem = body.document.layers[body.document.layers.length-1];
				l.clip.width = proto.menuWidth || body.clip.width;
				l.clip.height = proto.menuItemHeight || l.clip.height;
				if (n>0) {
					if( l.Menu.vertical ) l.top = body.document.layers[n-1].top + body.document.layers[n-1].clip.height + proto.menuItemBorder + proto.menuItemSpacing;
					else l.left = body.document.layers[n-1].left + body.document.layers[n-1].clip.width + proto.menuItemBorder + proto.menuItemSpacing;
				}
				l.hilite = l.document.layers[1];
				if (proto.bgImageUp) l.background.src = proto.bgImageUp;
				l.document.layers[1].isHilite = true;
				if (l.document.layers.length > 2) {
					l.childMenu = container.menus[i].items[n].menuLayer;
					l.document.layers[2].left = l.clip.width -13;
					l.document.layers[2].top = (l.clip.height / 2) -4;
					l.document.layers[2].clip.left += 3;
					l.Menu.childMenus[l.Menu.childMenus.length] = l.childMenu;
				}
			}
			if( proto.menuBgOpaque ) body.document.bgColor = proto.bgColor;
			if( proto.vertical ) {
				body.clip.width  = l.clip.width +proto.menuBorder;
				body.clip.height = l.top + l.clip.height +proto.menuBorder;
			} else {
				body.clip.height  = l.clip.height +proto.menuBorder;
				body.clip.width = l.left + l.clip.width  +proto.menuBorder;
				if( body.clip.width > window.innerWidth ) body.clip.width = window.innerWidth;
			}
			var focusItem = body.document.layers[n];
			focusItem.clip.width = body.clip.width;
			focusItem.Menu = l.Menu;
			focusItem.top = -30;
            focusItem.captureEvents(Event.MOUSEDOWN);
            focusItem.onmousedown = onMenuItemDown;
			if( proto.menuBgOpaque ) menu.document.bgColor = proto.menuBorderBgColor;
			var lite = menu.document.layers[0];
			if( proto.menuBgOpaque ) lite.document.bgColor = proto.menuLiteBgColor;
			lite.clip.width = body.clip.width +1;
			lite.clip.height = body.clip.height +1;
			menu.clip.width = body.clip.width + (proto.menuBorder * 3) ;
			menu.clip.height = body.clip.height + (proto.menuBorder * 3);
		}
	} else {
		if ((!document.all) && (container.hasChildNodes) && !window.mmIsOpera) {
			container.innerHTML=content;
		} else {
			container.document.open("text/html");
			container.document.writeln(content);
			container.document.close();	
		}
		if (!FIND("menuLayer0")) return;
		var menuCount = 0;
		for (var x=0; x<container.menus.length; x++) {
			var menuLayer = FIND("menuLayer" + x);
			container.menus[x].menuLayer = "menuLayer" + x;
			menuLayer.Menu = container.menus[x];
			menuLayer.Menu.container = "menuLayer" + x;
			menuLayer.style.zindex = 1;
		    var s = menuLayer.style;
			s.pixeltop = -300;
			s.pixelleft = -300;
			s.top = '-300px';
			s.left = '-300px';

			var menu = container.menus[x];
			menu.menuItemWidth = menu.menuWidth || menu.menuIEWidth || 140;
			if( menu.menuBgOpaque ) menuLayer.style.backgroundColor = menu.menuBorderBgColor;
			var top = 0;
			var left = 0;
			menu.menuItemLayers = new Array();
			for (var i=0; i<container.menus[x].items.length; i++) {
				var l = FIND("menuItem" + menuCount);
				l.Menu = container.menus[x];
				l.Menu.menuItemLayers[l.Menu.menuItemLayers.length] = l;
				if (l.addEventListener || window.mmIsOpera) {
					l.style.width = menu.menuItemWidth + 'px';
					l.style.height = menu.menuItemHeight + 'px';
					l.style.pixelWidth = menu.menuItemWidth;
					l.style.pixelHeight = menu.menuItemHeight;
					l.style.top = top + 'px';
					l.style.left = left + 'px';
					if(l.addEventListener) {
						l.addEventListener("mouseover", onMenuItemOver, false);
						l.addEventListener("click", onMenuItemAction, false);
						l.addEventListener("mouseout", mouseoutMenu, false);
					}
					if( menu.menuItemHAlign != 'left' ) {
						l.hiliteShim = FIND("menuItemHilite" + menuCount + "Shim");
						l.hiliteShim.style.visibility = "inherit";
						l.textShim = FIND("menuItemText" + menuCount + "Shim");
						l.hiliteShim.style.pixelWidth = menu.menuItemWidth - 2*menu.menuItemPadding - menu.menuItemIndent;
						l.hiliteShim.style.width = l.hiliteShim.style.pixelWidth;
						l.textShim.style.pixelWidth = menu.menuItemWidth - 2*menu.menuItemPadding - menu.menuItemIndent;
						l.textShim.style.width = l.textShim.style.pixelWidth;	
					}
				} else {
					l.style.pixelWidth = menu.menuItemWidth;
					l.style.pixelHeight = menu.menuItemHeight;
					l.style.pixelTop = top;
					l.style.pixelLeft = left;
					if( menu.menuItemHAlign != 'left' ) {
						var shim = FIND("menuItemShim" + menuCount);
						shim[0].style.pixelWidth = menu.menuItemWidth - 2*menu.menuItemPadding - menu.menuItemIndent;
						shim[1].style.pixelWidth = menu.menuItemWidth - 2*menu.menuItemPadding - menu.menuItemIndent;
						shim[0].style.width = shim[0].style.pixelWidth + 'px';
						shim[1].style.width = shim[1].style.pixelWidth + 'px';
					}
				}
				if( menu.vertical ) top = top + menu.menuItemHeight+menu.menuItemBorder+menu.menuItemSpacing;
				else left = left + menu.menuItemWidth+menu.menuItemBorder+menu.menuItemSpacing;
				l.style.fontSize = menu.fontSize + 'px';
				l.style.backgroundColor = menu.menuItemBgColor;
				l.style.visibility = "inherit";
				l.saveColor = menu.menuItemBgColor;
				l.menuHiliteBgColor = menu.menuHiliteBgColor;
				l.mmaction = container.menus[x].actions[i];
				l.hilite = FIND("menuItemHilite" + menuCount);
				l.focusItem = FIND("focusItem" + x);
				l.focusItem.style.pixelTop = -30;
				l.focusItem.style.top = '-30px';
				var childItem = FIND("childMenu" + menuCount);
				if (childItem) {
					l.childMenu = container.menus[x].items[i].menuLayer;
					childItem.style.pixelLeft = menu.menuItemWidth -11;
					childItem.style.left = childItem.style.pixelLeft + 'px';
					childItem.style.pixelTop = (menu.menuItemHeight /2) -4;
					childItem.style.top = childItem.style.pixelTop + 'px';
					l.Menu.childMenus[l.Menu.childMenus.length] = l.childMenu;
				}
				l.style.cursor = "hand";
				menuCount++;
			}
			if( menu.vertical ) {
				menu.menuHeight = top-1-menu.menuItemSpacing;
				menu.menuWidth = menu.menuItemWidth;
			} else {
				menu.menuHeight = menu.menuItemHeight;
				menu.menuWidth = left-1-menu.menuItemSpacing;
			}

			var lite = FIND("menuLite" + x);
			var s = lite.style;
			s.pixelHeight = menu.menuHeight +(menu.menuBorder * 2);
			s.height = s.pixelHeight + 'px';
			s.pixelWidth = menu.menuWidth + (menu.menuBorder * 2);
			s.width = s.pixelWidth + 'px';
			if( menu.menuBgOpaque ) s.backgroundColor = menu.menuLiteBgColor;

			var body = FIND("menuFg" + x);
			s = body.style;
			s.pixelHeight = menu.menuHeight + menu.menuBorder;
			s.height = s.pixelHeight + 'px';
			s.pixelWidth = menu.menuWidth + menu.menuBorder;
			s.width = s.pixelWidth + 'px';
			if( menu.menuBgOpaque ) s.backgroundColor = menu.bgColor;

			s = menuLayer.style;
			s.pixelWidth  = menu.menuWidth + (menu.menuBorder * 4);
			s.width = s.pixelWidth + 'px';
			s.pixelHeight  = menu.menuHeight+(menu.menuBorder*4);
			s.height = s.pixelHeight + 'px';
		}
	}
	if (document.captureEvents) document.captureEvents(Event.MOUSEUP);
	if (document.addEventListener) document.addEventListener("mouseup", onMenuItemOver, false);
	if (document.layers && window.innerWidth) {
		window.onresize = NS4resize;
		window.NS4sIW = window.innerWidth;
		window.NS4sIH = window.innerHeight;
		setTimeout("NS4resize()",500);
	}
	document.onmouseup = mouseupMenu;
	window.mmWroteMenu = true;
	status = "";
}

function NS4resize() {
	if (NS4sIW != window.innerWidth || NS4sIH != window.innerHeight) window.location.reload();
}

function onMenuItemOver(e, l) {
	MM_clearTimeout();
	l = l || this;
	a = window.ActiveMenuItem;
	if (document.layers) {
		if (a) {
			a.document.bgColor = a.saveColor;
			if (a.hilite) a.hilite.visibility = "hidden";
			if (a.Menu.bgImageOver) a.background.src = a.Menu.bgImageUp;
			a.focusItem.top = -100;
			a.clicked = false;
		}
		if (l.hilite) {
			l.document.bgColor = l.menuHiliteBgColor;
			l.zIndex = 1;
			l.hilite.visibility = "inherit";
			l.hilite.zIndex = 2;
			l.document.layers[1].zIndex = 1;
			l.focusItem.zIndex = this.zIndex +2;
		}
		if (l.Menu.bgImageOver) l.background.src = l.Menu.bgImageOver;
		l.focusItem.top = this.top;
		l.focusItem.left = this.left;
		l.focusItem.clip.width = l.clip.width;
		l.focusItem.clip.height = l.clip.height;
		l.Menu.hideChildMenu(l);
	} else if (l.style && l.Menu) {
		if (a) {
			a.style.backgroundColor = a.saveColor;
			if (a.hilite) a.hilite.style.visibility = "hidden";
			if (a.hiliteShim) a.hiliteShim.style.visibility = "inherit";
			if (a.Menu.bgImageUp) a.style.background = "url(" + a.Menu.bgImageUp +")";;
		} 
		l.style.backgroundColor = l.menuHiliteBgColor;
		l.zIndex = 1;
		if (l.Menu.bgImageOver) l.style.background = "url(" + l.Menu.bgImageOver +")";
		if (l.hilite) {
			l.hilite.style.visibility = "inherit";
			if( l.hiliteShim ) l.hiliteShim.style.visibility = "visible";
		}
		l.focusItem.style.pixelTop = l.style.pixelTop;
		l.focusItem.style.top = l.focusItem.style.pixelTop + 'px';
		l.focusItem.style.pixelLeft = l.style.pixelLeft;
		l.focusItem.style.left = l.focusItem.style.pixelLeft + 'px';
		l.focusItem.style.zIndex = l.zIndex +1;
		l.Menu.hideChildMenu(l);
	} else return;
	window.ActiveMenuItem = l;
}

function onMenuItemAction(e, l) {
	l = window.ActiveMenuItem;
	if (!l) return;
	hideActiveMenus();
	if (l.mmaction) eval("" + l.mmaction);
	window.ActiveMenuItem = 0;
}

function MM_clearTimeout() {
	if (mmHideMenuTimer) clearTimeout(mmHideMenuTimer);
	mmHideMenuTimer = null;
	mmDHFlag = false;
}

function MM_startTimeout() {
	if( window.ActiveMenu ) {
		mmStart = new Date();
		mmDHFlag = true;
		mmHideMenuTimer = setTimeout("mmDoHide()", window.ActiveMenu.Menu.hideTimeout);
	}
}

function mmDoHide() {
	if (!mmDHFlag || !window.ActiveMenu) return;
	var elapsed = new Date() - mmStart;
	var timeout = window.ActiveMenu.Menu.hideTimeout;
	if (elapsed < timeout) {
		mmHideMenuTimer = setTimeout("mmDoHide()", timeout+100-elapsed);
		return;
	}
	mmDHFlag = false;
	hideActiveMenus();
	window.ActiveMenuItem = 0;
}

function MM_showMenu(menu, x, y, child, imgname) {
	if (!window.mmWroteMenu) return;
	MM_clearTimeout();
	if (menu) {
		var obj = FIND(imgname) || document.images[imgname] || document.links[imgname] || document.anchors[imgname];
		x = moveXbySlicePos (x, obj);
		y = moveYbySlicePos (y, obj);
	}
	if (document.layers) {
		if (menu) {
			var l = menu.menuLayer || menu;
			l.top = l.left = 1;
			hideActiveMenus();
			if (this.visibility) l = this;
			window.ActiveMenu = l;
		} else {
			var l = child;
		}
		if (!l) return;
		for (var i=0; i<l.layers.length; i++) { 			   
			if (!l.layers[i].isHilite) l.layers[i].visibility = "inherit";
			if (l.layers[i].document.layers.length > 0) MM_showMenu(null, "relative", "relative", l.layers[i]);
		}
		if (l.parentLayer) {
			if (x != "relative") l.parentLayer.left = x || window.pageX || 0;
			if (l.parentLayer.left + l.clip.width > window.innerWidth) l.parentLayer.left -= (l.parentLayer.left + l.clip.width - window.innerWidth);
			if (y != "relative") l.parentLayer.top = y || window.pageY || 0;
			if (l.parentLayer.isContainer) {
				l.Menu.xOffset = window.pageXOffset;
				l.Menu.yOffset = window.pageYOffset;
				l.parentLayer.clip.width = window.ActiveMenu.clip.width +2;
				l.parentLayer.clip.height = window.ActiveMenu.clip.height +2;
				if (l.parentLayer.menuContainerBgColor && l.Menu.menuBgOpaque ) l.parentLayer.document.bgColor = l.parentLayer.menuContainerBgColor;
			}
		}
		l.visibility = "inherit";
		if (l.Menu) l.Menu.container.visibility = "inherit";
	} else if (FIND("menuItem0")) {
		var l = menu.menuLayer || menu;	
		hideActiveMenus();
		if (typeof(l) == "string") l = FIND(l);
		window.ActiveMenu = l;
		var s = l.style;
		s.visibility = "inherit";
		if (x != "relative") {
			s.pixelLeft = x || (window.pageX + document.body.scrollLeft) || 0;
			s.left = s.pixelLeft + 'px';
		}
		if (y != "relative") {
			s.pixelTop = y || (window.pageY + document.body.scrollTop) || 0;
			s.top = s.pixelTop + 'px';
		}
		l.Menu.xOffset = document.body.scrollLeft;
		l.Menu.yOffset = document.body.scrollTop;
	}
	if (menu) window.activeMenus[window.activeMenus.length] = l;
	MM_clearTimeout();
}

function onMenuItemDown(e, l) {
	var a = window.ActiveMenuItem;
	if (document.layers && a) {
		a.eX = e.pageX;
		a.eY = e.pageY;
		a.clicked = true;
    }
}

function mouseupMenu(e) {
	hideMenu(true, e);
	hideActiveMenus();
	return true;
}

function getExplorerVersion() {
	var ieVers = parseFloat(navigator.appVersion);
	if( navigator.appName != 'Microsoft Internet Explorer' ) return ieVers;
	var tempVers = navigator.appVersion;
	var i = tempVers.indexOf( 'MSIE ' );
	if( i >= 0 ) {
		tempVers = tempVers.substring( i+5 );
		ieVers = parseFloat( tempVers ); 
	}
	return ieVers;
}

function mouseoutMenu() {
	if ((navigator.appName == "Microsoft Internet Explorer") && (getExplorerVersion() < 4.5))
		return true;
	hideMenu(false, false);
	return true;
}

function hideMenu(mouseup, e) {
	var a = window.ActiveMenuItem;
	if (a && document.layers) {
		a.document.bgColor = a.saveColor;
		a.focusItem.top = -30;
		if (a.hilite) a.hilite.visibility = "hidden";
		if (mouseup && a.mmaction && a.clicked && window.ActiveMenu) {
 			if (a.eX <= e.pageX+15 && a.eX >= e.pageX-15 && a.eY <= e.pageY+10 && a.eY >= e.pageY-10) {
				setTimeout('window.ActiveMenu.Menu.onMenuItemAction();', 500);
			}
		}
		a.clicked = false;
		if (a.Menu.bgImageOver) a.background.src = a.Menu.bgImageUp;
	} else if (window.ActiveMenu && FIND("menuItem0")) {
		if (a) {
			a.style.backgroundColor = a.saveColor;
			if (a.hilite) a.hilite.style.visibility = "hidden";
			if (a.hiliteShim) a.hiliteShim.style.visibility = "inherit";
			if (a.Menu.bgImageUp) a.style.background = "url(" + a.Menu.bgImageUp +")";
		}
	}
	if (!mouseup && window.ActiveMenu) {
		if (window.ActiveMenu.Menu) {
			if (window.ActiveMenu.Menu.hideOnMouseOut) MM_startTimeout();
			return(true);
		}
	}
	return(true);
}

function hideChildMenu(hcmLayer) {
	MM_clearTimeout();
	var l = hcmLayer;
	for (var i=0; i < l.Menu.childMenus.length; i++) {
		var theLayer = l.Menu.childMenus[i];
		if (document.layers) theLayer.visibility = "hidden";
		else {
			theLayer = FIND(theLayer);
			theLayer.style.visibility = "hidden";
			if( theLayer.Menu.menuItemHAlign != 'left' ) {
				for(var j = 0; j < theLayer.Menu.menuItemLayers.length; j++) {
					var itemLayer = theLayer.Menu.menuItemLayers[j];
					if(itemLayer.textShim) itemLayer.textShim.style.visibility = "inherit";
				}
			}
		}
		theLayer.Menu.hideChildMenu(theLayer);
	}
	if (l.childMenu) {
		var childMenu = l.childMenu;
		if (document.layers) {
			l.Menu.MM_showMenu(null,null,null,childMenu.layers[0]);
			childMenu.zIndex = l.parentLayer.zIndex +1;
			childMenu.top = l.Menu.menuLayer.top + l.Menu.submenuYOffset;
			if( l.Menu.vertical ) {
				if( l.Menu.submenuRelativeToItem ) childMenu.top += l.top + l.parentLayer.top;
				childMenu.left = l.parentLayer.left + l.parentLayer.clip.width - (2*l.Menu.menuBorder) + l.Menu.menuLayer.left + l.Menu.submenuXOffset;
			} else {
				childMenu.top += l.top + l.parentLayer.top;	
				if( l.Menu.submenuRelativeToItem ) childMenu.left = l.Menu.menuLayer.left + l.left + l.clip.width + (2*l.Menu.menuBorder) + l.Menu.submenuXOffset;
				else childMenu.left = l.parentLayer.left + l.parentLayer.clip.width - (2*l.Menu.menuBorder) + l.Menu.menuLayer.left + l.Menu.submenuXOffset;
			}
			if( childMenu.left < l.Menu.container.clip.left ) l.Menu.container.clip.left = childMenu.left;
			var w = childMenu.clip.width+childMenu.left-l.Menu.container.clip.left;
			if (w > l.Menu.container.clip.width)  l.Menu.container.clip.width = w;
			var h = childMenu.clip.height+childMenu.top-l.Menu.container.clip.top;
			if (h > l.Menu.container.clip.height) l.Menu.container.clip.height = h;
			l.document.layers[1].zIndex = 0;
			childMenu.visibility = "inherit";
		} else if (FIND("menuItem0")) {
			childMenu = FIND(l.childMenu);
			var menuLayer = FIND(l.Menu.menuLayer);
			var s = childMenu.style;
			s.zIndex = menuLayer.style.zIndex+1;
			if (document.all || window.mmIsOpera) {
				s.pixelTop = menuLayer.style.pixelTop + l.Menu.submenuYOffset;
				if( l.Menu.vertical ) {
					if( l.Menu.submenuRelativeToItem ) s.pixelTop += l.style.pixelTop;
					s.pixelLeft = l.style.pixelWidth + menuLayer.style.pixelLeft + l.Menu.submenuXOffset;
					s.left = s.pixelLeft + 'px';
				} else {
					s.pixelTop += l.style.pixelTop;
					if( l.Menu.submenuRelativeToItem ) s.pixelLeft = menuLayer.style.pixelLeft + l.style.pixelLeft + l.style.pixelWidth + (2*l.Menu.menuBorder) + l.Menu.submenuXOffset;
					else s.pixelLeft = (menuLayer.style.pixelWidth-4*l.Menu.menuBorder) + menuLayer.style.pixelLeft + l.Menu.submenuXOffset;
					s.left = s.pixelLeft + 'px';
				}
			} else {
				var top = parseInt(menuLayer.style.top) + l.Menu.submenuYOffset;
				var left = 0;
				if( l.Menu.vertical ) {
					if( l.Menu.submenuRelativeToItem ) top += parseInt(l.style.top);
					left = (parseInt(menuLayer.style.width)-4*l.Menu.menuBorder) + parseInt(menuLayer.style.left) + l.Menu.submenuXOffset;
				} else {
					top += parseInt(l.style.top);
					if( l.Menu.submenuRelativeToItem ) left = parseInt(menuLayer.style.left) + parseInt(l.style.left) + parseInt(l.style.width) + (2*l.Menu.menuBorder) + l.Menu.submenuXOffset;
					else left = (parseInt(menuLayer.style.width)-4*l.Menu.menuBorder) + parseInt(menuLayer.style.left) + l.Menu.submenuXOffset;
				}
				s.top = top + 'px';
				s.left = left + 'px';
			}
			childMenu.style.visibility = "inherit";
		} else return;
		window.activeMenus[window.activeMenus.length] = childMenu;
	}
}

function hideActiveMenus() {
	if (!window.activeMenus) return;
	for (var i=0; i < window.activeMenus.length; i++) {
		if (!activeMenus[i]) continue;
		if (activeMenus[i].visibility && activeMenus[i].Menu && !window.mmIsOpera) {
			activeMenus[i].visibility = "hidden";
			activeMenus[i].Menu.container.visibility = "hidden";
			activeMenus[i].Menu.container.clip.left = 0;
		} else if (activeMenus[i].style) {
			var s = activeMenus[i].style;
			s.visibility = "hidden";
			s.left = '-200px';
			s.top = '-200px';
		}
	}
	if (window.ActiveMenuItem) hideMenu(false, false);
	window.activeMenus.length = 0;
}

function moveXbySlicePos (x, img) { 
	if (!document.layers) {
		var onWindows = navigator.platform ? navigator.platform == "Win32" : false;
		var par = img;
		var lastOffset = 0;
		while(par){
			if( par.leftMargin && ! onWindows ) x += parseInt(par.leftMargin);
			if( (par.offsetLeft != lastOffset) && par.offsetLeft ) x += parseInt(par.offsetLeft);
			if( par.offsetLeft != 0 ) lastOffset = par.offsetLeft;
			par = par.offsetParent;
		}
	} else if (img.x) x += img.x;
	return x;
}

function moveYbySlicePos (y, img) {
	if(!document.layers) {
		var onWindows = navigator.platform ? navigator.platform == "Win32" : false;
		var par = img;
		var lastOffset = 0;
		while(par){
			if( par.topMargin && !onWindows ) y += parseInt(par.topMargin);
			if( (par.offsetTop != lastOffset) && par.offsetTop ) y += parseInt(par.offsetTop);
			if( par.offsetTop != 0 ) lastOffset = par.offsetTop;
			par = par.offsetParent;
		}		
	} else if (img.y >= 0) y += img.y;
	return y;
}
function goPreview(resCtrl) {
		var w = 800
		var h = 600
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
		theUrl=resCtrl;
		reWin = window.open(theUrl,'preview','toolbar=yes,location=yes,directories=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes,'+winprops);}


function goHtmlEditor(resCtrl,sType) {
			var w = 770
			var h = 550
			if (screen.width>770)
				w = 770
			if (screen.height>570)
				h = 570
			var winl = (screen.width - w) / 2;
			var wint = (screen.height - h) / 2;
			winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
			theUrl="html_editor.asp?returnArg="+resCtrl+"&sType="+sType;
			reWin = window.open(theUrl,'htmlEditor','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,'+winprops);
}

function goNewHtmlEditor(resCtrl,sType) {
			var w = 770
			var h = 550
			if (screen.width>770)
				w = 770
			if (screen.height>570)
				h = 570
			var winl = (screen.width - w) / 2;
			var wint = (screen.height - h) / 2;
			winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
			theUrl="new_html_editor.asp?returnArg="+resCtrl+"&sType="+sType;
			reWin = window.open(theUrl,'htmlEditor','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,'+winprops);
}
function goNewHtmlEditor_full(resCtrl,sType) {
			var w = 770
			var h = 550
			if (screen.width>770)
				w = 770
			if (screen.height>570)
				h = 570
			var winl = (screen.width - w) / 2;
			var wint = (screen.height - h) / 2;
			winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl
			theUrl="new_html_editor_full.asp?returnArg="+resCtrl+"&sType="+sType;
			reWin = window.open(theUrl,'htmlEditor','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,'+winprops);
}



