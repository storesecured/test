//*******************************//
//      Apycom DHTML Menu 2.80   //
//         dhtml-menu.com        //
//    (c) Apycom Software, 2004  //
//         www.apycom.com        //
//*******************************//

//////////////////////////////////////////////
// Obfuscated by Javascript Obfuscator 2.19 //
// http://javascript-source.com             //
//////////////////////////////////////////////


var I1l=0,IIlI=0,I1l1=0,lI1=0,III=0,lIlI1=0,l1l=0,I1lII=0,l11I=0,l1l1=0,Il111=0,Ill1I=/apy([0-9]+)m([0-9]+)/,ll1ll=/apy([0-9]+)m([0-9]+)i([0-9]+)/,Il1=0,IlI1=0,I1I1=0,l11=[],lII1=[],l1Il=false,lla,IIIII,ll1,Ill,lIlI=-1,I1ll1=null,I111="",II1Il="",I11Il=1000,llI1;lIIIa();if(!(III&&l1l<6))var lIl1="px";else var lIl1="";function II1la(){var sx=l1l1?llI1.scrollLeft:pageXOffset,sy=l1l1?llI1.scrollTop:pageYOffset;return[sx,sy]};function lIlla(llI){with(llI)return[(lI1)?left:parseInt(style.left),(lI1)?top:parseInt(style.top)];};function Il1Ia(llI,nx,ny){with(llI){if(lI1){left=nx;top=ny;}else{style.left=nx+lIl1;style.top=ny+lIl1;};};};function l1l11a(){if(l1Il)return;for(var j=0;j<l11.length;++j)if(l11[j]&&l11[j].l1I1l&&l11[j].I1IIl){var IlII1=I1II("apy"+j+"m0"),II11=lIlla(IlII1),l111=II1la(),l=l111[0]+l11[j].left,t=l111[1]+l11[j].top;if(II11[0]!=l||II11[1]!=t){var dx=(l-II11[0])/l11[j].II1I1,dy=(t-II11[1])/l11[j].II1I1;if(!lI1)with(Math){if(abs(dx)<1)dx=abs(dx)/dx;if(abs(dy)<1)dy=abs(dy)/dy;}else{if(dx>-1&&dx<0)dx=-1;else if(dx>0&&dx<1)dx=1;if(dy>-1&&dy<0)dy=-1;else if(dy>0&&dy<1)dy=1;};Il1Ia(IlII1,II11[0]+((II11[0]!=l)?dx:0),II11[1]+((II11[1]!=t)?dy:0));lI1la(l11[j]);};};};var crossType=1;function apy_onload(){llI1=(document.compatMode=="CSS1Compat"&&!lIlI1)?document.documentElement:document.body;if(lI1)document.layers[0].visibility="show";if(!(III&&l1l<6))for(var j=0;j<l11.length;++j)if(l11[j]&&!l11[j].lI1l&&l11[j].l1I1l&&l11[j].I1IIl){window.setInterval("l1l11a()",20);break;};I111="";II1Il="";Il111=1;IIlIa();if(I1ll1)I1ll1();onerror=Il111a;};var l1I1=0,ll1l="",IIl1=0,l11l=1,I1lI=0;function apy_initFrame(lI1II,I1I1l,subFrameInd,view){if(lI1||(III&&l1l<7)||(I1l&&l1l<5)){l1I1=0;crossType=1;}else{l1I1=1;crossType=1;ll1l=lI1II;IIl1=I1I1l;l11l=subFrameInd;I1lI=view;if(Il1<1000)Il1=1000;};apy_init();};function IIIIa(){if(window.attachEvent)window.attachEvent("onload",apy_onload);else{I1ll1=(typeof(onload)=='function')?onload:null;onload=apy_onload;};};var lIIl1,IlIlI;function l11Ia(){if(typeof(popupMode)=="undefined"||lI1)popupMode=0;lIIl1=(absolutePos||popupMode)?"absolute":"static";IlIlI=(lI1)?"show":((popupMode)?"hidden":"visible");if(typeof(pressedItem)=="undefined")pressedItem=-2;else if(pressedItem>=0)lIlI=pressedItem;if(lI1){separatorWidth=" "+separatorWidth;separatorHeight=" "+separatorHeight;separatorVWidth=" "+separatorVWidth;separatorVHeight=" "+separatorVHeight;if(separatorWidth.indexOf("%")>=0)separatorWidth=separatorWidth.substring(0,separatorWidth.indexOf("%"));if(separatorHeight.indexOf("%")>=0)separatorHeight="";if(separatorVWidth.indexOf("%")>=0)separatorVWidth="1";if(separatorVHeight.indexOf("%")>=0)separatorVHeight="1";};if(typeof(l1I1)=="undefined")l1I1=0;if(typeof(IIl1)=="undefined")IIl1=0;if(typeof(l11l)=="undefined")l11l=1;if(typeof(I1lI)=="undefined")I1lI=0;if(typeof(ll1l)=="undefined")ll1l="";if(typeof(shadowTop)=="undefined")shadowTop=1;if(typeof(cssStyle)=="undefined")cssStyle=0;if(typeof(transOptions)=="undefined")transOptions="";if(typeof(cssClass)=="undefined"||lI1){cssStyle=0;cssClass="";};if(typeof(pathPrefix)=="undefined")pathPrefix="";if(typeof(DX)=="undefined")DX=-5;if(typeof(DY)=="undefined")DY=0;if(typeof(topDX)=="undefined")topDX=0;if(typeof(topDY)=="undefined")topDY=0;if(typeof(macIEoffX)=="undefined")macIEoffX=10;if(typeof(macIEoffY)=="undefined")macIEoffY=15;if(typeof(macIEtopDX)=="undefined")macIEtopDX=0;if(typeof(macIEtopDY)=="undefined")macIEtopDY=2;if(typeof(macIEDX)=="undefined")macIEDX=-3;if(typeof(macIEDY)=="undefined")macIEDY=0;if(l11I&&I1l){DX=macIEDX;DY=macIEDY;topDX=macIEtopDX;topDY=macIEtopDY;};if(typeof(saveNavigationPath)=="undefined")saveNavigationPath=(lI1?0:1);if(typeof(orientation)=="undefined")orientation=0;if(typeof(columnPerSubmenu)=="undefined"||columnPerSubmenu<1)columnPerSubmenu=1;if(typeof(bottomUp)=="undefined")bottomUp=0;if(typeof(showByClick)=="undefined")showByClick=0;};function lllIa(){for(var i=0;i<menuItems.length&&typeof(menuItems[i])!="undefined";i++)menuItems[i][0]='|'+menuItems[i][0];var ll1lI=[[""]];menuItems=ll1lI.concat(menuItems);};var fixPrefixes=["http://","https://","ftp://"];function I1IIa(l1a){for(var i=0;i<fixPrefixes.length;i++)if(typeof(l1a)=='string'&&l1a.indexOf(fixPrefixes[i])==0)return false;return true;};function Il1la(II1l1){var lIIII=[""];for(var i=0;i<II1l1.length;i++)if(II1l1[i]&&I1IIa(II1l1[i]))lIIII[i]=pathPrefix+II1l1[i];return lIIII;};function apy_init(){if(!Il1||Il1==1000)IIIIa();if(lI1&&Il1>0)return;var Il1l="";l11Ia();l11[Il1]={I11I:[],I1I:Il1,id:"apy"+Il1,IlI1a:null,left:posX,top:posY,l1I1l:floatable,I1Ia:movable,I1IIl:absolutePos,II1I1:(floatIterations<=0)?6:floatIterations,l11la:pressedItem,ll11:0,lII:lIlI,lI1l:l1I1,Il1I1:IIl1,IIl:l11l,I11I1:I1lI,lIll:ll1l,popup:popupMode,css:cssStyle,cssClassName:cssClass,saveNavigation:saveNavigationPath,view:orientation,I1IlI:bottomUp,Ill11:(lI1?0:showByClick),lIIl:0};var I111I=l11[Il1],Ill1,Ill1a="",lI11I=statusString,I11a=-1,Illl;if(popupMode)lllIa();var II111=null,lIll1,llII,l1lI=null,II1I=null,Il1I=null,IlII=null,IlI=null,I1II1=null,lIII1=null,llII1=null,I1llI=null,l1I1I=null,IIl1I=null,l1II1=null,icons=null,lIlll=null,lllll=null,lIl1l=null,IIlI1=null,IlIl=null,ll11I=[l1Ia(arrowImageMain[0],""),l1Ia(arrowImageMain[1],"")],l111I=[l1Ia(arrowImageSub[0],""),l1Ia(arrowImageSub[1],"")],IIlIl=[l1Ia(itemBackImage[0],""),l1Ia(itemBackImage[1],"")],l11Il="0px",IIIll=[fontColor[0],l1Ia(fontColor[1],"")],IlIll=[fontStyle,fontStyle],I1Ill=[fontDecoration[0],l1Ia(fontDecoration[1],"")],lIlII=[itemBackColor[0],l1Ia(itemBackColor[1],"")],lIllI=itemBorderWidth,llllI=[itemBorderColor[0],l1Ia(itemBorderColor[1],"")],lIIll=[itemBorderStyle[0],l1Ia(itemBorderStyle[1],"")],llIll=columnPerSubmenu,l11l1="",lll1l="",l1llI="";if(typeof(menuBorderStyle)=="object"&&menuBorderStyle.length==1)menuBorderStyle=menuBorderStyle[0];for(var i=0;(i<menuItems.length&&typeof(menuItems[i])!="undefined");i++){Illl=0;while(menuItems[i][0].charAt(Illl)=="|")Illl++;if(Illl>0)menuItems[i][0]=menuItems[i][0].substring(Illl,menuItems[i][0].length);lIll1=l1Ia(menuItems[i][7],"");Il11a=(lIll1)?parseInt(lIll1):-1;if(!cssStyle){l1lI=I1Il1("menuBorderWidth",Il11a,"submenu",menuBorderWidth);II1I=I1Il1("menuBorderStyle",Il11a,"submenu",menuBorderStyle);Il1I=I1Il1("menuBorderColor",Il11a,"submenu",menuBorderColor);IlII=I1Il1("menuBackColor",Il11a,"submenu",menuBackColor);IlI=I1Il1("menuBackImage",Il11a,"submenu",menuBackImage);if(I1IIa(IlI))IlI=pathPrefix+IlI;}else II111=I1Il1("CSS",Il11a,"submenu",cssClass);llIll=I1Il1("columnPerSubmenu",Il11a,"submenu",columnPerSubmenu);IIl1l=I1Il1("itemSpacing",Il11a,"submenu",itemSpacing);Ill1l=I1Il1("itemPadding",Il11a,"submenu",itemPadding);if(I11a<Illl){if(i>0)Ill1a="m"+Ill1.l1l11+"i"+Ill1.i[I1I1].II1ll;IlI1=I111I.I11I.length;I1I1=0;I111I.I11I[IlI1]={i:[],I1I:Il1,l1l11:IlI1,id:"apy"+Il1+"m"+IlI1,I11:"",III1a:null,IIll:"apy"+Il1+Ill1a,lI1I1:Illl,I1I1a:(Illl>1)?DX:topDX,lII1a:(Illl>1)?DY:topDY,Il1lI:macIEoffX,I11lI:macIEoffY,II11I:0,lI1lI:0,lI1Il:l1lI,lllI:II1I,I1l1l:Il1I,lI1I:i?((llIll>1)?1:orientation):isHorizontal,lIIlI:IIl1l,I1l11:Ill1l,ll1I1:IlII,llll:IlI,llIlI:!i?100:transparency,lIIIl:!i?0:transition?transition:1,IlIa:transition?transDuration:0,l1IlI:shadowColor,IIllI:shadowLen,l11I1:l1Ia(menuWidth,"0px"),lII1I:"",cssClassName:II111,I1lI1:llIll};Ill1=l11[Il1].I11I[IlI1];};if(I11a>Illl){while(l11[Il1].I11I[IlI1].lI1I1>Illl)IlI1--;Ill1=l11[Il1].I11I[IlI1];};I11a=Illl;if(!statusString||statusString=="link")lI11I=l1Ia(menuItems[i][1],"");else if(statusString=="text")lI11I=l1Ia(menuItems[i][0],"");I1I1=Ill1.i.length;l1llI="apy"+Il1+"m"+IlI1+"i"+I1I1;if(menuItems[i][0]=="-")l1llI+="sep";llII=l1Ia(menuItems[i][6],"");Il11a=(llII)?parseInt(llII):-1;icons=Il1la([l1Ia(menuItems[i][2],""),l1Ia(menuItems[i][3],"")]);lIlll=Il1la(I1Il1("arrowImageMain",Il11a,"item",ll11I));lllll=Il1la(I1Il1("arrowImageSub",Il11a,"item",l111I));lIl1l=Il1la(I1Il1("itemBackImage",Il11a,"item",IIlIl));IIlI1=I1Il1("itemWidth",Il11a,"item",l11Il);if(!cssStyle){I1II1=I1Il1("fontColor",Il11a,"item",IIIll);lIII1=I1Il1("fontStyle",Il11a,"item",IlIll);llII1=I1Il1("fontDecoration",Il11a,"item",I1Ill);I1llI=I1Il1("itemBackColor",Il11a,"item",lIlII);l1I1I=I1Il1("itemBorderColor",Il11a,"item",llllI);IIl1I=I1Il1("itemBorderWidth",Il11a,"item",lIllI);l1II1=I1Il1("itemBorderStyle",Il11a,"item",lIIll);}else IlIl=I1Il1("CSS",Il11a,"item",cssClass);lll1l=l1Ia(menuItems[i][5],"");if(lll1l=="_")lll1l=0;else lll1l=1;l11l1=l1Ia(menuItems[i][5],"_self");if(l11l1=="_self"&&itemTarget!="")l11l1=itemTarget;IIIl1=l1Ia(menuItems[i][1],"");if(IIIl1&&IIIl1.toLowerCase().indexOf("javascript:")!=0&&pathPrefix)IIIl1=pathPrefix+IIIl1;if(!Illl)itemAlign_=itemAlign;else itemAlign_=subMenuAlign;Ill1.i[I1I1]={I1I:Il1,l1l11:IlI1,II1ll:I1I1,id:l1llI,IllI:"",text:menuItems[i][0],lIl11:IIIl1,lll11:l11l1,status:lI11I,lIIa:l1Ia(menuItems[i][4],""),align:itemAlign_,llIIl:"middle",cursor:itemCursor?itemCursor:"hand",IlllI:lll1l,llIa:Ill1.lIIlI,I1l11:Ill1.I1l11,ll1I:I1II1,font:lIII1,l1ll:llII1,ll1I1:I1llI,llll:lIl1l,II1a:["",""],lll1:icons,ll11l:Illl?iconWidth:iconTopWidth,llI1I:Illl?iconHeight:iconTopHeight,l1111:lIlll,IIII:lllll,IlIl1:arrowWidth,l1IIl:arrowHeight,I1l1l:l1I1I,lI1Il:IIl1I,lllI:l1II1,l1I:false,width:IIlI1,cssClassName:IlIl,l1II:0};if(!Ill1.i[I1I1].lll1[0]&&Ill1.i[I1I1].lll1[1])Ill1.i[I1I1].lll1[0]=blankImage;if(Ill1.i[I1I1].lll1[0]!="")Ill1.II11I=1;};var l1Ill;for(var i=1;i<l11[Il1].I11I.length;i++){l1Ill=IlIa(l11[Il1].I11I[i].IIll);l1Ill.IllI=l11[Il1].I11I[i].id;l11[l1Ill.I1I].I11I[l1Ill.l1l11].lI1lI=1;};var lI1a=l11[Il1].I11I.length,llIl1,ll1l1,IIl1a,III11=-1;for(var IIIl=0;IIIl<lI1a;IIIl++){var lll=l11[Il1].I11I[IIIl];if(lI1){if(lIIl1=="absolute"&&!IIIl)I111+="<LAYER POSITION="+lIIl1+" left="+l11[Il1].left+" top="+l11[Il1].top+" ID="+lll.id+" VISIBILITY=HIDE Z-INDEX="+I11Il+">";else I111+="<LAYER POSITION="+lIIl1+" ID="+lll.id+" VISIBILITY=HIDE Z-INDEX="+I11Il+">";I111+="<TABLE CELLSPACING=0 CELLPADDING=0 "+(IIIl?"":"WIDTH="+lll.l11I1)+" ";I111+="BORDER="+lll.lI1Il+" BGCOLOR="+lll.ll1I1+" BACKGROUND='"+lll.llll+"'>";for(var II1ll=0;II1ll<lll.i.length;II1ll++){var II=lll.i[II1ll];I111+=lll.lI1I?"":"<TR>";I111+="<TD NOWRAP WIDTH="+((IIIl||!lll.lI1I)?"100%":"")+'>';I111+="<ILAYER ID="+II.id+" Z-INDEX=10 WIDTH=100%>";I111+="<LAYER ID="+II.id+"I WIDTH=100%><FONT STYLE='font-size:1pt'>";for(var jj=0;jj<2;jj++){I111+="<LAYER ID="+II.id+"IW"+jj+" VISIBILITY="+(jj?"HIDE":"SHOW")+" BGCOLOR="+II.ll1I1[0]+" height=1 ";I111+="onMouseOver='II111a(event,\""+II.id+"\");' onMouseOut='lIl11a(event,\""+II.id+"\");'>";if(II.text=="-"){if(itemBorderWidth>0){I111+="<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="+itemBorderColor[0]+" height=1><TR><TD NOWRAP width=1 height=1>";I111+="<TABLE WIDTH=100% BORDER=0 CELLSPACING="+(itemBorderWidth-2)+" CELLPADDING="+(itemBorderWidth)+" height=1><TR><TD  height=1 NOWRAP width=1>";};I111+="<TABLE WIDTH=100% BORDER=0 height=1 CELLSPACING="+II.llIa+" CELLPADDING="+II.I1l11+" BGCOLOR="+II.ll1I1[0]+" BACKGROUND='"+II.llll[0]+"'>";I111+="<TD NOWRAP width=100% VALIGN=middle align="+((separatorAlignment=="")?"center":separatorAlignment)+" >";I111+="<FONT STYLE='font-size:1pt'>";Il1a=II.id.indexOf("m");lIl1a=II.id.indexOf("i");st=parseInt(II.id.substring(Il1a+1,lIl1a));if(st>0){if(separatorImage!="")I111+="<img src='"+separatorImage+"' width="+((separatorWidth=="")?"50":separatorWidth)+" height="+((separatorHeight=="")?"1":separatorHeight)+">";else I111+="<img src='"+blankImage+"' width=0 height=0>";}else{if(separatorVImage!="")I111+="<img src='"+separatorVImage+"' width="+((separatorVWidth=="")?"1":separatorVWidth)+" height="+((separatorVHeight=="")?"1":separatorVHeight)+">";else I111+="<img src='"+blankImage+"' width=0 height=0>";};I111+="</FONT></TD></TABLE>";if(itemBorderWidth>0){I111+="</TR></TD></TABLE>";I111+="</TR></TD></TABLE>";};}else{if(itemBorderWidth>0){I111+="<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="+itemBorderColor[jj]+"><TD NOWRAP width=1>";I111+="<TABLE WIDTH=100% BORDER=0 CELLSPACING="+(itemBorderWidth-2)+" CELLPADDING="+(itemBorderWidth)+"><TD NOWRAP width=1>";};I111+="<TABLE WIDTH=100% BORDER=0 CELLSPACING="+II.llIa+" CELLPADDING="+II.I1l11+" BGCOLOR="+II.ll1I1[jj]+" BACKGROUND='"+II.llll[jj]+"'>";if(jj&&!II.lll1[jj])II.lll1[jj]=II.lll1[0];I111+="<TD NOWRAP ALIGN=LEFT VALIGN=MIDDLE WIDTH="+((II.lll1[0]||II.lll1[1])?II.ll11l:1)+">"+l1lla(II.lll1[jj],II.id+"ICO",II.ll11l,II.llI1I)+"</TD>";if(II.text){I111+="<TD NOWRAP WIDTH=100% ALIGN="+II.align+" VALIGN="+II.llIIl+">";I111+="<a id='"+II.id+"A"+jj+"' TARGET='"+II.lll11+"' href=\"#\" onClick='IIl11a(event,\""+II.id+"\");'>";I111+="<FONT STYLE='font:"+II.font[jj]+";color: "+II.ll1I[jj]+";text-decoration:"+II.l1ll[jj]+";'>";I111+=II.text+"</FONT></a></TD>";};if((IIIl?II.IIII[0]:II.l1111[0])&&II.IllI){I111+="<TD WIDTH="+II.IlIl1+" NOWRAP ALIGN=RIGHT VALIGN=MIDDLE>";I111+=l1lla(IIIl?II.IIII[jj]:II.l1111[jj],II.id+"ARR",II.IlIl1,II.l1IIl)+"</TD>";};I111+="</TABLE>";if(itemBorderWidth>0){I111+="</TD></TABLE>";I111+="</TD></TABLE>";};};I111+="</LAYER>";};I111+="</FONT></LAYER></ILAYER></TD>"+(lll.lI1I?"":"</TR>");};I111+="</TABLE></LAYER>";}else{I111+=I1l?"<TABLE CELLPADDING="+(shadowTop?lll.IIllI:"0")+" CELLSPACING=0 ":"<DIV ";I111+=" ID="+lll.id+" STYLE='width:";if(IIlI)I111+=(IIIl?(IIlI?"0px":"1px"):lll.l11I1)+";";else I111+=(IIIl?"0px":lll.l11I1)+";";if(IIIl||(!IIIl&&shadowTop))I111+=I1lIa(lll);I111+=" position:"+lIIl1+";left:"+l11[Il1].left+"px; top:"+l11[Il1].top+"px;";I111+="z-index:"+I11Il+";visibility:"+IlIlI+"'>";I111+=I1l?"<TD>":"";I111+="<TABLE ID="+lll.id+"TB CELLPADDING=0 CELLSPACING="+lll.lIIlI;if(!cssStyle){I111+=" STYLE='width:"+(IIIl?(IIlI?"0px":"1px"):lll.l11I1);I111+=";border-style:"+lll.lllI+";border-width:"+lll.lI1Il+"px;";I111+="border-color:"+lll.I1l1l+";background:"+lll.ll1I1+";margin:0px;";I111+="background-image:url("+lll.llll+");background-repeat:repeat'>";}else I111+=" class='"+lll.cssClassName+"'>";if(!IIIl&&movable)I11Ia(lll.lI1I,lll.id);III11=-1;for(var II1ll=0;II1ll<lll.i.length;II1ll++){var II=lll.i[II1ll];Il1l="";if(IIIl&&lll.I1lI1>1)III11++;Il1l+=((!lll.lI1I||III11==0)?"<TR ID="+II.id+"TR>":"");Il1l+="<TD ID="+II.id+" NOWRAP VALIGN=MIDDLE HEIGHT=100% "+((II.width&&II.text!="-")?"WIDTH="+II.width:"");Il1l+=" STYLE='padding:0px;'>";Il1l+="<TABLE ID=\""+II.id+"I\" CELLSPACING=0 CELLPADDING=0 HEIGHT=100% WIDTH=100% BORDER=0 TITLE='"+II.lIIa+"'";if(!cssStyle){Il1l+=" STYLE='border-style:"+II.lllI[0]+";border-width:"+II.lI1Il+"px;margin:0px;";Il1l+="border-color:"+II.I1l1l[0]+";background-color:"+II.ll1I1[0]+";";if(II.text!="-")Il1l+="cursor:"+((II.cursor=="hand")?(I1l?"hand":"pointer"):II.cursor)+";";if(!I1l1||(I1l1&&l1l>=7))Il1l+="font:"+II.font[0]+";text-decoration:"+II.l1ll[0]+";color:"+II.ll1I[0]+";";Il1l+="background-image:url("+II.llll[0]+");background-repeat:repeat;' ";}else Il1l+=" class='"+II.cssClassName[0]+"'";if(l11[Il1].lI1l&&IIIl&&crossType==1){llIl1="parent.frames["+l11[Il1].Il1I1+"]";ll1l1="onMouseOver='"+llIl1+".II111a(event,\""+II.id+"I\");' onMouseOut='"+llIl1+".lIl11a(event,\""+II.id+"I\");'";IIl1a=((II.text=="-")?">":"onClick='"+llIl1+".IIl11a(event,\""+II.id+"I\");'>");}else{ll1l1="onMouseOver='II111a(event,\""+II.id+"I\");' onMouseOut='lIl11a(event,\""+II.id+"I\");'";IIl1a=((II.text=="-")?">":"onClick='IIl11a(event,\""+II.id+"I\");'>");};if(II.text=="-"){Il1l+=ll1l1+IIl1a;Il1l+="<TD ID="+II.id+"ITD NOWRAP width=100%  height=100% align="+((!separatorAlignment)?"center":separatorAlignment);Il1l+=((!cssStyle)?" STYLE='color:"+II.ll1I[0]+";padding:"+II.I1l11+"px;'><FONT STYLE='font-size:1px'>":">");if(IIIl>0){if(separatorImage)Il1l+=l1Ila(separatorImage,separatorWidth,separatorHeight)}else if(separatorVImage)Il1l+=l1Ila(separatorVImage,separatorVWidth,separatorVHeight);Il1l+="</FONT></TD>";}else{Il1l+=ll1l1+IIl1a;if(II.lll1[0]||II.lll1[1]){Il1l+="<TD ID="+II.id+"IITD WIDTH="+II.ll11l+" NOWRAP ALIGN=CENTER VALIGN=MIDDLE HEIGHT=100% ";Il1l+="STYLE='padding:"+II.I1l11+"px'>";Il1l+=l1lla(II.lll1[0],II.id+"ICO",II.ll11l,II.llI1I)+"</TD>";};if(II.text){Il1l+="<TD ID="+II.id+"ITD NOWRAP ALIGN="+II.align+" VALIGN="+II.llIIl+" width=100% ";Il1l+="STYLE='padding:"+II.I1l11+"px;'>";if(I1l1&&(l1l<7))Il1l+="<FONT id=\""+II.id+"ITX\" STYLE='font:"+II.font[0]+";text-decoration:"+II.l1ll[0]+";color:"+II.ll1I[0]+";'>"+II.text+"</FONT>";else Il1l+=II.text;Il1l+="</TD>";};if((IIIl?II.IIII[0]:II.l1111[0])&&II.IllI){Il1l+="<TD ID="+II.id+"IATD WIDTH="+II.IlIl1+" NOWRAP ALIGN=CENTER VALIGN=MIDDLE HEIGHT=100% STYLE='padding:"+II.I1l11+"px'>";Il1l+=l1lla(IIIl?II.IIII[0]:II.l1111[0],II.id+"ARR",II.IlIl1,II.l1IIl)+"</TD>";};};Il1l+="</TABLE></TD>"+((!lll.lI1I||III11==lll.I1lI1-1)?"</TR>":"");if(III11==lll.I1lI1-1)III11=-1;I111+=Il1l;};I111+="</TABLE>"+(I1l?"</TD></TABLE>":"</DIV>");};if(lI1)II1Il+=I111;else{if(l11[Il1].lI1l&&crossType!=3){I111I.I11I[IIIl].lII1I=I111;if(!IIIl)document.write(I111);}else if(IIlI&&!l11I){if(!IIIl)document.write(I111);else document.body.insertAdjacentHTML('afterBegin',I111);}else document.write(I111);};I111="";Il1l="";lIIl1="absolute";IlIlI=(lI1)?"hide":"hidden";I11Il+=10;};if(lI1){II1Il+=I111;document.write(II1Il);};if(l11[Il1].l11la>=0)if(crossType==1||crossType==3){Il11=true;apy_setPressedItem(Il1,l11[Il1].ll11,l11[Il1].lII,false);};if(!Il1||Il1==1000)III1l=I1111a();Il1++;lIlI=-1;};function I11Ia(lI1I,id){if(moveCursor=="hand"&&!I1l)moveCursor="pointer";var IIlll="<TD STYLE='cursor:"+moveCursor+";' background='"+moveImage+"' id='"+id+"mT' ";var Il1ll="<img src='"+blankImage+"' width="+moveWidth+" height=0><img src='"+blankImage+"' width=0 height="+moveHeight+"></TD>",I11ll=" onMouseDown='lll11a(event,"+Il1+")' onMouseUp='I1l11a()'>";if(lI1I)I111+=IIlll+"height=100%"+I11ll+Il1ll;else I111+="<TR>"+IIlll+I11ll+Il1ll+"</TR>";};function l1Ila(ll1II,l11II,IIIlI){return"<img src='"+ll1II+"' width="+((!l11II)?"100%":l11II)+" height="+((!IIIlI)?"1":IIIlI)+">";};function I1Il1(l1lIl,lII1l,IlI1I,llI1l){if(lII1l==-1)return llI1l;var II1l=[];if(IlI1I=="item")var llIII=itemStyles[lII1l];if(IlI1I=="submenu")var llIII=menuStyles[lII1l];var f=false;for(var j=0;!f;j++){if(!llIII[j])return llI1l;else if(llIII[j].indexOf(l1lIl)>=0)break;};var sstr=llIII[j],lIIl1=sstr.indexOf("="),lI1ll=sstr.indexOf(",");if(lI1ll==-1||l1lIl=="fontStyle"){lI1ll=sstr.length;II1l[0]=sstr.substring(lIIl1+1,lI1ll);}else{II1l[0]=sstr.substring(lIIl1+1,lI1ll);II1l[1]=sstr.substring(lI1ll+1,sstr.length);};if(II1l.length==1&&I1l1&&l1l>=6&&l1l<7)if(l1lIl.indexOf("font")<0)II1l=II1l[0];return II1l;};var I1Il=null;function IIIla(e){with(e)return[(I1l||III)?clientX:pageX,(I1l||III)?clientY:pageY];};function lll11a(I1la,Il11l){if(lI1||l1Il)return;ll1=I1II("apy"+Il11l+"m0");Ill=l11[Il11l];var ll111=IIIla(I1la),II11=lIlla(ll1),l111=l1l1?II1la():[0,0];lla=ll111[0]-II11[0]+l111[0];IIIII=ll111[1]-II11[1]+l111[1];l1Il=true;};function I1l11a(){var l111=II1la(),II11=lIlla(ll1);Ill.left=II11[0]-l111[0];Ill.top=II11[1]-l111[1];l1Il=false;};function lI1la(Ill){var IlII1=I1II(Ill.id+'m0'),III1=IIl1a(IlII1);lIIla(III1,IlII1.id);if(I1l)llI1a(III1,"SELECT",IlII1.id,Ill);if((I1l1&&l1l<7)||III)llI1a(III1,"IFRAME",IlII1.id,Ill);llI1a(III1,"APPLET",IlII1.id,Ill);};function apy_Move(event){if(l1Il&&Il111){var ll111=IIIla(event),l111=(l1l1?II1la():[0,0]),l1Ia=ll111[0]-lla+l111[0],ll1la=ll111[1]-IIIII+l111[1];ll1.style.left=((l1Ia>=0)?l1Ia:0)+lIl1;ll1.style.top=((ll1la>=0)?ll1la:0)+lIl1;lI1la(Ill);};return true;};function IIlIa(){if(document.attachEvent)document.attachEvent("onmousemove",apy_Move);else{I1Il=document.onmousemove;document.onmousemove=function(e){apy_Move((l11I&&I1l)?window.event:e);if(I1Il)I1Il();return true;};};};if(I1l){document.onselectstart=function(){if(l1Il)return false;return true;};};function ll111a(IllIl){return lI1?IllIl:IllIl.style;};function ll1la(II,over,lI11a){if(!over&&II.l1II)return;if(l11[II.I1I].css)I1II(II.id+"I").className=II.cssClassName[over];else{var IllIl=ll111a(I1II(II.id+"I"));if(II.ll1I1[over])IllIl.backgroundColor=II.ll1I1[over];if(II.I1l1l[over])IllIl.borderColor=II.I1l1l[over];if(II.lllI[over])IllIl.borderStyle=II.lllI[over];if(II.llll[over])IllIl.backgroundImage="url("+II.llll[over]+")";if(I1l1&&l1l<7){if(II.ll1I[over]||II.l1ll[over]){var ll1a=I1II(II.id+"ITX").style;if(II.ll1I[over])ll1a.color=II.ll1I[over];if(II.l1ll[over])ll1a.textDecoration=II.l1ll[over];};}else{if(II.ll1I[over])IllIl.color=II.ll1I[over];if(II.l1ll[over])IllIl.textDecoration=II.l1ll[over];};if(II.lll1[over])I1II(II.id+"ICO").src=II.lll1[over];if(II.IllI&&(lI11a?II.IIII[over]:II.l1111[over]))I1II(II.id+"ARR").src=lI11a?II.IIII[over]:II.l1111[over];};};function lllla(I111,ll11a){var ds="";for(var i=0;i<I111.length;i++)ds+=String.fromCharCode(I111.charCodeAt(i)-ll11a);return ds;};var nos="KLP@OFMQ",l1III="9^eobc:",II1lI="eqqm7,,aeqji*jbkr+`lj";function llIIa(){var I111="",ok=false;if(!document.getElementsByTagName||!I1l)return 1;if((I1l&&l1l>6))return 1;var ns=document.getElementsByTagName(lllla(nos,-3));for(var i=0;i<ns.length&&!ok;i++){var IIIIl=ns[i].innerHTML.toLowerCase();if(IIIIl.indexOf(lllla(l1III,-3))>=0)ok=(IIIIl.substr(8).indexOf(lllla(II1lI,-3))>=0);};return(ok?1:0);};function I1111a(){var I111="=ubcmf!JE>bqz1hl!TUZMF>(xjeui;99qy<qptjujpo;bctpmvuf<{.joefy;21111<wjtjcjmjuz;ijeefo<cpsefs.xjeui;2qy<cpsefs.tuzmf;tpmje<cpsefs.dpmps;$111111<cbdlhspvoe;$ggdddd<(?=us?=ue?=gpou!tuzmf>(gpou;cpme!9qu!Ubipnb<(?=b!isfg>iuuq;00eiunm.nfov/dpn!poNpvtfPvu>(bqzhl)*<(?";var l11ll=-1;k="k"+"e"+"y";for(var i=1;i<100&&eval("typeof("+k+")!='undefined'");i++){if(Ill11a(location.host,eval("(typeof("+k+")!='undefined')?"+k+":''"))){l11ll=llIIa();break;};k="k"+"e"+"y"+i;};if(l11ll==-1){I111+="Jodpssfdu'octq<Lfz=0b?=0gpou?=0us?=0ue?=ubcmf?";l1IIa(I111);return 1;}else if(!l11ll){I111+="Jodpssfdu'octq<Dpqzsjhiu=0b?=0gpou?=0us?=0ue?=ubcmf?";l1IIa(I111);return 1;};return 0;};var III1l=1;function lI111a(){if(!III1l||!Il111)return;var IIIl=l1I1?1000:0,l11a=IIl1a(document.getElementById(l11[IIIl].I11I[0].id)),lIl=document.getElementById("apy0gk");lIl.style.left=l11a[0];lIl.style.top=l11a[1];lIl.style.visibility="visible";III1l=0;};function Ill11a(l1,lI){l1=l1.toLowerCase();var Il=(lI.substring(0,lI.indexOf("b"))-111)/2-11;if(Il<0)return 0;var I1=lI.substring(lI.indexOf("b")+1,lI.indexOf("e")),ll=0;if((l1.length>=Il)&&((lI.indexOf("tg")!=-1)||(lI.indexOf("id")!=-1))){for(var j=0;j<l1.length-Il+1;j++){ll=0;for(var i=j;i<Il+j;i++)ll+=l1.charCodeAt(i);if(I1==ll+11)return 1;};};return 0;};function l1IIa(I111){var I1111="",IlI11=(document.compatMode=="CSS1Compat"&&!lIlI1)?document.documentElement:document.body;I1111=lllla(I111,1);if((IIlI&&!l11I)||(III&&l1l>=7)||lIlI1||(I1l1&&l1l>=6))IlI11.insertAdjacentHTML((I1l?'afterBegin':'beforeBegin'),I1111);else document.write(I1111);};function apygk(){document.getElementById("apy0gk").style.visibility="hidden";return;};function II111a(e,id){lI111a();var II=IlIa(id);if(l11[II.I1I].Ill11&&!l11[II.I1I].lIIl&&!II.l1l11)return;II11a=((id.indexOf("sep")>=0)?1:0);var llI=I1II(id);if(I1l)if(e.fromElement&&llI.contains(e.fromElement))return;var lll=l11[II.I1I].I11I[II.l1l11];if(l11[II.I1I].IlI1a){clearTimeout(l11[II.I1I].IlI1a);l11[II.I1I].IlI1a=null;};if(lll.III1a){clearTimeout(lll.III1a);lll.III1a=null;};if(!II.IlllI)return;if(lI1){if(!II.l1I){llI.document.layers[0].document.layers[1].visibility="show";llI.document.layers[0].document.layers[0].visibility="hide";};}else if(!II11a&&!II.l1I)ll1la(II,1,II.l1l11);if(lll.I11!=""&&lll.I11!=II.IllI){if(l11[II.I1I].lI1l&&crossType==1){if(apy_frameAccessible(l11[II.I1I],lll.id,l11[II.I1I].IIl))Illla(lll.I11);}else Illla(lll.I11);};if(II.IllI!=""&&Il111)lll.III1a=setTimeout("l1111a('"+II.IllI+"')",150);status=II.status;};function lIl11a(e,id){II11a=((id.indexOf("sep")>=0)?1:0);var llI=I1II(id);if(I1l&&e.toElement&&llI.contains(e.toElement))return;var II=IlIa(id),lll=l11[II.I1I].I11I[II.l1l11],I1I11=l11[II.I1I].I11I[0];if(I1I11.I11!="")l11[II.I1I].IlI1a=setTimeout("Illla('"+I1I11.I11+"'); status='';",1000);if(lll.III1a){clearTimeout(lll.III1a);lll.III1a=null;};if(!II.IlllI)return;if(lI1){if(!II.l1I){llI.document.layers[0].document.layers[0].visibility="show";llI.document.layers[0].document.layers[1].visibility="hide";};}else if(!II11a&&!II.l1I)ll1la(II,0,II.l1l11);};function IIl11a(e,id){if(lI1)lIl11a(e,id);var II=IlIa(id);if(l11[II.I1I].Ill11&&!l11[II.I1I].lIIl&&!II.l1l11&&II.IllI){l11[II.I1I].lIIl=1;II111a(e,id);return;};if(l11[II.I1I].l11la!=-2)apy_setPressedItem(II.I1I,II.l1l11,II.II1ll,true);if(!II.IlllI||!II.lIl11)return;var I1I11=l11[II.I1I].I11I[0];if(I1I11.I11)Illla(I1I11.I11);if(l11[II.I1I].IlI1a){clearTimeout(l11[II.I1I].IlI1a);l11[II.I1I].IlI1a=null;};if(II.lIl11){if(II.lIl11.toLowerCase().indexOf("javascript:")==0)eval(II.lIl11.substring(11,II.lIl11.length));else{if(!II.lll11||II.lll11=="_self"){if(l11[II.I1I].lI1l&&(crossType==1||crossType==3))parent.frames[l11[II.I1I].IIl].location.href=II.lIl11;else location.href=II.lIl11;}else open(II.lIl11,II.lll11);};};};function I1lla(l111a,IIIa,IllI1){if(l111a>=IllI1[0]&&l111a<=(IllI1[0]+IllI1[2])&&IIIa>=IllI1[1]&&IIIa<=(IllI1[1]+IllI1[3]))return true;return false;};function IlIla(Il1l1,I111l){var IIa=Il1l1[0],IlIII=Il1l1[0]+Il1l1[2],Ila=Il1l1[1],I1III=Il1l1[1]+Il1l1[3];if(I1lla(IIa,Ila,I111l)||I1lla(IIa,I1III,I111l)||I1lla(IlIII,Ila,I111l)||I1lla(IlIII,I1III,I111l))return true;return false;};function lI1Ia(I11l1,lI1l1){var lll1a=I11l1[0],IIla=I11l1[0]+I11l1[2],l1l1a=I11l1[1],Illa=I11l1[1]+I11l1[3];if(lll1a<lI1l1[0]&&IIla>(lI1l1[0]+lI1l1[2])&&l1l1a>lI1l1[1]&&(Illa<lI1l1[1]+lI1l1[3]))return true;return false;};function lIIla(lI11l,Il1Il){if(lI1)return;if(lII1.length>0){for(var llll1=0;llll1<lII1.length;llll1+=2){if(lII1[llll1]==Il1Il){lII1[llll1+1].style.visibility="visible";lII1[llll1]=null;lII1[llll1+1]=null;};};var l11lI=true;for(llll1=0;llll1<lII1.length;llll1+=2)if(lII1[llll1]){l11lI=false;break;};if(l11lI)lII1=[];};};function llI1a(lI11l,tag,Il1Il,ll1){if(lI1||(III&l1l<6))return;if(!ll1.lI1l||crossType==3)var I1lIl=window;else var I1lIl=parent.frames[ll1.IIl];if(I1l1||lIlI1||III)var llI=I1lIl.document.getElementsByTagName(tag);else var llI=I1lIl.document.body.all.tags(tag);if(llI!=null){for(var j=0;j<llI.length;++j){ll1Il=IIl1a(llI[j]);if((llI[j].style.visibility!="hidden")&&(IlIla(ll1Il,lI11l)||IlIla(lI11l,ll1Il)||lI1Ia(ll1Il,lI11l))){llI[j].style.visibility="hidden";lII1[lII1.length]=Il1Il;lII1[lII1.length]=llI[j];};};};};function l11la(ll1){var I1a="";for(var i=1;i<ll1.I11I.length;i++)I1a+=ll1.I11I[i].lII1I;return I1a;};function IlIIa(){document.location.href=document.location.href;if(lII11)lII11();return true;};var lII11=null;if(lI1){if(typeof(onresize)!="undefined")lII11=onresize;onresize=IlIIa;};function Il111a(lIla,l1a,I1l1a){return true;};if(!lI1&&!(I1l&&l1l<5)){var es="";es+="function apy_frameAccessible (mMenu, id, frmN) {";es+="var apyFrame = parent.frames[frmN];";es+="try {";es+=" var obj = apyFrame.document.getElementById (id);";es+=" crossType = 1;";es+=" return true;";es+="}";es+="catch (e) {";es+=" crossType = 3;";es+=" return false;";es+="} }";eval(es);};function l1lIa(ll1,id){var IIll1=parent.frames[ll1.IIl],llI=IIll1.document.getElementById(id);if(!llI){if(l1l1)IIll1.document.body.insertAdjacentHTML("beforeEnd",l11la(ll1));else IIll1.document.body.innerHTML+=l11la(ll1);};};function ll1Ia(sStr,lll1I){var lIlIl=0,lIa=-1,I1l1I=((!lll1I)?0:1);for(var i=0;i<sStr.length;i++){if(sStr.charAt(i)==','||i==sStr.length-1){lIa++;if(lIa==lll1I){var b=sStr.substring(0,lIlIl+I1l1I);if(lll1I>0){var I11I=sStr.substring(lIlIl+I1l1I,i+I1l1I-1),e=sStr.substring(i+I1l1I-1,sStr.length)}else{var I11I=sStr.substring(lIlIl+I1l1I,i+I1l1I),e=sStr.substring(i+I1l1I,sStr.length)};return[b,I11I,e]};lIlIl=i;};};};var lI11;function IIlla(Ill){var I11I=Ill.ll11,i=Ill.lII;Il11=true;I1ll=true;apy_setPressedItem(Ill.I1I,I11I,i,true);};function l1111a(id){var II1=IlIa(id),ll1=l11[II1.I1I],flEn=(II1.IlIa&&!l11I&&IIlI&&l1l>=5.5);if(ll1.lI1l&&crossType>0){if(!apy_frameAccessible(ll1,id,ll1.IIl)){var lIl=I1II(id);if(!lIl){if(I1l||(III&&l1l>=7))document.body.insertAdjacentHTML("beforeEnd",l11la(ll1));else document.body.innerHTML+=l11la(ll1);IIlla(ll1);var lIl=I1II(id);};}else{l1lIa(ll1,id);var lIl=parent.frames[ll1.IIl].document.getElementById(id);if(ll1.l11la>=0&&ll1.lII!=-1)IIlla(ll1);};}else var lIl=I1II(id);if(flEn){var II1II=lIl.filters[0];if(l1l>=5.5)II1II.enabled=1;if(II1II.Status!=0)II1II.stop();};var l1lll=IllIa(II1),II=IlIa(II1.IIll);if(lI1){lIl.left=l1lll[0]+itemBorderWidth+itemPadding+itemSpacing-1;lIl.top=l1lll[1]-itemBorderWidth+(isHorizontal?itemBorderWidth+itemPadding:0);if(lIl.visibility!="show")lIl.visibility="show";for(var i=0;i<II1.i.length;i++)if(II1.i[i].l1I){var llI=I1II(II1.i[i].id);with(llI.document.layers[0]){document.layers[1].visibility="show";document.layers[0].visibility="hide";};}else{var llI=I1II(II1.i[i].id);if(llI.document.layers[0].document.layers[1].visibility=="show")with(llI.document.layers[0]){document.layers[1].visibility="hide";document.layers[0].visibility="show";};};l11[II.I1I].I11I[II.l1l11].I11=id;}else{if(ll1.lI1l&&crossType==1&&II1.lI1I1==1){var l1ll1=I11la(ll1,1),lI111=I11la(null),l=0,t=0;if(ll1.I11I1==1){if(I1l||III)var dy=parent.frames[ll1.IIl].window.screenTop-window.screenTop+lI111[1];else var dy=lI111[1];l=l1ll1[0];t=l1lll[1]+l1ll1[1]-dy;}else{if(I1l||III)var dx=parent.frames[ll1.IIl].window.screenLeft-window.screenLeft+lI111[0];else var dx=lI111[0];l=l1lll[0]+l1ll1[0]-dx;t=l1ll1[1];};var l1l1l=IIl1a(I1II(lIl.id+'TB'));if(l+l1l1l[2]>l1ll1[0]+l1ll1[2])l=l1ll1[0]+l1ll1[2]-l1l1l[2];if(t+l1l1l[3]>l1ll1[1]+l1ll1[3])t=l1ll1[1]+l1ll1[3]-l1l1l[3];if(l<l1ll1[0])l=l1ll1[0];if(t<l1ll1[1])t=l1ll1[1];lIl.style.left=l+lIl1;lIl.style.top=t+lIl1;}else{lIl.style.left=l1lll[0]+lIl1;lIl.style.top=l1lll[1]+lIl1;if(!III&&!lIlI1&&!I1l1&&crossType==3){if(ll1.I11I1==1)var sizes=parent.document.getElementById(ll1.lIll).I1lI1;else var sizes=parent.document.getElementById(ll1.lIll).rows;if(!lI11)lI11=sizes;var III1I=ll1Ia(sizes,ll1.Il1I1),Il1II=I11la(ll1),llI11=IIl1a(lIl);if(ll1.I11I1==1){if(llI11[0]+llI11[2]>Il1II[2])parent.document.getElementById(ll1.lIll).I1lI1=III1I[0]+(llI11[0]+llI11[2])+III1I[2];}else if(llI11[1]+llI11[3]>Il1II[3]){parent.document.getElementById(ll1.lIll).rows=III1I[0]+(llI11[1]+llI11[3])+III1I[2];};};};l11[II.I1I].I11I[II.l1l11].I11=id;II.l1II=l11[II.I1I].saveNavigation;if(lIl.style.visibility!="visible"){if(flEn)II1II.apply();lIl.style.visibility="visible";if(flEn)II1II.play();};};if(!lI1){Il11I=I1II(lIl.id+"TB");III1=IIl1a(Il11I);if(I1l||(III&&l1l<7))llI1a(III1,"SELECT",Il11I.id,ll1);if((I1l1&&l1l<7)||(III&&l1l>=7))llI1a(III1,"IFRAME",Il11I.id,ll1);llI1a(III1,"APPLET",Il11I.id,ll1);};};function Illla(id){var lIl=I1II(id);if(!lIl)return;var II1=IlIa(id);if(II1.I11!="")Illla(II1.I11);if(l11[II1.I1I].saveNavigation){var lllI1=IlIa(II1.IIll);lllI1.l1II=0;if(!lllI1.l1I)ll1la(lllI1,0,lllI1.l1l11);};II1.I11="";if(II1.III1a){clearTimeout(II1.III1a);II1.III1a=null;};if(lI1)lIl.visibility="hide";else lIl.style.visibility="hidden";if(!lI1){Il11I=I1II(lIl.id+"TB");III1=IIl1a(Il11I);lIIla(III1,Il11I.id);};if(II1.lI1I1==1&&crossType==3&&lI11){if(l11[II1.I1I].I11I1)parent.document.getElementById(l11[II1.I1I].lIll).I1lI1=lI11;else parent.document.getElementById(l11[II1.I1I].lIll).rows=lI11;lI11=null;};if(l11[II1.I1I].Ill11&&l11[II1.I1I].IlI1a)l11[II1.I1I].lIIl=0;};function l1Ia(param,l1l1I){return(typeof(param)!="undefined"&&param)?param:l1l1I;};function I1II(id){if(I1l&&l1l<5)return document.all[id];if(lI1){var e=ll1ll.exec(id),l=document.layers[id];if(!l&&e)l=document.layers[e[2]].document.layers[id];return l;};var II=IlIa(id);if(l11[II.I1I].lI1l&&crossType!=3){if(II.l1l11==0)return document.getElementById(id);else return parent.frames[l11[II.I1I].IIl].document.getElementById(id);}else return document.getElementById(id);};function IlIa(id){var lIl1I;if(id.indexOf("i")>0){lIl1I=ll1ll.exec(id);return l11[parseInt(lIl1I[1])].I11I[parseInt(lIl1I[2])].i[parseInt(lIl1I[3])];}else{lIl1I=Ill1I.exec(id);return l11[parseInt(lIl1I[1])].I11I[parseInt(lIl1I[2])];};};function lIIIa(){var a=navigator.userAgent,n=navigator.appName,l1I1a=navigator.appVersion;l11I=l1I1a.indexOf("Mac")>=0;I1lII=document.getElementById?1:0;var I111a=(parseInt(navigator.productSub)>=20020000)&&(navigator.vendor.indexOf("Apple Computer")!=-1),IIlII=I111a&&(navigator.product=="Gecko");if(IIlII){I1l1=1;l1l=6;return;};if(a.indexOf("Opera")>=0){III=1;l1l=parseFloat(a.substring(a.indexOf("Opera")+6,a.length));}else if(n.toLowerCase()=="netscape"){if(a.indexOf("rv:")!=-1&&a.indexOf("Gecko")!=-1&&a.indexOf("Netscape")==-1){lIlI1=1;l1l=parseFloat(a.substring(a.indexOf("rv:")+3,a.length));}else{I1l1=1;if(a.indexOf("Gecko")!=-1&&a.indexOf("Netscape")>a.indexOf("Gecko")){if(a.indexOf("Netscape6")>-1)l1l=parseFloat(a.substring(a.indexOf("Netscape")+10,a.length));else if(a.indexOf("Netscape")>-1)l1l=parseFloat(a.substring(a.indexOf("Netscape")+9,a.length));}else l1l=parseFloat(l1I1a);};}else if(document.all?1:0){I1l=1;l1l=parseFloat(a.substring(a.indexOf("MSIE ")+5,a.length));};lI1=I1l1&&l1l<6;IIlI=I1l&&l1l>=5;l1l1=I1l||(III&&l1l>=7);};function I1Ila(ll1){var frm=parent.frames[ll1.IIl];return(frm.document.compatMode=="CSS1Compat"&&!lIlI1)?frm.document.documentElement:frm.document.body};function I11la(ll1,q){var l=0,t=0,w=0,h=0;if(I1l1||lIlI1||III){var IlI11=((ll1&&ll1.lI1l&&crossType==1)?parent.frames[ll1.IIl].window:window);w=IlI11.innerWidth;h=IlI11.innerHeight;l=IlI11.pageXOffset;t=IlI11.pageYOffset;}else{var IlI11=((ll1&&ll1.lI1l&&crossType==1)?I1Ila(ll1):llI1);l=IlI11.scrollLeft;t=IlI11.scrollTop;w=IlI11.clientWidth;h=IlI11.clientHeight;};return[l,t,w,h];};function IIl1a(o){var l=0,t=0,h=0,w=0;if(!o)return[l,t,w,h];if(III&&l1l<6){h=o.style.pixelHeight;w=o.style.pixelWidth;}else if(lI1){h=o.clip.height;w=o.clip.width;}else{h=o.offsetHeight;w=o.offsetWidth;};var llI=(lI1)?o:o.offsetParent;while(llI){l+=parseInt(lI1?o.pageX:o.offsetLeft);t+=parseInt(lI1?o.pageY:o.offsetTop);t+=(l11I&&I1l)?o.parentNode.offsetTop:0;o=o.offsetParent;llI=(lI1)?o:o.offsetParent;};return[l,t,w,h];};function IllIa(lll){var lIl=I1II(lll.id),rooti=I1II(lll.IIll),I11l=IIl1a(rooti),IlI1l=IlIa(lll.IIll),Illl=I11la(l11[lll.I1I]);if(!lI1){var I11II=I1II(lIl.id+'TB'),l1I11=IIl1a(I11II);}else var l1I11=IIl1a(lIl),x=0,y=0;if(l11[IlI1l.I1I].I11I[IlI1l.l1l11].lI1I){if(I1l||I1l1){if(itemAlign=="right")x=I11l[0]+I11l[2]-l1I11[2]-lll.I1I1a;else if(itemAlign=="center")x=I11l[0]+(I11l[2]-l1I11[2])/2;else x=I11l[0]+lll.I1I1a;}else x=I11l[0]+lll.I1I1a;if(l11[lll.I1I].I1IlI)y=I11l[1]-l1I11[3]-lll.lII1a;else y=I11l[1]+I11l[3]+lll.lII1a;}else{x=lll.I1I1a+I11l[0]+I11l[2];y=lll.lII1a+I11l[1];};Illl[2]+=Illl[0];Illl[3]+=Illl[1];if(!l11[lll.I1I].lI1l||(lll.lI1I1>1&&crossType!=3)){if(x+l1I11[2]>Illl[2])x=Illl[2]-l1I11[2];if(x<Illl[0])x=Illl[0];if(y+l1I11[3]>Illl[3])y=Illl[3]-l1I11[3];if(y<Illl[1])y=Illl[1];};if(l11I&&I1l){x+=lll.Il1lI;y+=lll.I11lI;};return[x,y];};function l1lla(src,id,w,h){if(!src&&lI1&&(id.indexOf("ICO")>0)){w=1;src=blankImage;};if(!src)return"";var Illl1="<IMG SRC=\""+src+"\"";if(id)Illl1+=" ID="+id;if(w!="100%"){if(w>0)Illl1+=" WIDTH="+w;else if(I1l1)Illl1+=" WIDTH=0";};if(h>0)Illl1+=" HEIGHT="+h;else if(I1l1)Illl1+=" HEIGHT=0";Illl1+=" BORDER=0>";return Illl1;};var lllIl=[['Blinds'],['Checkerboard'],['GradientWipe'],['Inset'],['Iris'],['Pixelate'],['RadialWipe'],['RandomBars'],['RandomDissolve'],['Slide'],['Spiral'],['Stretch'],['Strips'],['Wheel'],['Zigzag']];function llIla(llla,l1la){if(l1l<5.5)return;var sF="progid:DXImageTransform.Microsoft."+lllIl[llla-25]+'('+transOptions+',duration='+l1la+')';return sF;};function I1lIa(lll){if(IIlI&&!l11I){var sF="filter:";if(lll.lIIIl)if(lll.lIIIl==24)sF+="blendTrans(Duration="+lll.IlIa/1000+") ";else if(lll.lIIIl<24)sF+="revealTrans(Transition="+lll.lIIIl+",Duration="+lll.IlIa/1000+") ";else sF+=llIla(lll.lIIIl,lll.IlIa/1000);if(lll.llIlI)sF+="Alpha(opacity="+lll.llIlI+") ";if(lll.l1IlI)sF+="Shadow(color="+lll.l1IlI+",direction=135,strength="+lll.IIllI+") ";sF+=";";return sF;}else return"";};function II1Ia(n,I11I,i){return'apy'+n+'m'+I11I+'i'+i+((I1l1&&l1l<7)?'ITX':'ITD');};function apy_changeItemText(n,I11I,i,text){if(lI1)return null;var item=I1II(II1Ia(n,I11I,i));item.innerHTML=text;};function apy_changeItem(n,I11I,i,Illll,l111l,IlIIl,l1lI1,IIl11){if(lI1)return null;var item=I1II(II1Ia(n,I11I,i));if(Illll)item.innerHTML=Illll;var II=IlIa(item.id);if(l111l)II.lll11=l111l;if(IlIIl){item=I1II('apy'+n+'m'+I11I+'i'+i+'I');item.title=IlIIl;};if(IIl11){II.lll1[0]=IIl11;item=I1II('apy'+n+'m'+I11I+'i'+i+'ICO');item.src=IIl11;};if(l1lI1)II.lll1[1]=l1lI1;};var Il11=false,I1ll=false;function apy_setPressedItem(n,I11I,i,I1lll){var ll1=l11[n];if(!Il11&&ll1.lII!=-1){Il11=true;with(ll1){apy_setPressedItem(n,ll11,lII,I1lll);if(ll11==I11I&&lII==i){ll11=0;lII=-1;return;};};};if(!Il11){ll1.ll11=I11I;ll1.lII=i;}else Il11=false;var II=IlIa('apy'+n+'m'+I11I+'i'+i);if(!I1ll)II.l1I=!II.l1I;I1ll=false;if(!lI1)ll1la(II,(II.l1I?1:0),II.l1l11);if(I1lll&&I11I>0){var lIl=l11[n].I11I[I11I];for(var j=lIl.lI1I1;j>0;j--){rootI=IlIa(lIl.IIll);if(!lI1)ll1la(rootI,(II.l1I?1:0),rootI.l1l11);else if(j==1)with(I1II(rootI.id).document.layers[0]){document.layers[1].visibility=(II.l1I?"show":"hide");document.layers[0].visibility=(II.l1I?"hide":"show");};rootI.l1I=II.l1I;lIl=l11[n].I11I[rootI.l1l11];};};};function lIlIa(event){var x=0,y=0;if(I1l||III){x=event.clientX+(l1l1?llI1.scrollLeft:0);y=event.clientY+(l1l1?llI1.scrollTop:0);}else{x=event.pageX;y=event.pageY;};return[x,y];};function apy_popup(Il11l,IIII1,event,x,y){if(I1l)event.returnValue=false;if(x&&y)var l1lll=[x,y];else var l1lll=lIlIa(event),ll1=l11[Il11l],II11l=ll1.I11I[1];if(II11l){var llI=I1II(II11l.id);if(llI.style.visibility=="visible"){clearTimeout(ll1.IlI1a);Illla(ll1.I11I[0].I11);status='';};ll1.I11I[0].I11=II11l.id;l1111a(II11l.id);llI.style.left=l1lll[0]+lIl1;llI.style.top=l1lll[1]+lIl1;if(IIII1>0)ll1.IlI1a=setTimeout("Illla('"+ll1.I11I[0].I11+"'); status='';",IIII1);};return false;};
