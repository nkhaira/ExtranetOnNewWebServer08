L_H_APP="MSN Maps &amp; Directions";
H_URL_BASE="http://help.msn.com/" + "en_us";
H_CONFIG="msnmappointv2.ini";
function DoHelpKey(key,text){bSearch=true;H_KEY=key;H_TOPIC='';L_H_TEXT=text;DoHelp()}
function DoHelpTopic(topic){bSearch=false;H_KEY='';H_TOPIC=topic+'.htm';DoHelp()}

var eDistanceUnits=0,eMapSize=1,eZoom=2,eRegion=3,eAltRegion=4,eMapStyle=5;
function StoreInCookie(Idx,Value){
var Nm="MapPoint3=",v="??????",s=document.cookie;
var i=s.indexOf(Nm);
if(i>-1){i+=Nm.length;if(i+6<=s.length)v=s.substring(i,i+6);}
v=v.substring(0,Idx)+unescape("%"+((Value-0)+65).toString(16))+v.substring(Idx+1,v.length);
var Exp=new Date();
Exp.setYear(Exp.getYear()+10);
document.cookie=Nm+v+"; expires="+Exp.toGMTString()+"; path=/";
}

function getPartnerId(){var s=document.cookie+";";var i=s.indexOf("mh="),p="";if(i>=0)p=s.substring(i+3,s.indexOf(";",i));if(null==p)p="";return p.substring(0,4);}
function logoImg(s){document.write("<img border=\"0\" width=\"118\" height=\"35\" src=\""+(s==''?'/I/logo.gif':s+"/global/c/lg/"+getPartnerId()+"_118x35.gif")+"\" alt=\"go to MSN.com\" title=\"go to MSN.com\"/>");}

function Ct(n){var c=[5,7,1,9,3,4,11,8,12,10,6,2,0,1,0,2,13,14,15,16];return c[Fc(n).options[Fc(n).selectedIndex].value]};
function oDirLinkClick(){
var t=Fc("StreetText").value,u="directionsfind.aspx?&src=MP";
if(Fc("SearchType")[1].checked){
if(""!=t)u+="&plce2="+t;
u+="&regn2="+Ct("PRegionSelect");
}else{
if(""!=t)u+="&strt2="+t;
t=Fc("CityText").value;if(""!=t)u+="&city2="+t;
t=Fc("StateText").value;if(""!=t)u+="&stnm2="+t;
t=Fc("ZipText").value;if(""!=t)u+="&zipc2="+t;
u+="&cnty2="+Ct("ARegionSelect");
}window.location=u;}

var aImg;
function LoadMapImg(){
var g=new Array('ZIn','ZOut','Zd','Za','print','email','pda','USEN_e','USEN_w','USEN_n','USEN_s','ne','nw','se','sw','MapTFin','MapBFin');
var j=g.length;
aImg=new Array(j);
while(j--)(aImg[j]=new Image()).src='/I/'+g[j]+'.gif';
}
