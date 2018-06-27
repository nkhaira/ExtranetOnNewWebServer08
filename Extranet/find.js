var aRegions=new Array(
"Austria",
"Belgium",
"Canada",
"Denmark",
"France",
"Germany",
"Italy",
"Luxembourg",
"Netherlands",
"Spain",
"Switzerland",
"United Kingdom",
"United States",
"Europe",
"North America",
"World Atlas",
"Finland (Helsinki)",
"Norway",
"Portugal",
"Sweden");

var aDS=[1,1,2,1,1,1,1,1,1,1,1,1,2,1,2,4,1,1,1,1],
aAddr=[1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,1,1,1,1],
aCfgs=[3,3,2,3,3,3,3,3,3,3,3,4,1,0,0,0,3,3,3,3],
fields = new Array("City","State","Zip"),
labs = new Array(
new Array("","",""),
new Array("City","State","ZIP Code"),
new Array("City","Province","Postal Code"),
new Array("City","","Postal Code"),
new Array("City","","Postcode"));

var gWith="";
function Ctl(n){return document.getElementById(gWith+n);}
function fcCtl(n){return document.getElementById("FndControl_"+n);}

function SetControls(w,a){
gWith=w;
Ctl("ARegionSelect").disabled=Ctl("PlaceRadio").checked=!a;
Ctl("PRegionSelect").disabled=Ctl("AddressRadio").checked=a;
SetAddrFields(w);
}

function SetAddrFields(w){
gWith=w;
var x,c=aCfgs[Ctl(Ctl("AddressRadio").checked?"ARegionSelect":"PRegionSelect").value];
for(i=0;i<3;++i){
x="hidden";
if(labs[c][i]!=""){Ctl(fields[i]+"L").innerHTML=labs[c][i];x="visible"}
Ctl(fields[i]+"Label").style.visibility=x;
}
Ctl("StreetL").innerHTML=(c==0?"Place Name":"Street Address");
}

function SaveToCookie(w){
gWith=w;
if(Ctl("ARegionSelect")){
var a=Ctl("AddressRadio").checked;
StoreInCookie(eRegion, Ctl(a?"ARegionSelect":"PRegionSelect").value);
StoreInCookie(eAltRegion, Ctl(a?"PRegionSelect":"ARegionSelect").value);}
}

function CheckFindControl(){
if(fcCtl("ARegionSelect")){
var a=fcCtl("AddressRadio").checked;
fcCtl("ARegionSelect").disabled=!a;
fcCtl("PRegionSelect").disabled=a;
SetAddrFields("FndControl_");
}}

function onARegionChange(){fcCtl("BkARegion").value=fcCtl("ARegionSelect").value;SetAddrFields("FndControl_");}
function onPRegionChange(){fcCtl("BkPRegion").value=fcCtl("PRegionSelect").value;}
function onPRegionClick(){if(!fcCtl("PlaceRadio").checked)onPlaceRadioClick();}
function onARegionClick(){if(!fcCtl("AddressRadio").checked)onAddressRadioClick();}
function onPlaceRadioClick(){SetControls("FndControl_",false);}
function onAddressRadioClick(){SetControls("FndControl_",true);}

function GetMap(){
if(fcCtl("AmbiguousSelect")){var s=fcCtl("AmbiguousSelect").selectedIndex;if(-1<s){AmbiguousClick(s);return;}}
if(fcCtl("ARegionSelect"))fcCtl("BkARegion").value=fcCtl("ARegionSelect").value;
if(fcCtl("PRegionSelect"))fcCtl("BkPRegion").value=fcCtl("PRegionSelect").value;
SaveToCookie("FndControl_");
if(document.FindForm.fireEvent)document.FindForm.fireEvent("onsubmit");
document.FindForm.submit();
}

function onKeyDown(){if("13"==window.event.keyCode){GetMap();window.event.returnValue=false;}}
function MoreResults(){fcCtl("isRegionChange").value="2";document.FindForm.submit();}
