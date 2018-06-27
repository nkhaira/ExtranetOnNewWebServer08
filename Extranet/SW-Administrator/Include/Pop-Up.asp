
<!-- Pop-Up Window Begin -->

<SCRIPT LANGUAGE=JAVASCRIPT>

<!--

function Checkitout(){

//      Gets Browser and Version

        var appver = "null";
        var browser = navigator.appName;
        var version = navigator.appVersion;
        if ((browser == "Netscape")) version = navigator.appVersion.substring(0, 3);
        if ((browser == "Microsoft Internet Explorer")) version = navigator.appVersion.substring(22, 25);

//      Gives AppVersion (appver) for Detect Strings

        if ((browser == "Microsoft Internet Explorer") && (version >= 3)) appver = "ie3+";
        if ((browser == "Netscape") && (version >= 3)) appver = "ns3+";
        if ((browser == "Netscape") && (version < 3)) appver = "ns2";


       if ((appver == "ie3+")) {
                return 0;
        }  else {
                return 1;
                }
}

function openit(DaURL, orient) {
  var File_Load;
  File_Load = window.open(DaURL,"File_Load","status=yes,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=no,links=no");
  File_Load.blur();
}

function openit_maxi(DaURL, orient) {
  var Maxi_pop_up;
  Maxi_pop_up = window.open(DaURL,"Maxi","status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=yes,links=no");
  Maxi_pop_up.focus();
}

function openit_mini(DaURL, orient) {
  var Mini_pop_up;
  Mini_pop_up = window.open(DaURL,"Mini","status=no,height=410,width=525,scrollbars=yes,resizable=yes,toolbar=no,links=no");
  Mini_pop_up.focus();
}

//-->

</SCRIPT>

<!-- Pop-Up Window End -->


