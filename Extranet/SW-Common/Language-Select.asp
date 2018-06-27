<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN">

<HTML>
<HEAD>
<TITLE>Custom Language Select Box</title>
<link rel="STYLESHEET" type="text/css" href="/include/webfx.css">
<link rel="STYLESHEET" type="text/css" href="/include/classic.css">

<script type="text/javascript" src="/include/SW-Fade.js" ></script>
<script type="text/javascript" src="/include/SW-Select.js"></script>
<script type="text/javascript" src="/include/SW-WriteSelect.js"></script>

</HEAD>

<BODY>
<BR><BR>
The following is a visual example of displaying country flag images within a drop-down selection box.  This works for all browser version 4.7+, others will just display the text only in the drop down box.  This rendition uses a fade-in/face-out feature which can be disabled.  The text font also needs to be adjusted, but I think you get the idea.
<BR><BR>
What do you think?
<BR><BR>
Graphical Country Site Selection Drop-Down
<script>

Language_Array = new Array();
Language_Array[0] = new Option('<nobr title="English"><img src="/images/flags/eng.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;">United States</span></nobr>', "http://support.fluke.com/");
Language_Array[1] = new Option('<nobr title="France"><img src="/images/flags/fre.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">French</span></nobr>', "http://support.fluke.com/");
Language_Array[2] = new Option('<nobr title="Denmark"><img src="/images/flags/dan.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Danish</span></nobr>', "http://support.fluke.com/");
Language_Array[3] = new Option('<nobr title="The Netherlands"><img src="/images/flags/dut.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;">Dutch</span></nobr>', "http://support.fluke.com/");
Language_Array[4] = new Option('<nobr title="Italy"><img src="/images/flags/ita.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Italian</span></nobr>', "http://support.fluke.com/");
Language_Array[5] = new Option('<nobr title="Sweden"><img src="/images/flags/swe.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Swedish</span></nobr>', "http://support.fluke.com/");
Language_Array[6] = new Option('<nobr title="China"><img src="/images/flags/chi.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Chinese</span></nobr>', "http://support.fluke.com/");
Language_Array[7] = new Option('<nobr title="Spanin"><img src="/images/flags/spa.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Spanish</span></nobr>', "http://support.fluke.com/");
Language_Array[8] = new Option('<nobr title="Portugual"><img src="/images/flags/por.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Portuguese</span></nobr>', "http://support.fluke.com/");
Language_Array[9] = new Option('<nobr title="Norway"><img src="/images/flags/nor.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Norwegeon</span></nobr>', "http://support.fluke.com/");
Language_Array[10] = new Option('<nobr title="Finland"><img src="/images/flags/fin.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Finnish</span></nobr>', "http://support.fluke.com/");
Language_Array[11] = new Option('<nobr title="Japan"><img src="/images/flags/jpn.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Japanese</span></nobr>', "http://support.fluke.com/");
Language_Array[12] = new Option('<nobr title="Thailand"><img src="/images/flags/tha.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Thai</span></nobr>', "http://support.fluke.com/");
Language_Array[13] = new Option('<nobr title="Korea"><img src="/images/flags/kor.gif" Height=12 WIDTH=20 border=1 align="absmiddle"> <span style="font-family: arial;", "selected">Korean</span></nobr>', "http://support.fluke.com/");

writeSelectBox(Language_Array, "select2", 1, "window.open(this.options[this.selectedIndex].value, '_top')", "margin-left: 10; width: 140;");

</script>

<!-- Footer Starts here -->


</BODY>
</HTML>
