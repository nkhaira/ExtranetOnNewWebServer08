<html>
<head>

<form name="hotmail" method="post"  action="http://lw10fd.law10.hotmail.msn.com/cgi-bin/HoTMaiL" style="margin-top:0px;margin-bottom:0px">

<table border=0 cellpadding=0 cellspacing=2 >
<tr>
<td nowrap  ><div class="bttn"><a href="javascript:document.hotmail._HMaction.value='delete';document.hotmail.submit();"><font class="snu" color="#000066">&nbsp;Delete&nbsp;</font></a></div></td>
<td nowrap  ><div class="bttn"><a href="javascript:document.hotmail._HMaction.value='blocksender';document.hotmail.submit();"><font class="snu" color="#000066">&nbsp;Block&nbsp;Sender(s)&nbsp;</font></a></div></td>
<td width="100%"></td>	
<td nowrap  ><div class="bttn"><a href="javascript:document.hotmail._HMaction.value='MoveTo';document.hotmail.submit();"><font class="snu" color="#000066">&nbsp;Move&nbsp;to&nbsp;</font></a></div></td>
<td align="right" nowrap>
	<select name="tobox" onChange="SynchDrop();">
	<option value="">-Move to Selected Folder-
	<option value="F000000001">Inbox
	<option value="F000000002">Sent Messages
	<option value="F000000003">Drafts
	<option value="F000000004">Trash Can
	<option value="F965370904">SFBC
	</select>
</td>
</tr>
</table>
<table border=0 cellpadding=0 cellspacing=1 width="100%" id="MsgTable">
<input type="hidden" name="curmbox" value="F000000001">

<input type="hidden" name="js" value="">
<input type="hidden" name="_HMaction" value="">
<input type="hidden" name="foo" value="inbox">
<input type="hidden" name="page" value="">	
<tr align="center" bgcolor="#333366">
<td width=20 height=23>
<input name="allbox" type="checkbox" value="Check All" onClick="CheckAll();">
</td>
<td width=28><nobr><img src='http://64.4.14.24/newmail.gif' width=11 height=11 hspace=5 border=0></nobr></td>
<td><nobr><a href="/cgi-bin/HoTMaiL?curmbox=F000000001&a=428947954428888282&sort=From" title="Sort by From Name"><font class="swb" color="ffffff">From</font></a></nobr></td>
<td><nobr><a href="/cgi-bin/HoTMaiL?curmbox=F000000001&a=428947954428888282&sort=Subject" title="Sort by Subject"><font class="swb" color="ffffff">Subject</font></a></nobr></td>
<td bgcolor="#cccc99"> <nobr><a href="/cgi-bin/HoTMaiL?curmbox=F000000001&a=428947954428888282&sort=Date" title="Sort by Date"><img src="http://64.4.14.24/desc.gif" height=7 hspace=3 width=7 border=0 alt="sorted in descending order"><font class="sbbd">Date</font></a></nobr></td>
<td><nobr><a href="/cgi-bin/HoTMaiL?curmbox=F000000001&a=428947954428888282&sort=Size" title="Sort by Size"><font class="swb" color="ffffff">Size</font></a></nobr></td>
</tr>
<tr bgcolor="#eeeecc">
	<td><input type="checkbox" name="MSG974815105.23" onClick="CheckCheckAll();"></td>
	<td>&nbsp;</td>
	<td name="">&nbsp;<a href="/cgi-bin/getmsg?curmbox=F000000001&a=428947954428888282&msg=MSG974815105.23&src=k.law10.hotmail.com:/home/d1/surveys/hmletter.001114:6100&mfs=302">Hotmail Member Servi...</a>&nbsp;</td>
	<td>&nbsp;New Hotmail Services and WebCourier Partners!&nbsp;</td>
	<td>&nbsp;Nov&nbsp;21&nbsp;2000&nbsp;</td>
	<td align="right">&nbsp;1k&nbsp;</td></tr>
<tr bgcolor="#eeeecc">
	<td><input type="checkbox" name="MSG974751368.20" onClick="CheckCheckAll();"></td>
	<td>&nbsp;</td>
	<td name="Mary.Kelly@bigplanet.com">&nbsp;<a href="/cgi-bin/getmsg?curmbox=F000000001&a=428947954428888282&msg=MSG974751368.20&start=308345&len=1319&msgread=1&mfs=302">Mary</a>&nbsp;</td>
	<td>&nbsp;Thanksgiving Thankfulness - from Mary&nbsp;</td>
	<td>&nbsp;Nov&nbsp;20&nbsp;2000&nbsp;</td>
	<td align="right">&nbsp;1k&nbsp;</td></tr>
<tr bgcolor="#eeeecc">
	<td><input type="checkbox" name="MSG974685584.10" onClick="CheckCheckAll();"></td>
	<td>&nbsp;</td>
	<td name="lasson4@juno.com">&nbsp;<a href="/cgi-bin/getmsg?curmbox=F000000001&a=428947954428888282&msg=MSG974685584.10&start=114743&len=193602&msgread=1&mfs=302">lasson4@juno.com</a>&nbsp;</td>
	<td>&nbsp;Fw: : Florida Ballot&nbsp;</td>
	<td>&nbsp;Nov&nbsp;19&nbsp;2000&nbsp;</td>
	<td align="right">&nbsp;189k&nbsp;</td></tr>

</table>
<font class="f" size=2 color="#990000">
</font>
<script language="JavaScript">
<!--
function CheckAll()
{
	for (var i=0;i<document.hotmail.elements.length;i++)
	{
		var e = document.hotmail.elements[i];
		if ((e.name != 'allbox') && (e.type=='checkbox'))
		e.checked = document.hotmail.allbox.checked;
	}
}
function CheckCheckAll()
{
	var TotalBoxes = 0;
	var TotalOn = 0;
	for (var i=0;i<document.hotmail.elements.length;i++)
	{
		var e = document.hotmail.elements[i];
		if ((e.name != 'allbox') && (e.type=='checkbox'))
		{
			TotalBoxes++;
		if (e.checked)
		{
			TotalOn++;
		}
		}
	}
	if (TotalBoxes==TotalOn)
	{document.hotmail.allbox.checked=true;}
	else
	{document.hotmail.allbox.checked=false;}
}
function DoEmpty()
	{
	if (confirm("Are you sure you want to permanently delete all messages in this folder?"))
		window.location = "http://lw10fd.law10.hotmail.msn.com/cgi-bin/HoTMaiL?curmbox=F000000001&a=428947954428888282&DoEmpty=1";
	}
	
function SynchDrop()
	{
	if (document.hotmail.nullbox)
		document.hotmail.nullbox.selectedIndex=document.hotmail.tobox.selectedIndex;
	}
//-->
</script>
</form>
<!-- File: bottomstuff.asp --> 
<p></center>
<!-- FILE: navmenuend.asp -->
</td>
<td width=10 bgcolor="#ffffff">&nbsp;</td>
</tr></table>



<table cellpadding=0 cellspacing=0 border=0 width="100%">
<tr align="center">
		<td  bgcolor="#cccc99" onmouseover='mOvr(this,"#eeeecc");' onmouseout='mOut(this,"#cccc99");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="/cgi-bin/HoTMaiL?curmbox=F000000001&a=428947954428888282"><font class="sbnub">Inbox </font></a></td>
		<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="/cgi-bin/compose?curmbox=F000000001&a=428947954428888282"><font class="swnub">Compose </font></a></td>
		<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="/cgi-bin/addresses?curmbox=F000000001&a=428947954428888282"><font class="swnub">Address&nbsp;Book </font></a></td>
		<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="/cgi-bin/folders?curmbox=F000000001&a=428947954428888282"><font class="swnub">Folders </font></a></td>
		<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="/cgi-bin/options?curmbox=F000000001&a=428947954428888282"><font class="swnub">Options </font></a></td>
		<td  bgcolor="#336699" style="border-right:solid white 1px;" width="100%">&nbsp;</td>
		<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="javascript:DoMsngrLink();"><font class="swnub">Messenger </font></a></td>
	
			<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="border-right:solid white 1px;padding-right:10px;padding-left:10px;"><a href="http://calendar.msn.com/?locale=1033"><font class="swnub">Calendar </font></a></td>
	<td  bgcolor="#336699" onmouseover='mOvr(this,"#333366");' onmouseout='mOut(this,"#336699");' onclick="mClk(this);" style="padding-right:10px;padding-left:10px;"><a href="javascript:DoHelp()"><font class="swnub">Help </font></a></td>
</tr>
</table>

</body></html>
<!-- H: F213.law10.internal.hotmail.com -->
<!-- V: WIN2K 09.02.00.0069 i -->
<!-- D: Nov 17 2000 15:09:21-->
