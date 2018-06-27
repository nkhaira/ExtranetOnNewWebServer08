<html>
<head>
<title> Test </title>
</head>

<script>

//	rows number:
var rowNum = 20;

//	columns number:
var colNum = 9;


//	Array that holds the columns width
var arrTDwidth = new Array (77, 31, 71, 113, 52, 53, 41, 64, 55, 90, 
70, 145, 23, 50, 30, 120, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 
30, 120, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 
70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 70, 180, 12, 90, 
70, 145, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 
70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 70, 180, 12, 90, 
70, 145, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 
70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 70, 180, 12, 90, 
70, 145, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 
70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 70, 180, 12, 90, 
70, 145, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 
70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 70, 180, 12, 90, 
70, 145, 200, 80, 70, 180, 12, 90, 70, 145, 23, 50, 30, 120, 200, 80, 
70, 180, 12, 90, 70);




//	function that create the table
function drawTab (rows, cols, color, id)
{
	sHTML = "<TABLE id='"+id+"' bgcolor='"+color+"' border=1 style='tableLayout:fixed'>";
	count=0;

	sHTML += "\n<THEAD id='downTblHead'>\n";
	
	i=0;
	
	for (var j=0; j < cols; j++)
	{
		width = arrTDwidth;
		sHTML += "<TH id='a0"+j+"' nowrap style='word-wrap: break-word;' style='width:"+arrTDwidth[count]+"pt;'>this is cell: Row"+i+"Col"+j+"</TH>";
		//alert ("this is cell: Row"+i+"Col"+j+"\n"+arrTDwidth[count]);
		count++;
	}
	count = 0;
	
	sHTML += "\n</THEAD>\n";
	sHTML += "\n<TBODY>\n";
	
	for (var i=1; i<rows; i++)
	{
		sHTML += "\n<TR>\n";
		
		for (var j=0; j<cols; j++)
		{
			width = arrTDwidth;
			sHTML += "<TD id='a"+i+j+"' style='word-wrap: break-word;' nowrap style='word-wrap: break-word;' style='width:"+arrTDwidth[count]+"pt;'>this is cell: Row"+i+"Col"+j+"</TD>";
			//alert ("this is cell: Row"+i+"Col"+j+"\n"+arrTDwidth[count]);
			count++;
		}
		count = 0
		sHTML += "\n</TR>\n";
	}

	sHTML += "\n</TBODY>\n";
	sHTML += "\n</TABLE>";

	return sHTML;
	
}


// function that clones the real header and changes the original IDs (for not duplicate any IDs).
function cloneHeader ()
{
	// clone header
	var x = document.all['downTblHead'].cloneNode(true);
	
	//	changing orig. IDs
	for (var j = 0; j < colNum; j++)
	{
		document.all['a0'+j].id = 'head'+j;
	}

	document.all['downTblHead'].id = "downTblHead4hidden";
	
	//	adding the cloned header to the upper table.
	document.all['upTbl'].appendChild (x);


	document.all['upTbl'].width = document.all
['downTblHead4hidden'].offsetWidth;
	
	synchronizeHeader ()
}



//	function that synchronize
function synchronizeHeader ()
{
	var ok = true;
	var max = 0;
	
	//	the table may not synchronize at first time; if not, I repeat the procedure (maximum 20 times)
	while (ok && max < 20)
	{
		max++;
		ok = false;

		for (var j = 0; j < colNum; j++)
		{
			if (document.all['a0'+j].offsetWidth != document.all['head'+j].offsetWidth)
			{
				ok = true;
				//alert ("before: \n" + document.all['a0'+j].offsetWidth + "\n" + document.all['head'+j].offsetWidth);
				document.all['a0'+j].style.width = document.all['head'+j].offsetWidth;
				//alert ("after: \n" + document.all['a0'+j].offsetWidth + "\n" + document.all['head'+j].offsetWidth);
			}
		}
	}

}




//	reading the scrollBars width (depending on the OS settings).
function getScrollBraWidth ()
{
	try
	{
		var elem = document.createElement("DIV");
		elem.id = "asdf";
		elem.style.width = 100;
		elem.style.height = 100;
		elem.style.overflow = "scroll";
		elem.style.position = "absolute";
		elem.style.visibility = "hidden";
		elem.style.top = "0";
		elem.style.left = "0";
		
		document.body.appendChild (elem);

		scrollWidth = document.all['asdf'].offsetWidth - document.all['asdf'].clientWidth;

		document.body.removeChild (elem);
		delete elem;
	}
	catch (ex)
	{
		return false;
	}

	return scrollWidth;
}


//	function that find the xPos of an HTML object;
function findPosX (obj)
{
	var curleft = 0;

	// go into a loop that continues as long as the object has an offsetParent
	while (obj.offsetParent)
	{
		// add the offsetLeft of the element relative to the offsetParent to curleft and set the object to this offsetParent
		curleft += obj.offsetLeft
		obj = obj.offsetParent;
	}
	return curleft;
}

</script>

<body bgcolor="#000000", text="#ffffff" onLoad="cloneHeader ();">

<div id="headerContainer" style="position: absolute; z-index:2;">
<TABLE  id='upTbl' bgcolor='gray' border=1 style="tableLayout:fixed">
</TABLE>
</div>

<div id="dataContainer" style="height:600;overflow:scroll; z-index:1; 
position: absolute;" >
<script language="JavaScript">
<!--
	//	writting the table
	document.write (drawTab (rowNum, colNum, "", "data"));

	//	adding scrollBar width to the dataContainer div
	document.all['dataContainer'].style.width = document.all
['data'].offsetWidth + getScrollBraWidth ();

	//	adding scrollBar width to the dataContainer div
	document.all['headerContainer'].style.width = document.all
['data'].offsetWidth;
//-->
</script>
</div>

</body>
</html>
