<%				
        if len(waybill) > 0 then
					if Instr(lcase(waybill),"b/o") > 0 then
						waybill = "back order"
						shp_method = "&nbsp;"
					elseif (Instr(ucase(shp_method),"UPS") = 1) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','UPS');"">" & waybill & "</a>"
					elseif (Instr(ucase(shp_method),"TNT") <> 0) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','TNT');"">" & waybill & "</a>"
          ' this shows all MFG orders with shipped dates as links to TNT
					elseif (ucase(strSrcSystem) = "MFG" and len(dbRS("TShip_Date")&"") > 0) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','TNT');"">" & waybill & "</a>"
					elseif (Instr(ucase(shp_method),"FDX") = 1) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" &_
  					  Trim(Replace(waybill,"+","")) & "','FDX');"">" & waybill & "</a>"
					elseif (Instr(ucase(shp_method),"PUR") = 1) then
						bShowWayUL = True
						waybill = "<a href=""javascript:Track('" & Trim(waybill) & "','PUR');"">" &_
              waybill & "</a>"
					elseif (Instr(ucase(shp_method),"DIR.SH") = 1) then
						waybill = ""
						bShowDirect = True
					end if
				end if
        %>
<script language="Javascript">
	function Track(num,carrier) {
		
		if (carrier == 'UPS') {
			var vHref = 'http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=5';
  			vHref += '&TypeOfInquiryNumber=T&HTMLVersion=4.0&InquiryNumber1=' + num;
		}
		else if (carrier == 'FDX') {
			var vHref = 'http://www.fedex.com/cgi-bin/tracking?action=track&language=english&';
        vHref += 'cntry_code=us&initial=x&tracknumbers=' + num;
		}
		else if (carrier == 'TNT') {
			var vHref = 'http://www.tntew.com/new_tracker/SaCGI.cgi/tracker.exe?';
        vHref += 'FNC=gotoresults__Adummy_html___conok='+num+'___ttype=R___lang=EN___page=0';
        vHref += '___laf=default';
		}
		else if (carrier == 'PUR') {
			var vHref = 'http://shipnow.purolator.com/shiponline/track/PurolatorTrackE.asp?';
        vHref += 'PINNO=' + num;
		}
		
		var wName = 'Track';
	
		newWind = window.open(vHref,wName);
	
		if (newWind.opener == null) {
		   newWind.opener = window;
		}
		// self.blur();
		newWind.focus();
	}
</script>
        
