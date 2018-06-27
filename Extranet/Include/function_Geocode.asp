<%
response.write GeoDistance(48.027419000000002,-122.062363,48.002966999999998,-122.256364,"m")

' --------------------------------------------------------------------------------------
'                                                                         
'  This routine calculates the distance between two points (given the     
'  latitude/longitude of those points). it is being used to calculate     
'  the distance between two zip codes or postal codes.                     
'                                                                         
'  definitions:                                                           
'    south latitudes are negative, east longitudes are positive           
'                                                                         
'  passed to function:                                                    
'    lat1, lon1 = latitude and longitude of point 1 (in decimal degrees)  
'    lat2, lon2 = latitude and longitude of point 2 (in decimal degrees)  
'    unit = the unit you desire for results                               
'           where: 'm' is statute miles                                   
'                  'k' is kilometers (default)                            
'                  'n' is nautical miles                                  
'                                                                         
' --------------------------------------------------------------------------------------

function GeoNearBy()



end function



function GeoDistance(lat1, lon1, lat2, lon2, unit)

  dim theta, dist
  
  pi = 3.14159265358979323846
  
  theta = lon1 - lon2
  dist = sin(deg2rad(lat1)) * sin(deg2rad(lat2)) + cos(deg2rad(lat1)) * cos(deg2rad(lat2)) * cos(deg2rad(theta))
  response.write "dist = " & dist & "<br>"
  dist = acos(dist)
  dist = rad2deg(dist)
  response.write "dist = " & dist & "<br>"
  GeoDistance = dist * 60 * 1.1515

  select case LCase(unit)
    case "k"
      GeoDistance = GeoDistance * 1.609344
    case "n"
      GeoDistance = GeoDistance * 0.8684
  end select

end function 

' --------------------------------------------------------------------------------------
'  Get the arccos function using arctan function   
' --------------------------------------------------------------------------------------':

function ACOS(rad)

  pi = 3.14159265358979323846

  if abs(rad) <> 1 then
    acos = pi/2 - atn(rad / sqr(1 - rad * rad))
  elseIf rad = -1 then
    acos = pi
  end if

end function

' --------------------------------------------------------------------------------------
'  Convert decimal degrees to radians             
' --------------------------------------------------------------------------------------

function Deg2Rad(deg)

  pi = 3.14159265358979323846
	Deg2Rad = cdbl(deg * pi / 180)

end function

' --------------------------------------------------------------------------------------
'  Convert radians to decimal degrees             
' --------------------------------------------------------------------------------------

function Rad2Deg(rad)
  
  pi = 3.14159265358979323846
	Rad2Deg = cdbl(rad * 180 / pi)
  
end function

' --------------------------------------------------------------------------------------
%>
