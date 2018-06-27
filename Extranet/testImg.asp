<%@LANGUAGE="VBSCRIPT"%>

<script language=javascript>
function resize1(img) {
var canvasWidth ,canvasHeight
canvasHeight = 270
canvasWidth =328

      var thumb =  img // document.getElementById(img);
      var imageWidth = thumb.offsetWidth;
      var imageHeight = thumb.offsetHeight;
      if ((imageWidth / imageHeight) >= (canvasWidth / canvasHeight)) {
         thumb.style.width = canvasWidth + "px";
         thumb.style.height = (imageHeight * canvasWidth / imageWidth) + "px";
      } else {
         thumb.style.width = (imageWidth * canvasHeight / imageHeight) + "px";
         thumb.style.height = canvasHeight + "px";
      }      
   }
   

</script>
<body>
<form id="form1" name="form1" method="post" action="">
<table width="400" border="1" align="center">
    <tr>
      <td colspan="2" align="center"><strong>Image Aspect Ratio using Java Script</strong></td>
    </tr>  
    <tr>
      <td><img src="images\Adrepr_w.jpg" alt = "" />
      </td>
      <td><img src="images\Fluke_RPM logo.jpg" alt ="" />
      </td>   
    
    <tr>
      <td><img src="images\Adrepr_w.jpg" alt ="" name ="img1"  onload ="Javascript:resize1(this);"/>
      </td>
      <td><img src="images\Fluke_RPM logo.jpg" alt = "" name ="img2" onload ="Javascript:resize1(this);" />
      </td>
  
    </tr>    
    </table>

</form>
</body>

