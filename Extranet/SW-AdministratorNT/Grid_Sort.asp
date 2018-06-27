<body>
<form name=myform>
<input type=radio name=tm value=1><img name="img1" src="/find-sales/download/thumbnail/0108_t.jpg"><input type=text name=foo1 disabled value=first><BR>
<input type=radio name=tm value=2><img name="img2" src="/find-sales/download/thumbnail/0105_t.jpg"><input type=text name=foo2 disabled value=second><BR>
<input type=radio name=tm value=3><img name="img3" src="/find-sales/download/thumbnail/0109_t.jpg"><input type=text name=foo3 disabled value=third><BR>
<input type=button Value"Up" onclick="mup();">
<input type=button Value"Down" onclick="mdw();">
</form>
<script language="JavaScript1.2">
function mup() {
  var old,val;
  var oldval,pwid,uwid,i;
  
  for (i=0;i<3;i++) {
    if (document.myform.tm[i].checked) {
      val = document.myform.tm[i].value;
      document.myform.tm[i].checked = false;
      break;
    }
  }
  
  old = val - 1;
  
  for (i=0;i<3;i++) {
    if (old == i && i != 0) {
      // Do Text Box Update
      pwid = eval('document.myform.foo'+old);
      uwid = eval('document.myform.foo'+val);
      oldval = pwid.value;
      pwid.value = uwid.value;
      uwid.value = oldval;
      // Do Image Update
      pwid = eval('document.img'+old);
      uwid = eval('document.img'+val);
      oldval = pwid.src;
      pwid.src = uwid.src;
      uwid.src = oldval;
      // Update Radio Button
      document.myform.tm[--old].checked = true;
    }
  }
}

function mdw() {
  var old,val;
  var oldval,pwid,uwid,i;
  
  for (i=0;i<3;i++) {
    if (document.myform.tm[i].checked) {
      val = document.myform.tm[i].value;
      document.myform.tm[i].checked = false;
      break;
    }
  }
  
  old = val + 1;
  alert('val is '+val);
  
  for (i=3;i>=1;i--) {
    if (old == i) {
      // Do Text Box Update
      pwid = eval('document.myform.foo'+old);
      uwid = eval('document.myform.foo'+val);
      oldval = pwid.value;
      pwid.value = uwid.value;
      uwid.value = oldval;
      // Do Image Update
      pwid = eval('document.img'+old);
      uwid = eval('document.img'+val);
      oldval = pwid.src;
      alert('pwid is '+pwid.src);
      alert('uwid is '+uwid.src);      
      pwid.src = uwid.src;
      uwid.src = oldval;


      document.myform.tm[++old].checked = true;
    }
  }
}



</script>