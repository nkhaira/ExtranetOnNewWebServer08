var enablepersist = true
var slidernodes = new Object() //Object array to store references to each content slider's DIV containers (<div class="contentdiv">)

function ContentSlider(sliderid, autorun){
  var slider = document.getElementById(sliderid)
  slidernodes[sliderid] = [] //Array to store references to this content slider's DIV containers (<div class="contentdiv">)
  var alldivs = slider.getElementsByTagName("div")
  for (var i = 0; i < alldivs.length; i++){
    if (alldivs[i].className == "contentdiv"){
      slidernodes[sliderid].push(alldivs[i]) //add this DIV reference to array
    }
  }
  ContentSlider.buildpagination(sliderid)
  var loadfirstcontent = true
  if (enablepersist && getCookie(sliderid) != ""){ //if enablepersist is true and cookie contains corresponding value for slider
    var cookieval = getCookie(sliderid).split(":") //process cookie value ([sliderid, int_pagenumber (div content to jump to)]
    if (document.getElementById(cookieval[0]) != null && typeof slidernodes[sliderid][cookieval[1]] != "undefined"){ //check cookie value for validity
      ContentSlider.turnpage(cookieval[0], parseInt(cookieval[1])) //restore content slider's last shown DIV
      loadfirstcontent = false
    }
  }
  if (loadfirstcontent == true) //if enablepersist is false, or cookie value doesn't contain valid value for some reason (ie: user modified the structure of the HTML)
  ContentSlider.turnpage(sliderid, 0) //Display first DIV within slider
  if (typeof autorun == "number" && autorun>0) //if autorun parameter (int_miliseconds) is defined, fire auto run sequence
  window[sliderid+"timer"] = setTimeout(function(){ContentSlider.autoturnpage(sliderid, autorun)}, autorun)
}

ContentSlider.buildpagination = function(sliderid){
  var paginatediv = document.getElementById("paginate-" + sliderid) //reference corresponding pagination DIV for slider
  var pcontent = ""
  for (var i = 0; i < slidernodes[sliderid].length; i++) //For each DIV within slider, generate a pagination link
    pcontent += '<a href="#" onClick=\"ContentSlider.turnpage(\'' + sliderid + '\', ' + i + '); return false\">' + (i + 1) + '</a> '
    pcontent += '<a href="#" style="font-weight: bold;" onClick=\"ContentSlider.turnpage(\''+sliderid+'\', parseInt(this.getAttribute(\'rel\'))); return false\">Next</a>'
    paginatediv.innerHTML=pcontent
    paginatediv.onclick = function(){ //cancel auto run sequence (if defined) when user clicks on pagination DIV
    if (typeof window[sliderid+"timer"] != "undefined")
    clearTimeout(window[sliderid+"timer"])
  }
}

ContentSlider.turnpage = function(sliderid, thepage){
  var paginatelinks = document.getElementById("paginate-"+sliderid).getElementsByTagName("a") //gather pagination links
  for (var i = 0; i < slidernodes[sliderid].length; i++){ //For each DIV within slider
    paginatelinks[i].className = "" //empty corresponding pagination link's class name
    slidernodes[sliderid][i].style.display = "none" //hide DIV
  }
  paginatelinks[thepage].className = "selected" //for selected DIV, set corresponding pagination link's class name
  slidernodes[sliderid][thepage].style.display = "block" //show selected DIV
  //Set "Next" pagination link's (last link within pagination DIV) "rel" attribute to the next DIV number to show
  paginatelinks[paginatelinks.length-1].setAttribute("rel", thenextpage = (thepage<paginatelinks.length-2)? thepage+1 : 0)
  if (enablepersist)
  setCookie(sliderid, sliderid + ":" + thepage)
}

ContentSlider.autoturnpage = function(sliderid, autorunperiod){
  var paginatelinks = document.getElementById("paginate-" + sliderid).getElementsByTagName("a") //Get pagination links
  var nextpagenumber = parseInt(paginatelinks[paginatelinks.length-1].getAttribute("rel")) //Get page number of next DIV to show
  ContentSlider.turnpage(sliderid, nextpagenumber) //Show that DIV
  window[sliderid + "timer"]=setTimeout(function(){ContentSlider.autoturnpage(sliderid, autorunperiod)}, autorunperiod)
}

function getCookie(Name){ 
  var re = new RegExp(Name + "=[^;]+", "i"); //construct RE to search for target name/value pair
  if (document.cookie.match(re)) //if cookie found
  return document.cookie.match(re)[0].split("=")[1] //return its value
  return ""
}

function setCookie(name, value){
  document.cookie = name + "=" + value
}