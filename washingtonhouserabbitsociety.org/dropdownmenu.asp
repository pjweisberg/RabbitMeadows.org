

<style type="text/css">

#dropmenudiv{
position:absolute;

border-bottom-width: 0;
font:normal 10px Verdana;
color:black;
line-height:18px;
z-index:100;

/*filter:alpha(opacity=100);*/
}

#dropmenudiv a{
width: 100%;
display: block;
text-indent: 3px;
border-bottom: 1px solid #000;
border-left: 1px solid #000;
border-right: 1px solid #000;
padding: 1px 0;
background-color: #CC9;
text-decoration: none;
font:normal 10px Verdana;
color: #000;

}

#dropmenudiv a:hover{ /*hover background color*/
background-color: #FD8;

}

.linktext {
font-family:Verdana,Arial,Helvetica; 
color:#000; 
font-size:10px; 
text-decoration: none
}

</style>



<script type="text/javascript">

/***********************************************
* AnyLink Drop Down Menu-  Dynamic Drive (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit http://www.dynamicdrive.com/ for full source code
***********************************************/

//Contents for menu 1


//Contents for Home Menu
var menuHome=new Array()
menuHome[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/index.asp">Home</a>'


//Contents for Adoption Menu
var menuAdopt=new Array()
menuAdopt[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/Adoptedcurrent.asp">Adopt Rabbits</a>'
menuAdopt[1]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/RodentCurrent.asp">Adopt Rodents</a>'
menuAdopt[2]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/RodentCurrent.asp">Adopt G. Pigs</a>'
menuAdopt[3]='<a href="http://www.washingtonhouserabbitsociety.org/washingtonhouserabbitsociety.org/FerretCurrent.asp">Adopt Ferrets</a>'

//Contents for News Menu
var menuNews=new Array()
menuNews[0]='<a href="http://www.washingtonhouserabbitsociety.org/washingtonhouserabbitsociety.org/news.asp">Rabbit News</a>'
menuNews[1]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/comingsoon.asp?name=Rodent News">Rodent News</a>'
menuNews[2]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/comingsoon.asp?name=Guinea Pig News">G. Pig News</a>'
menuNews[3]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/comingsoon.asp?name=Ferret News">Ferret News</a>'

//Contents for Vets Menu
var menuVets=new Array()
menuVets[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/vets.asp?animal=1">Rabbit Vets</a>'
menuVets[1]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/vets.asp?animal=2">Rodent Vets</a>'
menuVets[2]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/vets.asp?animal=4">G. Pig Vets</a>'
menuVets[3]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/vets.asp?animal=3">Ferret Vets</a>'

//Contents for FAQ's Menu
var menuFaq=new Array()
menuFaq[0]='<a href="http://www.washingtonhouserabbitsociety.org/washingtonhouserabbitsociety.org/FAQ.asp">Rabbit FAQs</a>'
menuFaq[1]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/comingsoon.asp?name=Rodent FAQs">Rodent FAQs</a>'
menuFaq[2]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/comingsoon.asp?name=Guinea Pig FAQs">G. Pig FAQs</a>'
menuFaq[3]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/comingsoon.asp?name=Ferret FAQs">Ferret FAQs</a>'

//Contents for Store Menu
var menuStore=new Array()
menuStore[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/store.asp">Shop our Store!</a>'

//Contents for Contact Menu
var menuCon=new Array()
menuCon[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/contactus.asp">Contact Info</a>'

var menuwidth='165px' //default menu width
var menubgcolor='#CC9'  //menu bgcolor
var disappeardelay=250  //menu disappear speed onMouseout (in miliseconds)
var hidemenu_onclick="yes" //hide menu when user clicks within menu?



/////No further editting needed

var ie4=document.all
var ns6=document.getElementById&&!document.all


if (ie4||ns6)
document.write('<div id="dropmenudiv"  style="visibility:hidden;width:'+menuwidth+';" onMouseover="clearhidemenu()" onMouseout="dynamichide(event)"></div>')


function getposOffset(what, offsettype){
var totaloffset=(offsettype=="left")? what.offsetLeft : what.offsetTop;
var parentEl=what.offsetParent;
while (parentEl!=null){
totaloffset=(offsettype=="left")? totaloffset+parentEl.offsetLeft : totaloffset+parentEl.offsetTop;
parentEl=parentEl.offsetParent;
}
return totaloffset;
}


function showhide(obj, e, visible, hidden, menuwidth){
if (ie4||ns6)
dropmenuobj.style.left=dropmenuobj.style.top=-500
if (menuwidth!=""){
dropmenuobj.widthobj=dropmenuobj.style
dropmenuobj.widthobj.width=menuwidth
}
if (e.type=="click" && obj.visibility==hidden || e.type=="mouseover")
obj.visibility=visible
else if (e.type=="click")
obj.visibility=hidden
}

function iecompattest(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function clearbrowseredge(obj, whichedge){
var edgeoffset=0
if (whichedge=="rightedge"){
var windowedge=ie4 && !window.opera? iecompattest().scrollLeft+iecompattest().clientWidth-15 : window.pageXOffset+window.innerWidth-15
dropmenuobj.contentmeasure=dropmenuobj.offsetWidth
if (windowedge-dropmenuobj.x < dropmenuobj.contentmeasure)
edgeoffset=dropmenuobj.contentmeasure-obj.offsetWidth
}
else{
var topedge=ie4 && !window.opera? iecompattest().scrollTop : window.pageYOffset
var windowedge=ie4 && !window.opera? iecompattest().scrollTop+iecompattest().clientHeight-15 : window.pageYOffset+window.innerHeight-18
dropmenuobj.contentmeasure=dropmenuobj.offsetHeight
if (windowedge-dropmenuobj.y < dropmenuobj.contentmeasure){ //move up?
edgeoffset=dropmenuobj.contentmeasure+obj.offsetHeight
if ((dropmenuobj.y-topedge)<dropmenuobj.contentmeasure) //up no good either?
edgeoffset=dropmenuobj.y+obj.offsetHeight-topedge
}
}
return edgeoffset
}

function populatemenu(what){
if (ie4||ns6)
dropmenuobj.innerHTML=what.join("")
}


function dropdownmenu(obj, e, menucontents, menuwidth){
if (window.event) event.cancelBubble=true
else if (e.stopPropagation) e.stopPropagation()
clearhidemenu()
dropmenuobj=document.getElementById? document.getElementById("dropmenudiv") : dropmenudiv
populatemenu(menucontents)

if (ie4||ns6){
showhide(dropmenuobj.style, e, "visible", "hidden", menuwidth)
dropmenuobj.x=getposOffset(obj, "left")
dropmenuobj.y=getposOffset(obj, "top")
dropmenuobj.style.left=dropmenuobj.x-clearbrowseredge(obj, "rightedge")+"px"
dropmenuobj.style.top=dropmenuobj.y-clearbrowseredge(obj, "bottomedge")+obj.offsetHeight+"px"
}

return clickreturnvalue()
}

function dropdownmenub(obj, e, menucontents, menuwidth){
if (window.event) event.cancelBubble=true
else if (e.stopPropagation) e.stopPropagation()
clearhidemenu()
dropmenuobj=document.getElementById? document.getElementById("dropmenudiv") : dropmenudiv
populatemenu(menucontents)

if (ie4||ns6){
showhide(dropmenuobj.style, e, "visible", "hidden", menuwidth)
dropmenuobj.x=getposOffset(obj, "left")
dropmenuobj.y=getposOffset(obj, "top")
dropmenuobj.style.left=dropmenuobj.x-clearbrowseredge(obj, "rightedge")+"px"
dropmenuobj.style.top=dropmenuobj.y-clearbrowseredge(obj, "bottomedge")+obj.offsetHeight+"px"
}

return clickreturnvalue()
}
function clickreturnvalue(){
if (ie4||ns6) return false
else return true
}

function contains_ns6(a, b) {
while (b.parentNode)
if ((b = b.parentNode) == a)
return true;
return false;
}

function dynamichide(e){
if (ie4&&!dropmenuobj.contains(e.toElement))
delayhidemenu()
else if (ns6&&e.currentTarget!= e.relatedTarget&& !contains_ns6(e.currentTarget, e.relatedTarget))
delayhidemenu()
}

function hidemenu(e){
if (typeof dropmenuobj!="undefined"){
if (ie4||ns6)
dropmenuobj.style.visibility="hidden"
}
}

function delayhidemenu(){
if (ie4||ns6)
delayhide=setTimeout("hidemenu()",disappeardelay)
}

function clearhidemenu(){
if (typeof delayhide!="undefined")
clearTimeout(delayhide)
}

if (hidemenu_onclick=="yes")
document.onclick=hidemenu

</script>

