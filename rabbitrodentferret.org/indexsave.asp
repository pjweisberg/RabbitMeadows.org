<% @LANGUAGE=VBScript %>
<% Option Explicit %>


<%
Function sendrequest
Dim strquestion

strquestion=Request("applicantname") & "<br>" & Request("email") & "<br>"

Dim MyMail
Set MyMail = Server.CreateObject("Persits.MailSender") 
MyMail.Host = "sendmail.brinkster.com" 
MyMail.body = "<html><body>"  & strquestion & "</body></html>" 
MyMail.IsHTML = True 
MyMail.From = "webmaster@rabbitrodentferret.org" 
MyMail.Username = "webmaster@rabbitrodentferret.org" 
MyMail.Password = "climber" 
MyMail.AddAddress "sandi@rabbitrodentferret.org"
MyMail.Subject = "Newsletter Sign Up" 
MyMail.AddCC "houserabbit@clearwire.net" 

If MyMail.Send Then
Response.write ""
Else
Response.write ""
End if


set MyMail = Nothing 
End Function


Dim News
News=0
If Not IsNull(Request("email")) And Request("email")<>"" Then
News=1
End If

If News=1 Then
sendrequest
End If
%>


<HTML>
<HEAD>
<TITLE>Rabbit Meadows - Quality Supplies for Companion Animals</TITLE><meta name="description" content="Best Little Rabbit & Rodent House Pet Supplies and Adoption">

<meta name="keywords" content="rabbit, rodent, pet, companion animal, shelter, boarding, kennel, pet supplies, pet food,
   squirrel, possum, guinea pig, rat, rodent, mice, chinchilla, prairie dog, adoption, washington, seattle, puget, house rabbit, rabbit, rabbits, pet, bunny, bunnies, 
care, breed, breeding, breeds, Humane Society, education, adoption, adopt, non-profit,
	behavior, faq, spay, neuter, animals, lapin, lapine, sanctuary, rabbit sanctuary, woodland park, DEA "> 

<meta name="copyright" content="Copyright 2000 Best Little Rabbit, Rodent and Ferret House. All rights reserved. 
        Contact author for reprint policies.">

<LINK REL=StyleSheet HREF="style.css" TYPE="text/css" MEDIA=screen>
<!--#include file="sandiedit.js"-->



<!--This part is in here so I can refer to the named table cell id's by name without using single quotes-->

<script type="text/javascript">

var td1="td1";
var td2="td2";
var td3="td3";
var td4="td4";
var td5="td5";
var td6="td6";
var td7="td7";
var td8="td8";
var td9="td9";
var td10="td10";


</script>

<SCRIPT language="JavaScript">
<!--hide
function legendpewindow(page)
{
window.open(page,'rescuewin','width=620,height=650,scrollbars=yes,resizable=yes');
}
//-->
</SCRIPT>


<SCRIPT LANGUAGE="JavaScript">
function validate() {

 if (document.newssignup.email.value.length < 5) {
 alert("Please enter your e-mail address so we can send you the newsletter!");
 return false;
}
window.open('newsletterthankyou.asp','thankyouwin','width=320,height=220,scrollbars=yes,resizable=yes');
return true;
}
</script>


<SCRIPT language="JavaScript">
<!--hide
function thankyouwindow(page)
{
window.open(page,'thankyouwin','width=600,height=600,scrollbars=yes,resizable=yes');
}
//-->
</SCRIPT>

<!-- function to change content based on status of above global variables (controls mouseout and content of central info box onclick) -->

<script type="text/javascript">
function changeContentConditional(id,shtml,link) {
   if (document.getElementById || document.all) {
      var el = document.getElementById? document.getElementById(id): document.all[id];
      if (link=="notclicked" && el && typeof el.innerHTML != "undefined") el.innerHTML = shtml;
	  
   }
}
</script>

<!-- function to change content regarless of status of above global variables (for mouseover event) -->

<script type="text/javascript">
function changeContentPure(id,shtml) {
   if (document.getElementById || document.all) {
      var el = document.getElementById? document.getElementById(id): document.all[id];
      if (el && typeof el.innerHTML != "undefined") el.innerHTML = shtml;
	  
   }
}
</script>

<!-- function to change content based on status of above global variables (controls content of link box onclick) -->

<script type="text/javascript">
function changeOnClick(id,msgA,msgB,link) {
   if (document.getElementById || document.all) {
      var el = document.getElementById? document.getElementById(id): document.all[id];
      if (link=="notclicked" && el && typeof el.innerHTML != "undefined") {
	  el.innerHTML = msgA;
	  }
	  else {
	  el.innerHTML = msgB;
	  }
	  
   }
}




var link1rab='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + rabadoption + '</td></tr></table>';
var link2rab='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + rabvetreferral +'</td></tr></table>';
var link3rab='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + rablinks +'</td></tr></table>';
var link4rab='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + rabvolunteer +'</td></tr></table>';
var link5rab='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + rabdonate +'</td></tr></table>';

var link1fer='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + feradoption +'</td></tr></table>';
var link2fer='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + fervetreferral +'</td></tr></table>';
var link3fer='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + ferlinks +'</td></tr></table>';
var link4fer='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + fervolunteer +'</td></tr></table>';
var link5fer='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + ferdonate +'</td></tr></table>';

var link1mea='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + meaabout + '</td></tr></table>';
var link2mea='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + meavolunteer + '</td></tr></table>';
var link3mea='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + meadonate + '</td></tr></table>';
var link4mea='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + meawoodland + '</td></tr></table>';


var link1mem='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + meminfo + '</td></tr></table>';
var link2mem='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + memfamily + '</td></tr></table>';
var link3mem='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + memdonate + '</td></tr></table>';

var link1rod='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + rodadoption + '</td></tr></table>';
var link2rod='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + rodvetreferral + '</td></tr></table>';
var link3rod='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + rodlinks + '</td></tr></table>';
var link4rod='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + rodvolunteer + '</td></tr></table>';
var link5rod='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + roddonate + '</td></tr></table>';

var link1sto='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + storabbits +'</td></tr></table>';
var link2sto='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + storodents +'</td></tr></table>';
var link3sto='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + stoferrets +'</td></tr></table>';

var link1gui='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + guiadoption + '</td></tr></table>';
var link2gui='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + guivetreferral + '</td></tr></table>';
var link3gui='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + guilinks + '</td></tr></table>';
var link4gui='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + guivolunteer + '</td></tr></table>';
var link5gui='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + guidonate + '</td></tr></table>';

var link1don='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + dondonate + '</td></tr></table>';


var msg2rab='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + textmsg2rab + '</td></tr></table>';

var msg2fer='<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + textmsg2fer + '</td></tr></table>';

var msg2mea = '<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + textmsg2mea + '</td></tr></table>';

var msg2sto = '<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + textmsg2sto + '</td></tr></table>';

var msg2rod = '<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Blue">' + textmsg2rod + '</td></tr></table>';

var msg2mem = '<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Orange">' + textmsg2mem + '</td></tr></table>';

var msg2gui = '<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Purple">' + textmsg2gui + '</td></tr></table>';

var msg2don= '<table border="0" cellspacing="0" cellpadding="7" width="100%" height="100%"><tr><td class="Green">' + textmsg2don + '</td></tr></table>';


</script>



<SCRIPT TYPE="text/javascript" SRC="slideshow.js">
</SCRIPT> 

<SCRIPT TYPE="text/javascript">
<!--
SLIDES = new slideshow("SLIDES");



s = new slide();
s.src = "pics/Romeo.jpg";
s.text = "<font face=arial color=white>Romeo was recently adopted and is living the good life!</font>";
SLIDES.add_slide(s);

s = new slide();
s.src = "pics/Snowbunny.jpg";
s.text = "<font face=arial color=white>Our low-maintenance model!  </font>";
SLIDES.add_slide(s);

s = new slide();
s.src = "pics/babies.jpg";
s.text = "<font face=arial color=white>These babies are still looking for a  home.</font>";
SLIDES.add_slide(s);


s = new slide();
s.src = "pics/blackjack.jpg";
s.text = "<font face=arial color=white>All Blackjack wants for Christmas is a forever home. </font>";
SLIDES.add_slide(s);

s = new slide();
s.src = "pics/Kramer1.jpg";
s.text = "<font face=arial color=white>Kramer thinks everyone should sleep with a rabbit!</font>";

SLIDES.add_slide(s);

s = new slide();
s.src = "pics/NessieandNeeps.jpg";
s.text = "<font face=arial color=white>Nessie and Neeps were adopted in January and are doing great in their new home!</font>";

SLIDES.add_slide(s);


s = new slide();
s.src = "pics/pic8.jpg";
s.text = "<font face=arial color=white>I finally found my forever home!</font>";

SLIDES.add_slide(s);

s = new slide();
s.src = "pics/Daisy.jpg";
s.text = "<font face=arial color=white>Daisy got adopted by a family!</font>";
SLIDES.add_slide(s);


s = new slide();
s.src = "pics/SnowBuns.jpg";
s.text = "<font face=arial color=white>Enjoying a rare snowfall at the santuary.</font>";
SLIDES.add_slide(s);

s = new slide();
s.src = "pics/FibiandFooFoo.jpg";
s.text = "<font face=arial color=white>Foo and Fibi relaxing at home.</font>";
SLIDES.add_slide(s);


s = new slide();
s.src = "pics/pic9.jpg";
s.text = "<font face=arial color=white>I GOT ADOPTED!</font>";
SLIDES.add_slide(s);

s = new slide();
s.src = "pics/Bunz.jpg";
s.text = "<font face=arial color=white>Bunz was a stray taken in by a kind person who now dotes on him.</font>";
SLIDES.add_slide(s);

s = new slide();
s.src = "pics/pic7.jpg";
s.text = "<font face=arial color=white>Scamp was adopted but this picture is too cute to take down!</font>";

SLIDES.add_slide(s);

//-->
</SCRIPT>


<!--#include file="dropdownmenu.asp"-->

<script type="text/javascript">

/***********************************************
* AnyLink Drop Down Menu-  Dynamic Drive (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit http://www.dynamicdrive.com/ for full source code
***********************************************/

//Contents for menu 1

var menurab=new Array()
menurab[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/Adoptedcurrent.asp" onmouseover="changeContentPure(td1,link1rab)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'

menurab[1]='<a href="vets.asp?animal=1" onmouseover="changeContentPure(td1,link2rab)"  onmouseout="changeContentPure(td1,msg100)" >Vet Referral</a>'
menurab[2]='<a href="http://hrabbit.brinkster.net/washingtonhouserabbitsociety.org/sites.asp" onmouseover="changeContentPure(td1,link3rab)"  onmouseout="changeContentPure(td1,msg100)">Links</a>'
menurab[3]='<a href="volunteerrab.asp" onmouseover="changeContentPure(td1,link4rab)"  onmouseout="changeContentPure(td1,msg100)">Volunteer</a>'
menurab[4]='<a href="donate.asp?name=rabbits" onmouseover="changeContentPure(td1,link5rab)"  onmouseout="changeContentPure(td1,msg100)">Donate</a>'

var menufer=new Array()
menufer[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/FerretCurrent.asp" onmouseover="changeContentPure(td1,link1fer)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'
menufer[1]='<a href="vets.asp?animal=3" onmouseover="changeContentPure(td1,link2fer)"  onmouseout="changeContentPure(td1,msg100)">Vet Referral</a>'
menufer[2]='<a href="comingsoon.asp?name=Ferret Links" onmouseover="changeContentPure(td1,link3fer)"  onmouseout="changeContentPure(td1,msg100)">Links</a>'
menufer[3]='<a href="volunteer.asp?name=ferrets" onmouseover="changeContentPure(td1,link4fer)"  onmouseout="changeContentPure(td1,msg100)">Volunteer</a>'
menufer[4]='<a href="donate.asp?name=ferrets" onmouseover="changeContentPure(td1,link5fer)"  onmouseout="changeContentPure(td1,msg100)">Donate</a>'

var menumea=new Array()
menumea[0]='<a href="RabbitMeadowsSanctuary.asp" onmouseover="changeContentPure(td1,link1mea)" onmouseout="changeContentPure(td1,msg100)" >Our Sanctuary</a>'
menumea[1]='<a href="volunteer.asp?name=Sanctuary Rabbits" onmouseover="changeContentPure(td1,link2mea)" onmouseout="changeContentPure(td1,msg100)" >Volunteer</a>'
menumea[2]='<a href="donate.asp?name=sanctuary rabbits" onmouseover="changeContentPure(td1,link3mea)" onmouseout="changeContentPure(td1,msg100)" >Donate</a>'
menumea[3]='<a href="http://www.woodlandparkrabbits.org" onmouseover="changeContentPure(td1,link4mea)" onmouseout="changeContentPure(td1,msg100)" >Woodland Park Project</a>'


var menusto=new Array()
menusto[0]='<a href="store.asp" onmouseover="changeContentPure(td1,link1sto)"  onmouseout="changeContentPure(td1,msg100)" >Rabbits</a>'
menusto[1]='<a href="store.asp" onmouseover="changeContentPure(td1,link2sto)"  onmouseout="changeContentPure(td1,msg100)" >Rodents & Guinea Pigs</a>'
menusto[2]='<a href="store.asp" onmouseover="changeContentPure(td1,link3sto)"  onmouseout="changeContentPure(td1,msg100)" >Ferrets</a>'


var menurod=new Array()
menurod[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/RodentCurrent.asp" onmouseover="changeContentPure(td1,link1rod)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'
menurod[1]='<a href="vets.asp?animal=2" onmouseover="changeContentPure(td1,link2rod)"  onmouseout="changeContentPure(td1,msg100)" >Vet Referral</a>'
menurod[2]='<a href="comingsoon.asp?name=Rodent Links" onmouseover="changeContentPure(td1,link3rod)"  onmouseout="changeContentPure(td1,msg100)" >Links</a>'
menurod[3]='<a href="volunteer.asp?name=Rodents" onmouseover="changeContentPure(td1,link4rod)"  onmouseout="changeContentPure(td1,msg100)" >Volunteer</a>'
menurod[4]='<a href="donate.asp?name=rodents" onmouseover="changeContentPure(td1,link5rod)"  onmouseout="changeContentPure(td1,msg100)" >Donate</a>'

var menumem=new Array()
menumem[0]='<a href="http://barbaradeeb.org/BarbaraDeeb.org/index.html" onmouseover="changeContentPure(td1,link1mem)"  onmouseout="changeContentPure(td1,msg100)">Memorial Fund</a>'
menumem[1]='<a href= "http://barbaradeeb.org/BarbaraDeeb.org/index.html" onmouseover="changeContentPure(td1,link2mem)"  onmouseout="changeContentPure(td1,msg100)">Life and Work</a>'
menumem[2]='<a href="http://barbaradeeb.org/BarbaraDeeb.org/Donate.html" onmouseover="changeContentPure(td1,link3mem)"  onmouseout="changeContentPure(td1,msg100)">Donate</a>'

var menugui=new Array()
menugui[0]='<a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/GuineaCurrent.asp" onmouseover="changeContentPure(td1,link1gui)"  onmouseout="changeContentPure(td1,msg100)" >Adoptions</a>'
menugui[1]='<a href="vets.asp?animal=4" onmouseover="changeContentPure(td1,link2gui)" onmouseout="changeContentPure(td1,msg100)">Vet Referral</a>'
menugui[2]='<a href="comingsoon.asp?name=Guinea Pig" onmouseover="changeContentPure(td1,link3gui)"  onmouseout="changeContentPure(td1,msg100)">Links</a>'
menugui[3]='<a href="volunteer.asp?name=Guinea Pigs" onmouseover="changeContentPure(td1,link4gui)"  onmouseout="changeContentPure(td1,msg100)">Volunteer</a>'
menugui[4]='<a href="donate.asp?name=guinea pigs" onmouseover="changeContentPure(td1,link5gui)"  onmouseout="changeContentPure(td1,msg100)">Donate</a>'

var menudon=new Array()
menudon[0]='<a href="donate.asp" onmouseover="changeContentPure(td1,link1don)"  onmouseout="changeContentPure(td1,msg100)" >Donate</a>'


</script>

</style>
</head>

<BODY onLoad="SLIDES.update()">

<!--main table-->

<center>

<TABLE WIDTH="775" HEIGHT=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
<tr>
<td colspan=2>
<!--#include file="headerfile.asp"-->
</td>
</tr>
<!-- </td></tr>  -->
	
<tr>
<td align="right" valign="top">

<!--table containing links and slideshow-->
<table border="0"  cellspacing="8" cellpadding="0" width=775>

<!--first row with spacer cells -->
<tr><td><img src="spacer.gif" width="130" height="25"></td>

<td rowspan=5  bgcolor="#ffffff" WIDTH="300" HEIGHT="250" align="center" valign="top">

<!--inner table with slideshow and info cell -->

<table  width="525" cellspacing="0" cellpadding="10" border="0" >
<tr>
    <td><img src="spacer.gif" height="1" width="1" border="0"></td>
    <td colspan=2><img src="spacer.gif" height="1" width="324" border="0"></td>
</tr>

<tr>
	<td rowspan="10"><img src="spacer.gif" height="240" width="1"></td>
</tr>

<tr>
    <td>
    <!-- Silent Auction -->
    <table width="100%" height="100%">

        <tr>
        <td bgcolor="#C0C0C0" align="center">
     <td bgcolor="#660099" align="center">   
<h1><img src="2010PosterforHomePage.jpg" width=90 align="right"  border=1><font color="white">
<font size="6">2010 Silent Auction & Dinner<br> <font size="5">April 24th <br>Please JOIN US</font>
<font size="6"><br>Buy Your Tickets Now!

<br><center>

<!-- Begin Auction PayPal Logo -->

<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="hosted_button_id" value="CTWTH7MEM7H6Q">
<input type="image" src="auctiontickets.jpg" border="0" name="submit" alt="PayPal - The safer, easier way to pay online!">
<img alt="" border="0" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" height="1">
</form>

<!-- End PayPal Logo -->
<font size="4">


<a href="Auction_Catalog_2010.pdf" target="_blank">Auction Catalog</a><br>

<align="left">The following companies have donated towards this year's auction. Join in the fun! Your webpage will be added when you donate to our auction: <br><br>
<align="center">
<li><a href="http://larkin-art.deviantart.com/"><font color="white"> Larkin Art</a>
<li><a href="http://www.strayframes.com/"><font color="white"> Art of John Pirtle</a>
<li><a href="http://rmcf.com/WA/Seattle50435/"><font color="white"> Rocky Mountain Chocolate Factory</a>
<li><a href="http://www.dane-dane.com"> <font color="white"> dane + dane Studios</a>
<li><a href="http://www.mightyo.com/"><font color="white"> Mighty-O Donuts</a>   
<li><a href="http://www.RideTheDucksofSeattle.com/"><font color="white"> Ride The Ducks </a>
<li><a href="http://www.Ivars.net/"><font color="white"> Ivar's </a>
<li><a href="http://www.cornish.edu/"><font color="white"> Cornish College of the Arts </a>
<li><a href="http://www.BenJerry.com/"><font color="white"> Ben & Jerry's </a>
<li><a href="http://www.MuseumofFlight.org/"><font color="white"> Museum of Flight </a>
<li><a href="http://www.thymeforhealth.com/"><font color="white"> Thyme For Health </a>
<li><a href="http://www.undergroundtour.com/"><font color="white"> Seattle's Underground Tour</a>
<li><a href="http://www.busybunny.com/"><font color="white"> Busy Bunny</a>
<li><a href="http://www.thechildrensmuseum.org/"><font color="white"> The Children's Museum</a>
<li><a href="http://www.jazzalley.com/"><font color="white"> Dimitriou's Jazz Alley</a>
<li><a href="http://www.youcouldbedancing.com/"><font color="white"> Arthur Murray School of Dance</a>
<li><a href="http://seattle.mariners.mlb.com/index.jsp?c_id=sea"><font color="white"> Seattle Mariners Tickets</a>
<li><a href="http://www.michaelgreaganartist.com"><font color="white"> Michael Reagan's Artroom </a>
<li><a href="http://northseattlevetclinic.vetsuite.com/Templates/PetPortalStyle.aspx"><font color="white"> North Seattle Vet Clinic </a>
<li><a href="http://www.comedyunderground.com"><font color="white"> The Comedy Underground </a>
<li><a href="http://www.clothworkstextiles.com"><font color="white"> Clothworks Textiles - Celebrations, Seattle </a>
<li><a href="http://www.santorinipizza.com"><font color="white"> Santorini Pizza & Pasta, 35th Ave NE</a>
<li><a href="http://www.YogaSeattle.com"><font color="white"> Center for Yoga of Seattle</a>
<li><a href="http://www.TimsChips.com"><font color="white"> Tim's Cascade Snacks</a>
<li><a href="http://www.veganme.com/"><font color="white"> Vegan*Me</a>
<li><a href="http://www.AHouseofClocks.com/"><font color="white"> A House of Clocks</a>
<li><a href="http://www.jetcityimprov.com/"><font color="white"> Jet City Improv/Wing-it Productions</a>
<li><a href="http://www.waywardvegancafe.com/"><font color="white"> Wayward Vegan Cafe</a>
<li><a href="http://www.SidecarForPigsPeace.com/"><font color="white"> Sidecar for Pigs Peace</a>
<li><a href="http://www.emerysgarden.com/"><font color="white"> Emery's Garden</a>
<li><a href="http://www.oxbowanimalhealth.com/"><font color="white"> Oxbow</a>
<li><a href="http://www.ste-michelle.com/"><font color="white"> Chateau Ste. Michelle</a>
<li><a href="http://www.conifer-inc.com/"><font color="white"> Conifer Specialties </a>
<li><a href="http://www.empsfm.org/"><font color="white"> Experience Music Project/SFM</a>
<li><a href="http://www.bunniesbythebay.com/"><font color="white"> Bunnies By the Bay</a>
<li><font color="white"> Two Classic Zunes 16GB & Dock Packs from Microsoft</a>
<li><a href="http://www.LeithPetwerks.com/"><font color="white"> Leith Petwerks</a>
<li><a href="http://www.rabbithaven.org/"><font color="white"> Rabbit Haven</a>



<br>--------- Donations from Individual Donors ---------

<li><font color="white"> iPod Classic 160GB</a>
<li><font color="white">Sony Bravia 26" LCD TV</a>
<li><font color="white">Samsung Blu-Ray Disc Player</a>
<li><font color="white">$99 Gift Certificate to Best Buy Store</a>
<li><font color="white"> Kelly Turnbull NW Artist - pen & ink drawing</a>
<li><font color="white">Teeth Whitening-Jaymor Kim, DDS</a>
<li><font color="white">Vegan Biscotti Biscuits</a>
<li><font color="white">Motorola H721 Universal Bluetooth Wireless Headset</a>
<li><font color="white">Altec Lansing Orbit Portable Speaker-Universal 3.5MM</a>
<li><font color="white">Hand Made Felted Bunny & Gift Certificate</a>
<li><font color="white">Numerous rabbit artwork by Christa</a>
<li><font color="white">Family Fun Game Basket</a>

    </td>
        </tr>
    </table>
<TR>
<TD>
<HR WIDTH="75%">
</TD>
</TR>


 <Hr width="100%">
<ul>
</ul>

    <td>
    <!-- Green Box Challenge -->
    <table width="100%">
        <tr>
        <td bgcolor="#999933" align="center">
        <h1><font color="white">Will You Help us Pay for our New Well?</font></H1>
        </td>
        </tr>
    </table>
    
    <table width="100%">
         <tr>
          <td>
         <font size="4"> 
         <center>
         <b><a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/DroughtHitsRabbitMeadows.html">Drought Hits Rabbit Meadows</a></b> 
         </center>
         <br />
             </tr>
    </table>
    
    <!-- Well pics-->
    <left>
    <table>
        <tr>
        <td bgcolor="white">
	    <!-- img src="Well.jpg" align="left" width="262" height="247" hspace="10" -->        
        <img src="thirsty_rabbits.jpg" width="70%"> <!-- Pump.jpg -->    
        <br>The balance due on the loan we took out for the new well is down to $8,000. A very kind rabbit lover loaned us the money which allowed us to have a new well drilled, and we have done fairly well paying it down. However, the balance must be paid back asap. Whatever amount you contribute will be put towards the loan balance. (Indicate that your donation is for the well.) You can use paypal or 
        mail your check to: Rabbit Meadows, 14317 Lake City Way NE, Seattle WA 98125</font>

        </td>                
        </tr>        
    </table>
    </left>
    
    <table>        
        <tr>
  ...      
       <td>
...        
        </td>
        
<center>
<!-- Begin PayPal Logo -->
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">

<input type="hidden" name="cmd" value="_xclick">
<input type="hidden" name="business" value="Sandi@RabbitRodentFerret.org">
<input type="hidden" name="no_shipping" value="1">
<input type="hidden" name="shipping" value="0.00">
<input type="hidden" name="tax" value="0">

<input type="hidden" name="return" value="http://www.washingtonhouserabbitsociety.org/thanks">
<input type="hidden" name="cancel_return" value="http://http://www.washingtonhouserabbitsociety.org/cancel">
<input type="image" src="http://images.paypal.com/images/x-click-but04.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!">
</form>
<!-- End PayPal Logo -->
</center>        
</td></tr></table>

<br>

	<tr><td bgcolor="#999933">
	<p><font color="#FFFFFF" size="4">We need <font size="5"><B>your help</B></font> to complete <b>Rabbit Meadows Sanctuary</b> project.  Read more about our plans and how you can be a part <font color="#FFFFFF"><a href="rabbitmeadowssanctuary.asp">here. </a></font></td></tr>
	<tr><td><img src="spacer.gif height="20" width="1" border="0"></td></tr>
	
	<tr><td><H3 align="center"><a name="jack">MEET JACK,</a> Bunny of the Month</H3>
	<img align="left" src="bunnyimages/Jack066.jpg" width="227" height="246" hspace="5" vspace="5"> Jack was dumped along with 23 other badly treated rabbits in a remote part of King County. All four feet were so badly abscessed & his feet were so painful that Jack could not stand for two months after he came to us. Daily soaks, antibiotics and lots of TLC brought about a complete cure for 3 of his feet. We thought he was well enough to be neutered and scheduled that appointment. We would continue to work on his left front paw. However, when he came back from being neutered, his left foot was wrapped in a bandage. I asked the vet what on earth happened, and he said that he'd decided to clean out the abscess. When I protested that I hadn't given permission for that, he said "well I didn't charge you for it!" As a result of this unasked for treatment, Jack's abscess then spread into the joint and up into the bone. We were left with no option, except to amputate that leg.<p>With all that Jack's been through, he's still a "love bug" wanting lots of pets and attention.  Jack would like to have a girl friend and would love to have his very own home.    <b>Maybe you're the one to give him that chance? </B><a href="donate.asp"><b>Donate to Jack's Medical Bills</b></a>	
	</td></tr>

	<tr><td align="left" colspan="2"><br><font color="black" size="5"><b><img src="http://www.rabbitrodentferret.org/rabbitrodentferret.org/products/condos.jpg" width="125" height="124" align=left>  Condos Now Available!!</b></font><p>
	<p>We now have the Leith Petwerks condos (48") available at our Seattle Store! Make the trip to Seattle and save on expensive shipping fees. We do not ship this item, so if you are not in Seattle you can order directly from Leith Petwerks <a href="http://www.leithpetwerks.com">www.www.leithpetwerks.com</a>
	<p>You can still get <B>T-shirts</B>, mugs, and other items with this year's beautiful logo design at <a href="http://www.cafepress.com/blrrfh">www.cafepress.com/blrrfh</a>.  Proceeds benefit the animals!  See clips from our <B>bunny photo shoot</B> at the sanctuary <a href="photoshoot.asp">here</a>.
		
	</td></tr>	
</td></tr>
</table>

<!--end of slideshow rowspan section -->

</td>
</tr>
<!--end of first row with spacer cells^-->

<tr>

<td align="right" valign="top">

<!--beginning of row 1, rabbit section (all section cells are a main 2 cells inside, a title cell and main info cell containing a table) -->
<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltgreen" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkgreen" align="center" height="15"><a class="dkgreen" href="donate.asp"  onMouseover="changeContentPure('td1',msg2don)" onMouseout="changeContentPure('td1',msg100);">Donate</a></td></tr>

		<tr><td id="td10" bgcolor="#FFFFFF" height="40">
				
				<img src="pics/donatelink.gif" align="left" border="0" width="41" height="46" vspace="2"> <script type="text/javascript">
document.write (msg1don) </script>
				
		</td></tr>
		</table>
	</td></tr>
</table><br>

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="1" border="0">
	<tr><td valign="top">
		<table class="ltorange" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkorange" align="center" height="15"><a class="dkorange" href="javascript://nothing" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menurab, '90px'); changeContentPure('td1',msg2rab)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Rabbits</a></td></tr>

		<tr><td id="td2" bgcolor="#FFFFFF" height="40">					
				<img src="pics/rabadoptlink.jpg" align="left" border="0" vspace="2"> <script type="text/javascript">
document.write (msg1rab)</script>
				
		</td></tr>
		</table>
	</td></tr>
</table>	

<br>


<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltblue" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkblue" align="center" height="15"><a class="dkblue" href="javascript://nothing" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menurod,'90'); changeContentPure('td1',msg2rod)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Rodents</a></td></tr>

		<tr><td id="td7" bgcolor="#FFFFFF" height="40">
				
				<img src="pics/rodentlink.jpg"  align="left" border="0" width="42" height="41" vspace="2"> <script type="text/javascript">
document.write (msg1rod) </script>
				
		</td></tr>
		</table>
	</td></tr>
</table><br>

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltgreen" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkgreen" align="center" height="15"><a class="dkgreen" href="javascript://nothing" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menugui,'90px'); changeContentPure('td1',msg2gui)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Guinea Pigs</a></td></tr>

		<tr><td id="td9" bgcolor="#FFFFFF" height="40">
				
				<img src="pics/guinealink.jpg" align="left" border="0" width="45" height="42" vspace="2"> <script type="text/javascript">
document.write (msg1gui) </script>
				
		</td></tr>
		</table>
	</td></tr>
</table><br>


<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="1" border="0">
	<tr><td valign="top">
		<table class="ltpurple" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkpurple" align="center" height="15"><a class="dkpurple" href="javascript://nothing" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menumea,'90px'); changeContentPure('td1',msg2mea)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Rabbit Meadows</a></td></tr>

		<tr><td id="td4" bgcolor="#FFFFFF" height="40">
								<img src="pics/rmslink.jpg" align="left" border="0" vspace="2"> <script type="text/javascript">
document.write (msg1mea) </script>
				
		</td></tr>
		</table>
</td></tr>
</table><br>

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltblue" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkblue" align="center" height="15"><a class="dkblue" href="store.asp"  onMouseover="changeContentPure('td1',msg2sto)" onMouseout="changeContentPure('td1',msg100)">Store</a></td></tr>

		<tr><td id="td6" bgcolor="#FFFFFF" height="40">				
				<img src="pics/storelink.jpg" align="left" border="0" width="41" height="44" vspace="2"> <script type="text/javascript">
document.write (msg1sto) </script>
				
		</td></tr>
		</table>
</td></tr>
</table>

<br>
<!-- Ferret section should be going out soon-->
<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="1" border="0">
<tr><td valign="top">
		<table class="ltpurple" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkpurple" align="center" height="15"><a class="dkpurple" href="javascript://nothing" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menufer,'90px'); changeContentPure('td1',msg2fer)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Ferrets</a></td></tr>

		<tr><td id="td5" bgcolor="#FFFFFF" height="40">				
				<img src="pics/ferretlink.jpg" align="left" border="0" width="41" height="44" vspace="2"><script type="text/javascript">
				                                                                                             document.write(msg1fer) </script>				
		</td></tr>
		</table>
</td></tr>
</table><br>


<!--end of rabbit link section-->
</td>
<td align="left">

</td></tr>

<tr><td align="right">

</td>
<td>
<!--store section -->


<!--end of store section and end of row 2 -->
</td></tr>

<tr><td align="right">



</td>
<td>

<!--beginning of row 2, rabbit meadows link section -->


<!--end of rabbit meadows section -->



</td></tr>
<tr><td align="right">

<!--beginning of row 4, Guinea Pig section -->



<!--end of Guinea Pig section -->
</td>
<td align="left">

<!--beginning scholarship section 

<table bgcolor="#FFFFFF" width="130" height="65" cellpadding="0" cellspacing="0" border="0">
	<tr><td valign="top">
		<table class="ltorange" cellpadding=4 cellspacing=2 width="130">
		<tr><td class="dkorange" align="center" height="15"><a class="dkorange" href="javascript://nothing" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menumem,'90px'); changeContentPure('td1',msg2mem)" onMouseout="changeContentPure('td1',msg100);delayhidemenu()">Memorial Fund</a></td></tr>

		<tr><td id="td8" bgcolor="#FFFFFF" height="40">
				
				<img src="pics/Barbdeeblink.jpg" align="left" border="0" width="43" height="42" vspace="2"> <script type="text/javascript">
document.write (msg1mem) </script>
				
		</td></tr>
		</table>
	</td></tr>
</table>

  end of scholorship section and end of row 3 -->



</td></tr>

<!-- deleted section here -->
</table></td>

<td valign="top" bgcolor="white">
	<table cellpadding="0" cellspacing="0" border="0" bgcolor="white" width="145">

	<tr><td  align="left">
	<img src="spacer.gif" width="301" height="1" border="0" alt="auction for BLRRFH"></td></tr>

	<tr><td>
	<form method="post" action="http://oi.vresp.com?fid=e001acedbb" target="vr_optin_popup" onsubmit="window.open( 'http://www.verticalresponse.com', 'vr_optin_popup', 'scrollbars=yes,width=600,height=450' ); return true;" >
  <div style="font-family: verdana; font-size: 11px; width: 270px; padding: 5px; border: 1px solid #000000; background: #dddddd;">
    <strong><span style="color: #333333;">Sign Up For our Monthly Newsletter!</span></strong><p>
    
    <label style="color: #333333;">Email Address:</label>    <span style="color: #f00">* </span>
    <br/>
    <input name="email_address" size="25" style="margin-top: 5px; margin-bottom: 5px; border: 1px solid #999; padding: 3px;"/>
    <br/>
    <label style="color: #333333;">Name:</label>
    <br/>
    <input name="first_name" size="25" style="margin-top: 5px; margin-bottom: 5px; border: 1px solid #999; padding: 3px;"/>
    <br/>
    <input type="submit" value="Join Now" style="margin-top: 5px; border: 1px solid #999; padding: 3px;"/>
  </div>
</form>

	
<hr>
</td></tr>
	<td class="volunteerforboard"><p class="philossml">
<img src="carrot017.gif" align=left>
 <font color="black"><b><center>WISH LIST</b> <br>For our New Shelter: <br>Small Farm Tractor with front loader and blade <br> Lumber: 2x4's and 2x6'<br>Nailes & screws<br>Electrical Supplies</font> </b>
	</center>
	</td></tr>

<tr><td bgcolor="black">
			<table cellpadding="5" bgcolor="black" border="0">
			<tr><td bgcolor="white" align="left"><p align="center">

			<b>Bunnycam coming soon to this spot!</b><p align="center">
			<a href="Bunnyimages/snowbunny.jpg" border="0"><img src="Bunnyimages/snowbunny.jpg" width="220" height="165"></a><br><b><font size="4"><p align="center">In the meantime, enjoy this melting snow bunny!</a></font></B> <br><p>

<p>

                        		</td></tr>
			</table>
	</td></tr>	
	
	</table>





</td></tr>


<tr><td colspan=4>
	<center><table border="0" cellspacing="7" cellpadding="5" >
		
	<td class="philosophy1" width="250" valign="top">
 <p class="philos"> <span class="orange">T</span>he welfare of all rabbits and rodents are our primary consideration. We believe that all are <b>equally valuable</b> regardless of breed purity, temperament, state of health, or relationship to humans.</td>
	
   
	<td class="philosophy2" width="250" valign="top">
 <p class="philos"><span class="purple">I</span>t is in the best interest of rabbits and rodents to be <b>neutered/spayed</b>, to live in human housing where <b>supervision</b> and <b>protection</b> are provided, and to be <b>treated for illnesses</b> by veterinarians.
 </td>

 
<td class="philosophy3" width="250" valign="top">
 <p class="philos"><span class="green">R</span>abbits and rodents are <b>intelligent and social</b> animals who require mental stimulation, toys, exercise, and social interaction from humans and other animals.

 </td></tr></TABLE></center>
 </td></tr>


<tr ><td colspan=4 align="center">


<%
dim fso
dim getcount, countfile, oldcount, count
dim latestcount
set fso=server.createobject _
("scripting.filesystemobject")
countfile="C:\sites\Single16\hrabbit\database\countmain.txt"

set getcount=fso.opentextfile(countfile,1,false)
oldcount=trim(getcount.readline)
count=oldcount+1
set latestcount=fso.createtextfile(countfile,True)
latestcount.writeline(count)

Function addcount(whichfile)
dim getcount, lastcount, number, newcount, latestcount
set getcount=fso.opentextfile(whichfile,1,false)
lastcount=trim(getcount.readline)
newcount=lastcount+1
set latestcount=fso.createtextfile(whichfile,True)
latestcount.writeline(newcount)
End Function

dim address
address=request.servervariables("HTTP_HOST")

address=lcase(address)
If address= "www.rabbithouse.org" then
call addcount("C:\sites\Single16\hrabbit\database\rabbit.txt")
Elseif address="www.rodenthouse.org" then
call addcount("C:\sites\Single16\hrabbit\database\rodent.txt")
Elseif address="www.ferrethouse.org" then
call addcount("C:\sites\Single16\hrabbit\database\ferret.txt")
elseif address="www.rabbitrodentferret.org" then
call addcount("C:\sites\Single16\hrabbit\database\rarofe.txt")
'Elseif address=("hrabbit.brinkster.net") then
'Response.Redirect "http://www.washingtonhouserabbitsociety.org"
end if
%>

<!-- table Thank you -->
<table width="700" cellspacing="0" cellpadding="0" border="0">
<tr>
<td align="center" valign="middle" class="philosophy2">
<font face="arial, arial" size=4 color=#000000><B>THANK YOU FOR SUPPORTING US!</B></font><br><b>ALL</b> donations to <b>Rabbit Meadows</b> (BLRRFH) go directly to help support our rabbits, rodents and guinea pigs.
</td>
</tr>

<tr>
<td align="center">
<br><Font face=arial color=#999933 size=3><b>You are visitor number <% Response.Write count %>!  </b></font><br>&nbsp;
</td>
</tr>

<tr>
<td>
<!--#include file="footer.asp"-->
</td>
</tr>

</table>
<!--end of inner table with links and slideshow -->

</td>

</tr>

</table>
</center>

</body>
</html>