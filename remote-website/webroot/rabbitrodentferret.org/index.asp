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


<html>
<head>
<title>Rabbit Meadows - Quality Supplies for Companion Animals</title>

<meta name="description" content="Best Little Rabbit & Rodent House Pet Supplies and Adoption"/>
<meta name="keywords" content="rabbit, rodent, pet, companion animal, shelter, boarding, kennel, pet supplies, pet food,
   squirrel, possum, guinea pig, rat, rodent, mice, chinchilla, prairie dog, adoption, washington, seattle, puget, house rabbit, rabbit, rabbits, pet, bunny, bunnies, 
care, breed, breeding, breeds, Humane Society, education, adoption, adopt, non-profit,
	behavior, faq, spay, neuter, animals, lapin, lapine, sanctuary, rabbit sanctuary, woodland park, DEA "/> 

<meta name="copyright" content="Copyright 2000-2011 Best Little Rabbit, Rodent and Ferret House. All rights reserved. 
        Contact author for reprint policies."/>

<link rel="StyleSheet" href="style.css" type="text/css" media="screen"/>
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

<script type="text/javascript">
<!--hide
    function legendpewindow(page)
    {
        window.open(page,'rescuewin','width=620,height=650,scrollbars=yes,resizable=yes');
    }
//-->
</script>


<script type="text/javascript">
    function validate() {

        if (document.newssignup.email.value.length < 5) {
            alert("Please enter your e-mail address so we can send you the newsletter!");
            return false;
        }
        window.open('newsletterthankyou.asp','thankyouwin','width=320,height=220,scrollbars=yes,resizable=yes');
        return true;
    }
</script>


<script type="text/javascript">
<!--hide
    function thankyouwindow(page)
    {
        window.open(page,'thankyouwin','width=600,height=600,scrollbars=yes,resizable=yes');
    }
//-->
</script>

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



<script type="text/javascript" src="slideshow.js">
</script> 

<script type="text/javascript">
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
</script>


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

var menumea=new Array()
menumea[0]='<a href="RabbitMeadowsSanctuary.asp" onmouseover="changeContentPure(td1,link1mea)" onmouseout="changeContentPure(td1,msg100)" >Our Sanctuary</a>'
menumea[1]='<a href="volunteer.asp?name=Sanctuary Rabbits" onmouseover="changeContentPure(td1,link2mea)" onmouseout="changeContentPure(td1,msg100)" >Volunteer</a>'
menumea[2]='<a href="donate.asp?name=sanctuary rabbits" onmouseover="changeContentPure(td1,link3mea)" onmouseout="changeContentPure(td1,msg100)" >Donate</a>'
menumea[3]='<a href="http://www.woodlandparkrabbits.org" onmouseover="changeContentPure(td1,link4mea)" onmouseout="changeContentPure(td1,msg100)" >Woodland Park Project</a>'


var menusto=new Array()
menusto[0]='<a href="store.asp" onmouseover="changeContentPure(td1,link1sto)"  onmouseout="changeContentPure(td1,msg100)" >Rabbits</a>'
menusto[1]='<a href="store.asp" onmouseover="changeContentPure(td1,link2sto)"  onmouseout="changeContentPure(td1,msg100)" >Rodents & Guinea Pigs</a>'


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

<script type="text/javascript">
 var _gaq = _gaq || [];
 _gaq.push(['_setAccount', 'UA-21632761-1']);
 _gaq.push(['_trackPageview']);
(function() {
 var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
 ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
 var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
 })();
</script>

</head>

<body onload="SLIDES.update()">

<!--main table-->

<center>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td colspan="2">
<!--#include file="headerfile.asp"-->
</td>
</tr>
	
<tr>
<td align="right" valign="top">

<!--table containing links and slideshow-->
<table border="0"  cellspacing="8" cellpadding="0" width="775">

<!--first row with spacer cells -->
<tr>
<td><img src="spacer.gif" height="25" alt=""/></td>
<td rowspan="5" bgcolor="#ffffff" width="100%" height="250" align="center" valign="top">
<hr width="100%" />
<table width="100%" cellspacing="0" cellpadding="10" border="0"/>
</td>
</tr>
<tr>
    <td><img src="spacer.gif" height="1" width="1" border="0" alt=""/></td>
    <td colspan=2><img src="spacer.gif" height="1" width="324" border="0" alt=""/></td>
</tr>

<tr>
	<td rowspan="10"><img src="spacer.gif" height="240" width="1" alt=""/></td>
</tr>
</td>
<tr>
<td>
<hr width="75%"/>
</td>
</tr>


    <td>
    <table width="100%">
        <tr>
        <td bgcolor="#999933" align="center">
        <h1><font color="white">Celebrating Chinese Year of the Rabbit! 2011</font></h1>
        </td>
        </tr>
    </table>
        
    <table>
        <tr>
        <td bgcolor="white">
	    <img src="Web-bldg1.jpg" align="right" width="360" height="185" hspace="10" alt="Conceptual image of new shelter" style="float:right" />
        <br/>Here's an idea of what our new shelter will look like! Permit application was submitted to King County in Dec 2010. It's a slow process, but it's moving along. 
                The building will be 38' x 48' with a second level loft of 14' x 48'. The first level will contain our reception area; adoption area with a place to spend time with potential companions; an education room; a boarding room; utility room; bathroom.
                The loft area will contain areas to quarantine incomming animals and those waiting to be spayed/neutered. And, a clinic room where animals can receive veterinary care. 2-14-11 (More info soon.)

        </td>                
        </tr>        
    </table>
<hr />        
<table width="100%">
         <tr>
          <td>
         <center>
         <font size="4">   
     <b>Past Newsletters:</b>
<br/><a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/Newsletters/News from Rabbit Meadows.htm">November 2010</a>
<br/><a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/Newsletters/NoseWigglesfromRabbitMeadows.htm">December 2010</a>
<br/><a href="http://www.rabbitrodentferret.org/rabbitrodentferret.org/Newsletters/RabbitMeadowsCelebratesYearoftheRabbit.htm">January 2011</a>
</font>
</center>

<!-- Begin PayPal Logo -->
<form action="https://www.paypal.com/cgi-bin/webscr" method="post">

<input type="hidden" name="cmd" value="_xclick"/>
<input type="hidden" name="business" value="Sandi@RabbitRodentFerret.org"/>
<input type="hidden" name="no_shipping" value="1"/>
<input type="hidden" name="shipping" value="0.00"/>
<input type="hidden" name="tax" value="0"/>

<input type="hidden" name="return" value="http://www.rabbitmeadows.org/"/>
<input type="hidden" name="cancel_return" value="http://www.rabbitmeadows.org/"/>
<input type="image" src="http://images.paypal.com/images/x-click-but04.gif" name="submit" alt="Make payments with PayPal - it's fast, free and secure!"/>
</form>
<!-- End PayPal Logo -->
 
</td></tr></table>

<br/>

	<tr><td bgcolor="#999933">
	<p><font color="#FFFFFF" size="4">We need <font size="5"><B>your help</B></font> to complete <b>Rabbit Meadows Sanctuary</b> project.  Read more about our plans and how you can be a part <font color="#FFFFFF"><a href="rabbitmeadowssanctuary.asp">here. </a></font></td></tr>
	<tr><td><img src="spacer.gif" height="20" width="1" border="0"/></td></tr>
    
    <tr><td><H3 align="center"><a name="Pepper">MEET Pepper</a> Bunny of the Month</H3>
      <img align="left" src="bunnyimages/Pepper-72.jpg" width="198" height="163" hspace="5" vspace="5"> Pepper was found as a stray in Bellingham. He was likely last year's "Easter Bunny" and was discarded when the "kids lost interest" or "he started spraying/marking everything!" Now neutered, Pepper is the life of his party and wants lots of attention.  <br><br>Come visit Pepper at our Redmond Shelter. He's looking for someone who will pay lots of attention to him and let him out into the house for several hours of exercise each day.
    </td></tr>

	<tr><td align="left" colspan="2"><br><font color="black" size="5"><b><img src="http://www.rabbitrodentferret.org/rabbitrodentferret.org/products/condos.jpg" width="125" height="124" align=left>  Condos Now Available!!</b></font><p>
	<p>We now have the Leith Petwerks condos (48") available at our Seattle Store! Make the trip to Seattle and save on expensive shipping fees. We do not ship this item, so if you are not in Seattle you can order directly from Leith Petwerks <a href="http://www.leithpetwerks.com">www.www.leithpetwerks.com</a>
	 Proceeds benefit the animals!  See clips from our <B>bunny photo shoot</B> at the sanctuary <a href="photoshoot.asp">here</a>.
		
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
 <font color="black"><b><center>WISH LIST For our New Shelter</b> <br>Lumber: 2x4's and 2x6'<br>Nailes & screws<br>Electrical Supplies<br>Ceiling fans(6) <br>Drywall<br>Windows<br>Gift Certificates to Home Depot, Lowes, McLendon's, Second Use, Millworks of Maple Valley</font>
	</center>
	</td></tr>

<tr><td bgcolor="black">
			<table cellpadding="5" bgcolor="black" border="0">
			<tr><td bgcolor="white" align="left"><p align="center">

			<b>Bunnycam coming soon to this spot!</b><p align="center">
			<img src="Bunnyimages/snowbunny.jpg" alt="Bunny Snowman"/><br/>
            <b><font size="4"><p align="center">In the meantime, enjoy this melting snow bunny!</font></b> <br/><p>

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

 </td></tr></table></center>
 </td></tr>


<tr ><td colspan=4 align="center">

<%
<!-- table Thank you -->
<table width="700" cellspacing="0" cellpadding="0" border="0">
<tr>
<td align="center" valign="middle" class="philosophy2">
<font face="arial" size="4" color="#000000"><B>THANK YOU FOR SUPPORTING US!</B></font><br/>
<b>ALL</b> donations to <b>Rabbit Meadows</b> (BLRRFH) go directly to help support our rabbits, rodents and guinea pigs.
</td>
</tr>
<tr>
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