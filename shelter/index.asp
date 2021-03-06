﻿<% @LANGUAGE=VBScript %>
<% Option Explicit %>

<!-- #include file="correct-domain.asp"-->

<!-- #include file="paypal_paybutton.asp"-->
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
    MyMail.AddAddress "sandi@rabbitmeadows.org"
    MyMail.Subject = "Newsletter Sign Up"
    MyMail.AddCC "info@rabbitmeadows.org"

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

<meta name="description" content="Rabbit Meadows, Sanctuary and Adoption Center"/>
<meta name="keywords" content="rabbit, rodent, pet, companion animal, shelter, boarding, kennel, pet supplies, pet food,
   squirrel, possum, guinea pig, rat, mice, chinchilla, prairie dog, adoption, washington, seattle, puget, house rabbit, rabbit, rabbits, pet, bunny, bunnies,
care, breed, breeding, breeds, Humane Society, education, adoption, adopt, non-profit,Fuzzy Lop, Holland lop, mini lop, fench lop, rex, giant, dwarf, new zealand, lion head, jersey wooley,
	behavior, faq, spay, neuter, animals, lapin, lapine, sanctuary, rabbit sanctuary, woodland park, adoption "/>

<meta name="copyright" content="Copyright 2000-2011 Rabbit Meadows, Sanctuary & Adoption Center. All rights reserved.
        Contact author for reprint policies."/>

<link rel="StyleSheet" href="style.css" type="text/css" media="screen"/>
<!--#include file="sandiedit.js"-->

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

<!--#include file="dropdownmenu.asp"-->

<!--#include file="google-analytics.js"-->

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
<body>

<center>
<!--#include file="headerfile.asp"-->
_________________________
<table border="0"  cellspacing="10px">
    <tr>
      <td align="right" valign="top">
        <!--#include file="sidebar_left.asp"-->
      </td>
      <td style="width:540px">
        <h1 class="banner" style="text-align:center">
            Memorialize Your Companion Animal on our New Shelter Walls
        </h1>
        <img src="BunnyImages/Tiles-small.jpg" alt="Tile Examples" style="float:right" />
        <p>
        Have you heard about the new shelter we're building in Redmond? (See below.) We need <i>your</i> support to bring this project to completion.
        The walls of our new shelter will be adorned with photos of your companion animals on ceramic tiles.  If you want your furry friend to be
        forever memorialized on our shelter walls, please use the Paypal button below, or fill out
        <a href="Photo Tile Requirements.pdf">this form</a> and send it in with your donation.
        </p>
        <center>
        <%
                dim choices(7)
                choices(0) = "4x4"
                choices(1) = 25
                choices(2) = "3x6"
                choices(3) = 30
                choices(4) = "6x6"
                choices(5) = 50
                choices(6) = "6x8"
                choices(7) = 75

                paypal_multi "Tile", choices
        %>
        </center>
        <p>
        Rabbits, hamsters, cats, dogs, and critters of all descriptions are welcome on our walls!
        </p>
        <p>
        We'll need a high-quality digitial photo of your furry friend to create the ceramic tile.  If your digital camera has a "best" setting, use that.
        If you just have an image sitting on your hard drive, for a good photo-quality print, you should make sure it's a high enough resolution for the
        tile size you select.  Here's what we recommend:
        </p>
        <ul>
            <li>4x4 -- 800x800</li>
            <li>3x6 -- 600x1200</li>
            <li>6x6 -- 1200x1200</li>
            <li>6x8 -- 1200x1600</li>

        </ul>
        <p>
        The critters of Rabbit Meadows thank you for your support!
        </p>
      <hr/>
              <h2 style="border-style:groove; text-align:center">Our Shelter is going up!</h2>
    Look!! Our New Entrance Gate - Artwork created & donated by Chris Hart (future Eagle Scout) and his dad! Thank YOU!!<br>

      	<img src="New Gate2.jpg" width="288" height="230" hspace="10" alt="Shelter Progress" style="float:left" />
    	<img src="New Gate3.jpg" width="288" height="240" hspace="10" alt="Shelter Progress" style="float:right" />

      	<img src="ShelterFeb2012a.jpg" width="300" height="180" hspace="10" alt="Shelter Progress" style="float:left" />
    	<img src="ShelterFeb2012b.jpg" width="252" height="189" hspace="10" alt="Shelter Progress" style="float:right" />
        Progress: Our new shelter is taking shape!
                The building is 38' x 48' with a second level loft of 14' x 48'. The first level will contain our reception area; adoption area with a place to spend time with potential companions; an education room; a boarding room; utility room; bathroom.
                The loft area will contain areas to quarantine incomming animals and those waiting to be spayed/neutered. And, a clinic room where animals can receive veterinary care.
                <br><br>The building now has siding on the back and north side ground floor. We have one "half-round" window in the front, but still need another 4'x2' half-round for the left side of the door. (Thanks to Photoshop you can see what it will look like in the future.)
				 Other than that, we have all the ground floor windows and only need four more 6'x2' windows for upstairs. We have the gutters and down spouts
				 but they need to be installed. We have all of the siding and have a potential volunteer to install it. We also need a 20' x 10' asphalt apron if you know of anyone with that experience who might help, please ask them to contact us.
			<br><br><b>And then the rabbits, guinea pigs, rats, mice, etc. can move in!</b>
            <br><br>The address is (approximately) 8510 250th Ave NE, Redmond, WA 98053. (Address still to be assigned by the county.)  Please be considerate of our neighbors and drive no more than 20 mph when you turn into our private road.
                <a href="mailto:Sandi@RabbitMeadows.org?subject=Helping%20With%20the%20New%20Shelter">Contact us</a> if you have a few hours available to help.

          <center><hr />
        <img src="Shelter-JanetMatthew.jpg" width="300" height="225" hspace="10" alt="Shelter Progress-B" style="float:right" />

          <b>Past Newsletters:</b>
			<br/><a href="/shelter/Newsletters/March2012.pdf">March 2012</a>
			<br/><a href="/shelter/Newsletters/February2012.pdf">February 2012</a>
			<br/><a href="/shelter/Newsletters/January2012news.pdf">January 2012</a>
			<br/><a href="/shelter/Newsletters/NoseWigglesfromRabbitMeadows.htm">December 2011</a>
			 <br/><a href="http://hosted.verticalresponse.com/258167/f746d9532b/1371019405/b25efa5963/">October 2011</a>
           <br/><a href="http://hosted.verticalresponse.com/258167/a2995f900d/1371010277/b25efa5963/">April 2011</a>
           <br/><a href="http://hosted.verticalresponse.com/258167/6a5c7b27a3/1371010277/d924055a98/">March 2011</a>
          <br/><a href="/shelter/Newsletters/RabbitMeadowsCelebratesYearoftheRabbit.htm">January 2011</a>
          <br/><a href="/shelter/Newsletters/NoseWigglesfromRabbitMeadows.htm">December 2010</a>
          <br/><a href="/shelter/Newsletters/News from Rabbit Meadows.htm">November 2010</a>

          <div style="padding:10px"/>
          <% paypal_donatebutton %>
        </center>

        <div class="banner" style="font-size:1.1em">
          We need <strong style="font-size:1.25em">your help</strong> to complete <strong>Rabbit
          Meadows Sanctuary</strong> project.  Read more about our plans and how you can be a part
          <a href="rabbitmeadowssanctuary.asp">here.</a>
        </div>

        <hr/>

        <h3 align="center">MEET Jumpin' Jack & Popcorn, Bunnies of the Month</h3>
        <p>
          <img align="left" src="Web-JumpinJack-Popcorn.jpg" width="198" height="163" hspace="5" vspace="5/ alt="JumpinJack+Pepper"">
          These adorable bunnies were adopted for only 3 weeks and then returned when the previously allergy-free individuals came down with an allergy to cleaning litterboxes. Jumpin' Jack & Popcorn were originally turned into Seattle Animal Shelter by a breeder who hadn't been able to sell them at Easter. Now, almost two years old, these brothers would sure like a home to call their own.
        </p>

        <hr/>

      <img src="/shelter/condo.jpg" width="125" height="124" align="left" alt="Condo"/>
      <span style="font-weight:bold; font-size:1.35em">Condos Now Available!!</span>
	<p>We now have the Leith Petwerks condos (48&quot;) available at our Seattle Store! Make the trip to Seattle and save on expensive shipping fees. We do not ship this item, so if you are not in Seattle you can order directly from Leith Petwerks <a href="http://www.leithpetwerks.com">www.leithpetwerks.com</a>
	 Proceeds benefit the animals!</p>

      <hr/>

      See clips from our <strong>bunny photo shoot</strong> at the sanctuary <a href="photoshoot.asp">here</a>.

      <td valign="top" bgcolor="white">
        <!--#include file="sidebar_right.asp"-->
	  </td>
    </tr>

    <tr>
        <td colspan="3">
	        <center>
                <table border="0" cellspacing="7" cellpadding="5" width="750px">
                    <tr>
                    	<td class="philosophy1" width="250px" valign="top">
                            <p class="philos">
                            <span class="orange">T</span>he welfare of all rabbits and rodents are our primary consideration. We believe that all are <b>equally valuable</b> regardless of breed purity, temperament, state of health, or relationship to humans.
                            </p>
                         </td>
	                    <td class="philosophy2" width="250px" valign="top">
                            <p class="philos">
                            <span class="purple">I</span>t is in the best interest of rabbits and rodents to be <b>neutered/spayed</b>, to live in human housing where <b>supervision</b> and <b>protection</b> are provided, and to be <b>treated for illnesses</b> by veterinarians.
                            </p>
                        </td>
                        <td class="philosophy3" width="250px" valign="top">
                            <p class="philos">
                            <span class="green">R</span>abbits and rodents are <b>intelligent and social</b> animals who require mental stimulation, toys, exercise, and social interaction from humans and other animals.
                            </p>
                        </td>
                    </tr>
                </table>
            </center>
        </td>
    </tr>
    <tr>
    <td colspan="3" align="center">
        <table width="700" cellspacing="0" cellpadding="0" border="0">
            <tr>
                <td align="center" valign="middle" class="philosophy2">
                    <font face="arial" size="4" color="#000000"><B>THANK YOU FOR SUPPORTING US!</B></font><br/>
                    <b>ALL</b> donations to <b>Rabbit Meadows</b> go directly to help support our rabbits, rodents and guinea pigs.
                </td>
            </tr>
            <tr>
                <td align="center">
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
                    %>
                    <br/><font face="arial" color="#999933" size="3"><b>You are visitor number <% Response.Write count %>!  </b></font><br/>&nbsp;
                </td>
            </tr>
        </table>
    </td>
    </tr>
</table>
<!--#include file="footer.asp"-->
</center>

</body>
</html>