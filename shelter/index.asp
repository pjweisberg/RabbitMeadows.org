<% @LANGUAGE=VBScript %>
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
See Photos of our successful <a href="http://www.facebook.com/RabbitMeadows#!/RabbitMeadows#!/photo.php?fbid=10150167349094759&set=pu.75949809758&type=1&theater">Easter Egg Hunt</a> on Facebook
<table border="0"  cellspacing="10px">
    <tr>
      <td align="right" valign="top">
        <!--#include file="sidebar_left.asp"-->
      </td>
      <td style="width:540px">
        <h1 class="banner" style="text-align:center">
            Annual Silent Auction & Dinner<br/>
            Saturday  June 4, 2011 <br/>
            <span style="font-size:.75em">Sponsored by Rusty Pelican Cafe</span>
        </h1>
 <a href="2011 Poster.pdf"><img src="2011PosterSmall.jpg" alt="2011 Auction Poster" style="float:right;" /></a>
<center><a href="AuctionDinner2011.asp">See our current list of Auction Items</a></center>
<p>
Our 5th annual Silent Auction is fast approaching! Tickets are only $35 for the auction and a delicious pasta or lasagna dinner, plus a fantastic dessert that Rusty Pelican is famous for. We have a limited number of tickets, so purchase yours now. 
</p>
<center>
          <%
            paypal_paybutton "Seat at Rabbit Meadows Dinner & Silent Auction", "35.00"
          %>
</center>
<p>
    We are still accepting items for the auction. If you have a service (house cleaning, pet sitting, etc.) or product that you can donate, please fill out and send us
    <a href="2011 Auction Letter & Procurement Form.pdf">this form</a> along with any brochures, buisness cards, etc. you'd like to have displayed alongside the item at the auction.
    You can help to make this year's auction a huge success. The rabbits, guinea pigs, rats & other small rodents thank you!
</p>
<p>
    <strong>Help us spread the word!</strong>  Print out <a href="2011 Poster.pdf">our poster</a> and hang it on your office door, the bulletin board in the break room, or anywhere
    else <em>that you have permission</em> to post it!
</p>

               <p>            
    <!--        <a href="2011Auction.asp">Click here</a> for more information or email
            <a href="mailto:Auction@RabbitMeadows.org?subject=Silent%20Auction%20& Dinner%20">Auction@RabbitMeadows.org</a>
            or call 206-365-9105.
    -->
        </p>
        <hr/>
        
        <h2 style="border-style:groove; text-align:center">We Have Our Building Permit!</h2>
      	<img src="Web-bldg1.jpg" width="360" height="185" hspace="10" alt="Conceptual image of new shelter" style="float:right" />
        Here's an idea of what our new shelter will look like!  
                The building will be 38' x 48' with a second level loft of 14' x 48'. The first level will contain our reception area; adoption area with a place to spend time with potential companions; an education room; a boarding room; utility room; bathroom.
                The loft area will contain areas to quarantine incomming animals and those waiting to be spayed/neutered. And, a clinic room where animals can receive veterinary care. We need volunteers to help clear the building area. <a href="mailto:Sandi@RabbitMeadows.org?subject=Helping%20With%20the%20New%20Shelter">Contact us</a> if you have a few hours available to help.
<hr />

        <center>
          <b>Past Newsletters:</b>
          <br/><a href="/shelter/Newsletters/News from Rabbit Meadows.htm">November 2010</a>
          <br/><a href="/shelter/Newsletters/NoseWigglesfromRabbitMeadows.htm">December 2010</a>
          <br/><a href="/shelter/Newsletters/RabbitMeadowsCelebratesYearoftheRabbit.htm">January 2011</a>
           <br/><a href="/shelter/Newsletters/HoponoverMarch2011.htm">March 2011</a>
           <br/><a href="/shelter/Newsletters/FunWaystoSupportRabbitMeadows.htm">April 2011</a>
   
          <div style="padding:10px"/>
          <!--#include file="paypal_logo.html"-->
        </center>

        <div class="banner" style="font-size:1.1em">
          We need <strong style="font-size:1.25em">your help</strong> to complete <strong>Rabbit
          Meadows Sanctuary</strong> project.  Read more about our plans and how you can be a part
          <a href="rabbitmeadowssanctuary.asp">here.</a>
        </div>

        <hr/>
        
        <h3 align="center">MEET Pepper, Bunny of the Month</h3>
        <p>
          <img align="left" src="bunnyimages/Pepper-72.jpg" width="198" height="163" hspace="5" vspace="5/ alt="Pepper"">
          Pepper was found as a stray in Bellingham. He was likely last year's "Easter Bunny" and was discarded when the "kids lost interest" or "he started spraying/marking everything!" Now neutered, Pepper is the life of his party and wants lots of attention.  <br><br>Come visit Pepper at our Redmond Shelter. He's looking for someone who will pay lots of attention to him and let him out into the house for several hours of exercise each day.
        </p>

        <hr/>

      <img src="/shelter/products/condos.jpg" width="125" height="124" align="left" alt="Condo"/>
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
                    <b>ALL</b> donations to <b>Rabbit Meadows</b> (BLRRFH) go directly to help support our rabbits, rodents and guinea pigs.
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
