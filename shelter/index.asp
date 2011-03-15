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
<table border="0"  cellspacing="10px">
    <tr>
      <td align="right" valign="top">
        <!--#include file="sidebar_left.asp"-->
      </td>
      <td style="width:540px">
        <h1 class="banner" style="text-align:center">
            <span style="font-size:0.75em">Come to Where the Real Rabbits Live for our</span><br />
            First Annual Easter Egg Hunt!
        </h1>
        <h2 style="text-align:center">Saturday, April 23, 2011: 10&nbsp;am&nbsp;&ndash;&nbsp;2&nbsp;pm</h2>
        <img src="images/bunny-in-basket.gif" alt="Easter Basket" style="float:right;" />
        <p>
            The easter bunny will be making an appearance at Rabbit Meadows this April.
            Come with your family, have your photo taken with him, search for plastic eggs with toys
            inside, and find the gold or silver eggs for a special prize!
        </p>
        <p>
            There will be three-legged races and other fun activities for all ages.  While you're here,
            you'll also get the chance to meet a few of the real live rabbits who have a home at
            Rabbit Meadows because of your support.
            Learn how <i>you</i> can help
            us in our mission to save homeless and abandoned bunny-rabbits, guinea pigs, and other furry
            critters. You can also learn the difference between the easter bunny and living breathing 
            rabbits who can live 10+ years.
        </p>
        <p>
            <a href="EggHunt2011.asp">Click here</a> for more information.  To RSVP, email
            <a href="mailto:EasterEggHunt@RabbitMeadows.org?subject=Easter%20Egg%20Hunt%20RSVP">EasterEggHunt@RabbitMeadows.org</a>
            or call 425-836-8925.  (RSVPs are appreciated, but not required.)
        </p>

        <hr/>
        
        <h2 style="border-style:groove; text-align:center">Celebrating Chinese Year of the Rabbit! 2011</h2>
      	<img src="Web-bldg1.jpg" width="360" height="185" hspace="10" alt="Conceptual image of new shelter" style="float:right" />
        Here's an idea of what our new shelter will look like! Permit application was submitted to King County in Dec 2010. It's a slow process, but it's moving along. 
                The building will be 38' x 48' with a second level loft of 14' x 48'. The first level will contain our reception area; adoption area with a place to spend time with potential companions; an education room; a boarding room; utility room; bathroom.
                The loft area will contain areas to quarantine incomming animals and those waiting to be spayed/neutered. And, a clinic room where animals can receive veterinary care. 2-14-11 (More info soon.)

        <hr />

        <center>
          <b>Past Newsletters:</b>
          <br/><a href="/shelter/Newsletters/News from Rabbit Meadows.htm">November 2010</a>
          <br/><a href="/shelter/Newsletters/NoseWigglesfromRabbitMeadows.htm">December 2010</a>
          <br/><a href="/shelter/Newsletters/RabbitMeadowsCelebratesYearoftheRabbit.htm">January 2011</a>
          <div style="padding:10px"/>
          <!-- Begin PayPal Logo -->
          <form action="https://www.paypal.com/cgi-bin/webscr" method="post">

            <input type="hidden" name="cmd" value="_xclick"/>
            <input type="hidden" name="business" value="Sandi@RabbitRodentFerret.org"/>
            <input type="hidden" name="no_shipping" value="1"/>
            <input type="hidden" name="shipping" value="0.00"/>
            <input type="hidden" name="tax" value="0"/>

            <input type="hidden" name="return" value="http://www.rabbitmeadows.org/shelter/"/>
            <input type="hidden" name="cancel_return" value="http://www.rabbitmeadows.org/shelter/"/>
            <input type="image" src="http://images.paypal.com/images/x-click-but04.gif" name="submit" alt="Make payments with PayPal - it's fast, free and secure!"/>
          </form>
          <!-- End PayPal Logo -->
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

</body>
</html>
