<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->

<!-- #include file="paypal_paybutton.asp"-->
<html>
    <head>
        <title>Charity Dinner & Silent Auction for Rabbit Meadows</title>

        <link rel="StyleSheet" href="style.css" type="text/css" media="screen"/>
        <!--#include file="sandiedit.js"-->
        <!--#include file="google-analytics.js"-->
    </head>
    <body>

    <!--#include file="dropdownmenu.asp"-->
    <!--#include file="headerfile.asp"-->
    <center>
    <table cellspacing="10px">
        <tr>
            <td valign="top" align="right">
                <!--#include file="sidebar_left.asp"-->
            </td>
            <td width="540px" style="padding:0" valign="top">
            <center>
                <a href="http://www.facebook.com/media/set/?set=a.76291579758.73220.75949809758">
                    See pictures from our 2009 Auction
                </a>
            </center>
            <br />
            <div style="border-color:Black; border-style:double; padding:5px">
                <center>
                    Our Fifth Annual<br />
                    <h1>Dinner and Silent Auction</h1>
                    <h2>Saturday, June 4, 2011 &ndash; 2 pm to 5 pm</h2>
                    <h3><a href="http://www.rustypelicancafe.com"> Rusty Pelican Café</a><br />
                    1924 N 45th St., Seattle</h3>
                </center>
                <img src="BunnyImages/Josie1.jpg" alt="Josie, our 2011 Auction Mascot" style="float:right" />
                <p>
                    For a $35 donation, you get a three course meal with a choice of Vegetarian Lasagna or Greek Pasta,
                    and of course one of the homemade desserts The Rusty Pelicans is famous for!
                </p>
                <center>
                    <p style="font-size:.75em">
                        Reserve your tickets now with:
                        <%
                            paypal_paybutton "Seat at Rabbit Meadows Dinner & Silent Auction", "35.00"
                        %>
                    </p>
                </center>
                <p>
                    There is no other shelter like Rabbit Meadows! In fact, nearly all of the small animals we take in
                    come from other shelters that simply don’t have the resources to care for their special needs.
                </p>
                <p>
                    100% of the <em>gross</em> proceeds from this event benefit the lifesaving work of Rabbit Meadows.  The venue,
                    the food, and the service, are all donated by our friends at The Rusty Pelican Café.  The auction items are gifts
                    from generous individuals and businesses.  We could never do this without their support.
                </p>
                <center>
                    <table>
                        <tr>
                            <td>
                                <img src="carrot017.gif" /> 
                            </td>
                            <td valign="middle" style="font-family:'Comic Sans MS', cursive, sans-serif; text-align:center">
                                We get by with a little help from our friends!
                            </td>
                            <td>
                                <img src="carrot017.gif" /> 
                            </td>
                        </tr>
                    </table>
                </center>
                <p>
                    Do you know a generous individual or business that might want their item added to the list below?  Forward them
                    <a href="2011%20Auction%20Letter%20&%20Procurement%20Form.pdf">this letter</a>.
                    Our rabbits and other critters couldn't be more greatful.
                    (Unless you also gave them a piece of banana or something.)
                </p>
                <center>
                    <h1>Auction Items List</h1>
                    More to be added as the date approaches
                </center>
                <ul>
                    <li>
                       <a href="http://www.petwerks.com/Cat_images/BA200.jpg" target="_blank">
                         48 in. Double Level Bunny Abode Condo</a>  
                    </li>
                    <li>
                       <a href="http://www.petwerks.com/Cat_images/MS910.jpg" target="_blank">
                         Petwerks Vacation Villa (and more from Leith Petwerks!)</a>  
                    </li>
                    <li>
                       <a href="http://www.kamakaexoticvet.com/" target="_blank">
                         Certificate for Kamaka Exotic Animal Veterinary Services (you'll love Dr. Kamaka!)</a>  
                    </li>
                    <li>
                       <a href="http://www.northseattlevetclinic.vetsuite.com/Templates/PetPortalStyle.aspx" target="_blank">
                         Certificate for North Seattle Veterinary Clinic (you'll love Dr. Nathanson!)</a>  
                    </li>

                    <li>
                       <a href="http://www.ridetheducksofseattle.com/" target="_blank">Two
                        adult admissions to Ride the Ducks</a>
                    </li>

                    <li>
                        <a href="http://www.ahouseofclocks.com/" target="_blank">House of Clocks: Seiko
                        Atomic Clock &amp; Kit Kat Clock</a>
                    </li>
                    <li>
                        MP3 Player                                                                                   
                    </li>
                    <li>
                        Champagne flutes, set of 6                                                    
                    </li>
                    <li>
                        Three Rabbits (print)                                                                
                    </li>
                    <li>
                        Bedtime Bunny Set                                                                   
                    </li>
                    <li>
                        Bread Maker                                                                                
                    </li>
                    <li>
                        Four Course Vegan Dinner
                        for Six                                       
                    </li>
                    <li>
                        <a href="http://larkin-art.deviantart.com/gallery/" target="_blank">Larkin's                        original art - Scroll down to see "hannah"</a>
                    </li>
                    <li>
                        <a href="http://www.dermagenics.com/" target="_blank">Dermagenics Skin
                        Treatment Baskets</a>
                    </li>
                    <li>
                        <a href="http://www.queensryche.com/" target="_blank">Queensryche Gift Bag</a>
                    </li>
                    <li>
                        <a href="http://www.busybunny.com/" target="_blank">Busy Bunny
                        Gift Basket</a>
                    </li>
                                       <li>
                        <a href="http://www.jetcityimprov.com/" target="_blank">Admission
                        for 4 to Jet City Seattle&#39;s best improvisational comedy show!</a>
                    </li>
                    <li>
                        <a href="http://www.museumofflight.org/" target="_blank">Admission
                        for 2 to Museum of Flight</a>
                    </li>
                    <li>
                        <a href="http://www.thechildrensmuseum.org/" target="_blank">Admission
                        for 4 to The Children&#39;s Museum</a>
                    </li>
                    <li>
                        The Chopper 2 in 1 Custom Styler                                       
                    </li>
                    <li>
                        Zoom In Office Teeth Whitening -Dentist Jaymor Kim
                    </li>
                    <li>
                        <a href="http://www.veganme.com/" target="_blank">Custom Painted
                        Animal Portrait-Vegan*Me</a> 
                    </li>
                    <li>
                        Snowshoe &amp; Lunch Trek                                                         
                    </li>
                    <li>
                        Jewelry: earrings, bracelets, etc.
                    </li>
                    <li>
                        <a href="http://www.GenerationThrive.com" target="_blank">Thrive
                        Gift Card</a>
                    </li>
                    <li>
                        <a href="http://www.custompure.com/" target="_blank">Gift
                        Certificate for Custom Pure Water filters</a>
                    </li>
                    <li>
                        <a href="http://www.reddoorspas.com/" target="_blank">Red Door
                        Spas Makeover Party for 6</a>
                    </li>
                    <li>
                        <a href="http://www.parlorcollection.com/" target="_blank">2
                        Two-Guest passes to Parlor Live Comedy Club</a>
                    </li>
                    <li>
                        Annette Funicello -
                        Collectable Ballerina Bear                
                    </li>
                    <li>
                        Silver Rabbit Poop Earrings with Bright Swirls (yes poop!)                 
                    </li>
                    <li>
                        "Bunny Blossom" print                                                                
                    </li>
                    <li>
                        "Uprisings" print                                                                             
                    </li>
                    <li>
                        <a href="http://www.ste-michelle.com//" target="_blank">Winery tour & Tasting for 4</a>
                    </li>
                    <li>
                        Bunny Baking Basket                                                                                  
                    </li>
                    <li>
                        Large Stuffed sitting Bunny holding a baby bunny                                                                                 
                    </li>
                    <li>
                        4 pairs of bunny earrings                                                                                  
                    </li>
                    <li>
                        Bunny Fleece Gloves                                                                                  
                    </li>
                    <li>
                        Bunny Magnetic board, bunny magnets, bunny mouse pad in Basket                                                                                  
                    </li>
                    <li>
                        Wine Hostess Gift (3)                                                                                  
                    </li>
                    <li>
                        Two bottles of St Michelle Wines, 1 red, 1 white Gift Wrapped                                                                                  
                    </li>
                    <li>
                        Numerous different Rabbit prints                                                                                  
                    </li>
                    <li>
                        Earthenware Steamer                                                                                  
                    </li>
                    <li>
                        Heart Bundt Baker                                                                                 
                    </li>
                    <li>
                        Framed Prints by Spokane artist Kit Jagoda                                                                              
                    </li>
                    <li>
                        <a href="http://www.craftsandframes.com/" target="_blank">$50
                        Gift Certificate from Ben Franklin</a>
                    </li>
                    <li>
                        <a href="http://www.portagebaygoods.com/" target="_blank">Portage Bay Goods Gift Basket/</a> 
                    </li>
                    <li>
                        <a href="http://www.rmcf.com/" target="_blank">Rocky Mountain
                        Chocolate Gift Basket</a>
                    </li>
                    <li>
                        <a href="http://www.undergroundtour.com/" target="_blank">Admission
                        for 2 to the Seattle Underground Tour</a>
                    </li>
                    <li>
                        <a href="http://www.spinalley.com/" target="_blank">Spin Alley
                        Bowling Gift Certificate</a>
                    </li>
                    <li>
                        <a href="http://www.empsfm.org/" target="_blank">Four admissions
                        to Experience Music Project/Science Fiction Museum</a>
                    </li>
                    <li>
                        <a href="http://www.premiergc.com/" target="_blank">Interbay-Round of Mini-Golf for 8 players</a>
                    </li>
                    <li>
                        <a href="http://www.zoo.org/" target="_blank">2 Adult & 2 Children - Admissions to
                        Woodland Park Zoo</a>
                    </li>
                    <li>
                        <a href="http://www.LynnwoodIceCenter.com/" target="_blank">Lynnwood Ice Center, 10 free admissions and skate rental</a>
                    </li>
                    <li>
                        <a href="http://www.essentialBaking.com/" target="_blank">Essential Bakery, 1 loaf of bread a month for 13 months</a>
                    </li>
                    <li>
                        <a href="http://www.todobienwellness.com/" target="_blank">1 hour massage treatment: Todo Bien! Wellness Center</a>
                    </li>
                    <li>
                        <a href="http://www.meanylodge.org/" target="_blank">Snowshoe Trek & Lunch at Meany Lodge</a>
                    </li>

                </ul>
            </div>
            </td>       
            <td valign="top" align="left">
                <!--#include file="sidebar_right.asp"-->
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <!--#include file="footer.asp"-->        
            </td>
        </tr>
    </table>
    </center>
    </body>
</html>
