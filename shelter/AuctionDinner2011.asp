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
                    <h3><a href="http://www.rustypelicancafe.com/">Rusty Pelican Café</a><br />
                    1924 N 45th St., Seattle</h3>
                </center>
                <img src="BunnyImages/Josie1.jpg" alt="Josie, our 2011 Auction Mascot" style="float:right" />
                <p>
                    For a $35 donation, you get a three-course meal including your choice of Vegetarian Lasagna or Greek Pasta,
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
                    <a href="2011%20Auction%20Letter%20&%20Procurement%20Form.pdf">this letter</a>
                </p>
                <center>
                    <h1>Auction Items List</h1>
                    More to be added as the date approaches
                </center>
                <ul>
                    <li>
                       <a href="http://www.ridetheducksofseattle.com/" target="_blank">Two
                        adult admissions to Ride the Ducks</a>
                    </li>
                    <li>
                        <a href="http://www.ahouseofclocks.com/" target="_blank">Seiko
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
                        <a href="http://larkin-art.deviantart.com/" target="_blank">Multi
                        Media Art Picture of Lop Rabbit</a>
                    </li>
                    <li>
                        <a href="http://www.dermagenics.com/" target="_blank">Skin
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
                        <a href="http://www.dermagenics.com/" target="_blank">Dermagenics Collagen Recovery Cream</a>
                    </li>
                    <li>
                        <a href="http://www.jetcityimprov.com/" target="_blank">Admission
                        for 4 to Seattle&#39;s best improvisational comedy show!</a>
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
                        <a href="http://www.Generation%20Thrive.com" target="_blank">Thrive
                        Gift Card</a>
                    </li>
                    <li>
                        <a href="http://www.custompure.com/" target="_blank">Gift
                        Certificate to Custom Pure</a>
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
                        <a href="http://www.museumofflight.org/" target="_blank">4
                        Tickets to Museum of Flight</a>
                    </li>
                    <li>
                        Silver Rabbit Poop Earrings with Bright Swirls                 
                    </li>
                    <li>
                        Bunny Blossom print                                                                
                    </li>
                    <li>
                        Uprisings print                                                                             
                    </li>
                    <li>
                        Rabbit print                                                                                  
                    </li>
                    <li>
                        Framed Prints                                                                              
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
                        <a href="http://www.zoo.org/" target="_blank">Admission to
                        Woodland Park Zoo</a>
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
