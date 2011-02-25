<% @LANGUAGE=VBScript %>
<% Option Explicit %>

<HTML><HEAD><TITLE>Best Little Rabbit, Rodent and Ferret House - Quality Supplies for Companion Animals</TITLE><meta name="description" content="Best Little Rabbit, Rodent and Ferret House Pet Supplies and Adoption">

<meta name="keywords" content="rabbits, rodents, ferrets, pets, companion animals, shelter, pet supplies, pet food,
   guinea pig, rats, mice, chinchilla, prairie dog, adoption, washington, seattle, puget"> 

<meta name="copyright" content="Copyright 2000 Best Little Rabbit, Rodent and Ferret House. All rights reserved. 
        Contact author for reprint policies.">

	<!--#include file="dropdownmenu.asp"-->

<style type="text/css">
		span.purple{
font-weight: bold; 
color:#669; 
font-style: normal; 
font-size: 200%; 
line-height: 1.2; 
font-family:  Arial,Helvetica;
} 

span.orange{
font-weight: bold; 
color:#F93; font-style: 
normal; font-size: 200%; 
line-height: 1.2; 
font-family: Arial,Helvetica;
} 
td.philosophy3 {				
  background-color:#EEC;
  border:1px dotted #993;
  padding:25px;
}
p.philos {
font-weight: normal; 
font-style: normal; font-size: 16px; 
line-height: 1.6; 
font-family: Arial,Helvetica;} 

span.green{font-weight: bold; color:#993; font-style: normal; font-size: 200%; line-height: 1.2; font-family: Arial,Helvetica;} 

.linktext {
font-family:Verdana,Arial,Helvetica; 
color:#000; 
font-size:10px; 
text-decoration: none
}
</style>
</head>
<bodY>
<!--main table-->

<TABLE WIDTH=95% HEIGHT=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 valign="top">


<tr><td valign="top">
	<!--#include file="headerfile.asp"-->

<%
Dim referer,page
referer = Trim(Request.ServerVariables("HTTP_REFERER"))
if referer="" then
referer="index.asp"
end if
page=request.querystring("name")
if page="" then page = "animals"

%>

</td></tr>
<tr><td valign="top" align="center"><font face="arial" color="#FF9933"><H1>Please Help the <%=Ucase(left(page,1))%><%=Right(page,(Len(page)-1))%>!</font>
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
</td></tr>
<tr><td valign="top" align="center"><center><table border="1" cellpadding="25" ><tr><td class="philosophy3"  valign="top">
 <p align="center" class="philos"><b><span class="orange">Y</span>OU CAN MAKE A DIFFERENCE!</b>

<p class="philos" align="left"><span class="green">W</span>e are taking care of more animals than we ever have, and with less money.
 We are in the process of moving our shelter to Rabbit Meadows on Union Hill in Redmond. We need lumber for enclosures; nails and hardware for enclosure doors; since we're on a private well, we need to install a holding tank (depending upon the size it will be from $3,000 - $4,000.) We need to put up outside lights, do some grading for parking spaces, etc.  
 
 <p class="philos" align="left"><span class="purple">O</span>ur future goal is to build a permanent shelter for our small friends. This move is the begining of making that dream a reality. 

<p class="philos" align="left"><span class="green">W</span>e remain committed to providing proper care for our adoptable and permanent residents, and to rescuing as many unwanted, homeless and abandoned animals as we can afford to help.  All money donated goes directly to their care. Any amount will help and is very much appreciated.

<p class="philos" align="left">
We are a non-profit 501(c)(3) organization. Your donations are tax deductible 91-1873550
</p>

<p class="philos" align="center"><i>This page will soon list our currently needed items.<br>  In the meantime, if you wish to donate items, contact the shelter <a href="contactus.asp" >here</a></i>

<p class="philos" align="center"><b>THANK YOU!</b>


</td></tr></table></center>

</table>

</body>



<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>