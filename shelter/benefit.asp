<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->


<HTML><HEAD><TITLE>Best Little Rabbit, Rodent and Ferret House - Quality Supplies for Companion Animals</TITLE><meta name="description" content="Best Little Rabbit, Rodent and Ferret House Pet Supplies and Adoption">

<meta name="keywords" content="rabbits, rodents, ferrets, pets, companion animals, shelter, pet supplies, pet food,
   guinea pig, rats, mice, chinchilla, prairie dog, adoption, washington, seattle, puget"> 

<meta name="copyright" content="Copyright 2000-2011 Best Little Rabbit, Rodent and Ferret House. All rights reserved. 
        Contact author for reprint policies.">

	<!--#include file="dropdownmenu.asp"-->

<LINK REL=StyleSheet HREF="style.css" TYPE="text/css">
<style type="text.css">


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
<!--#include file="google-analytics.js"-->
</head>
<bodY>
<!--main table-->

<center><TABLE WIDTH="550" HEIGHT=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 valign="top" >


<tr><td   valign="top">
	
<!--#include file="headerfile.asp"-->
<%
Dim referer,page
referer = Trim(Request.ServerVariables("HTTP_REFERER"))
if referer="" then
referer="index.asp"
end if
page=request.querystring("name")

%>
</td></tr>
<tr ><td valign="top" background="benefit.jpg"  STYLE="background-repeat: no-repeat;" >
<table>

<tr><td valign="top" align="center">
<font face="arial" size="2">
<br><br>
<b>Join us for a <br> Dinner and Silent Auction hosted by</b><br>
<h2><b>The Rusty Pelican Cafe</B></h2> 

to benefit the lifesaving work of<p>
<b>The Best Little Rabbit, Rodent & Ferret House</b> <p>
<b>Saturday, April 21st, 2007 </b><br>
<b>2:30 - 5:30 p.m.</b><p><p align="Left">
at the Rusty Pelican Cafe<br>
1924 N 45th Street, Seattle<p align="left">
Choice of two vegetarian entrees<br>
$35.00 per person<br>
100% of proceeds benefiting BLRRFH<P align="Left">
Make your reservation today!<br>
206.365.9105<br>
BLRRFHauction@gmail.com<br><br>

or pay by PayPal <br>(this will appear as a donation -just type in your information)

<!--#include file="paypal_logo.html"-->


<p align="Left">

 <a href="BLRRFH Auction Poster_revised.pdf"><b>PDF flyer</b></a>
<br>
More info on the <a href="BLRRFH Auction Procurement Letter.pdf"><b>auction</b></a>.<br>
Auction item donation <a href="BLRRFH 2007 Auction Procurement Form.pdf "><b>form</b></a>.
<hr>

</font>
</td></tr></table>
</td>

</tr>

</table></center>
<!--#include file="footer.asp"-->
</body>



<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>
