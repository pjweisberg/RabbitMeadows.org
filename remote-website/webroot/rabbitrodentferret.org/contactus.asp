<% @LANGUAGE=VBScript %>
<% Option Explicit %>

<HTML><HEAD><TITLE>Best Little Rabbit, Rodent and Ferret House - Quality Supplies for Companion Animals</TITLE><meta name="description" content="Best Little Rabbit, Rodent and Ferret House Pet Supplies and Adoption">

<meta name="keywords" content="rabbits, rodents, ferrets, pets, companion animals, shelter, pet supplies, pet food,
   guinea pig, rats, mice, chinchilla, prairie dog, adoption, washington, seattle, puget"> 

<meta name="copyright" content="Copyright 2000-2011 Best Little Rabbit, Rodent and Ferret House. All rights reserved. 
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
td.philosophy1 {				
  background-color:#FD9;
  border: 1px dotted #F93;
    padding:5px;
}

td.philosophy2 {				
  background-color:#dde;
  border:1px dotted #669;
  padding:5px;
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
.dklink {
font-family:Verdana,Arial,Helvetica; 
color:#000; 

font-weight: normal;
font-size:14px;
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

%>

</td></tr>
<tr><td valign="top" align="center">
<font face="arial" color="#FF9933"><H1>CONTACT US</font></td></tr>

<tr><td valign="top">

<center><table border="0" cellpadding="25" >
<tr>
<td class="philosophy3"  valign="top"><p class="philos">Phone: </td><td colspan= "2" class="philosophy3"><p class="philos">(206)  3 6 5 - 9 1 05</td></tr>

<tr><td class="philosophy3"><p class="philos">E-mail:</td><td class="philosophy3"><p class="philos">Questions about adoption or products:</td><td class="philosophy2"><p class="philos"><a href="mailto:store@rabbitrodentferret.org" class="dklink">Info@rabbitrodentferret.org</a></td></tr>

<tr><td class="philosophy3"><p class="philos">E-mail:</td><td class="philosophy3"><p class="philos">Volunteering or Donating Items:</td><td class="philosophy2"><p class="philos"><a href="mailto:store@rabbitrodentferret.org" class="dklink">Store@rabbitrodentferret.org</a></td></tr>

<tr><td class="philosophy3"><p class="philos">E-mail:</td><td class="philosophy3"><p class="philos">Questions or comments about website</td><td class="philosophy2"><p class="philos"><a href="mailto:rebecca@rabbitrodentferret.org" class="dklink">Rebecca@rabbitrodentferret.org</a></td></tr>
</td></tr></table></center>

</table>
<!--#include file="footer.asp"-->
</body>



<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>