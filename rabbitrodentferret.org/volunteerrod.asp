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
	
<center><table width="700" cellspacing="0" cellpadding="0" border="0">
<tr><td rowspan="3" valign="top"><img src="header2a.gif" width="150" height="95" border="0" alt=""></td>
<td valign="top">
<img src="header2b.gif" width="550" height="56"" border="0" alt=""></td></tr>
<tr><td valign="top">
<img src="header2c.gif" width="550" height="8" border="0" alt=""></td></tr>
<tr><td valign="top">
	<Table border="0"><tr><td><img src="spacer.gif" height="25" width="0"></td>

	<td valign="top"><img  src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuHome, '45px')" onMouseout="delayhidemenu()">Home</a> </td>

	<td valign="top"><img  src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuAdopt, '110px')" onMouseout="delayhidemenu()">Adoption</a> </td>

	<td valign="top"><img  src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuStore, '110px')" onMouseout="delayhidemenu()">Store</a> </td>

	<td valign="top"><img   src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuNews, '100px')" onMouseout="delayhidemenu()">News</a> </td>

	<td valign="top"><img   src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuVets, '95px')" onMouseout="delayhidemenu()">Vets</a> </td>

	<td valign="top"><img  src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuFaq, '95px')" onMouseout="delayhidemenu()">FAQs</a> </td>

	<td valign="top"><img   src="links.gif" width="8" height="10"><a class="linktext" href="default.htm" onClick="return clickreturnvalue()" onMouseover="dropdownmenu(this, event, menuCon, '100px')" onMouseout="delayhidemenu()">Contact Us</a> </td></tr></table>
	</td></tr>
<!--end of page header section-->
</table></center>
<%
Dim referer,page
referer = Trim(Request.ServerVariables("HTTP_REFERER"))
if referer="" then
referer="index.asp"
end if
page=request.querystring("name")

%>

</td></tr>
<tr><td valign="top" align="center"><font face="arial" color="#FF9933"><H1>Volunteer to Help the Rodents!</font></td></tr>
<tr><td valign="top"><center><table border="1" cellpadding="25" ><tr><td class="philosophy3"  valign="top">
 <p class="philos"><span class="orange">W</span>e would love to have you as a volunteer!
 <p class="philos"><span class="green">T</span>his page will soon have more content and a list of positions. 
<p class="philos"><span class="purple">I</span>n the meantime, please contact the shelter <a href="contactus.asp">here</a>.


<p align="center"><Font face="arial" color="#F93" size="5">
<a href="<%=referer%>">BACK</a></font>
<%=page%>
</td></tr></table></center>

</table>

</body>



<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>