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
</td></tr>
<%
Dim referer,page
referer = Trim(Request.ServerVariables("HTTP_REFERER"))
if referer="" then
referer="index.asp"
end if
page=request.querystring("name")


Dim blurb
Select Case page

Case "Sanctuary Rabbits":

blurb = "<p class=""philos""><span class=""purple"">R</span><b>abbit Meadows Sanctuary</b> is located in Redmond, WA, near Novelty Hill/Union Hill Road.  <p class=""philos""><span class=""purple"">A</span>ll of the rabbits at <B>Rabbit Meadows Sanctuary</B> are feral rabbits rescued from a variety of locations in the greater Seattle area. They have all been spayed and neutered and live naturally in family groups in a safe and protected environment.  <p class=""philos""><span class=""purple"">R</span><B>abbit Meadows Sanctuary</B> is looking for volunteers who love animals, are dependable and dont mind getting a little dirty. <p class=""philos""><span class=""purple"">G</span>ood hygiene is essential to the health of the animals. A volunteer's work will primarily be cleaning up after them. This work is not glamorous, but it is critical to their health and well-being. <p class=""philos""><span class=""purple"">W</span>e ask for a commitment of one weekend day every 6-8 weeks.  The typical schedule is 2 adults every Sunday (or Saturday if you prefer). It takes 2 adults about 2 hours to clean the rabbit runs and put out fresh veggies. <p class=""philos""><span class=""purple"">I</span>f you or someone you know loves animals and would like to help put at <B>Rabbit Meadows Sanctuary</B>, please email <a href=""mailto:volunteer@rabbitrodentferret.org"">Volunteer@rabbitrodentferret.org</a>."

Case else:
blurb = "<p class=""philos""><span class=""orange"">W</span>e would love to have you as a volunteer! <p class=""philos""><span class=""green"">T</span>his page will soon have more content and a list of positions open. <p class=""philos""><span class=""purple"">I</span>n the meantime, please contact the shelter <a href=""contactus.asp"">here</a>."

End Select
%>



<tr><td valign="top" align="center"><font face="arial" color="#FF9933"><H1>Volunteer to Help the <%=Ucase(left(page,1))%><%=Right(page,(Len(page)-1))%></font></td></tr>
<tr><td valign="top"><center><table border="1" cellpadding="25" width="70%"><tr><td class="philosophy3"  valign="top">
 <%=blurb%>


<p align="center"><Font face="arial" color="#F93" size="5">
<a href="<%=referer%>">BACK</a></font>

</td></tr></table></center>

</table>
<!--#include file="footer.asp"-->
</body>



<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>