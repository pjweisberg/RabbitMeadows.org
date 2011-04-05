<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>HouseRabbit.org - Membership</title>
<meta NAME="Keywords" CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, non-profit, rabbit information, rescue, rabbit rescue, rodent rescue,
         pet rats, pet mice, pet hamsters, pet gerbils, pet ferrets, pet guinea pigs">

<meta NAME="Description" CONTENT="
		WashingtonHouseRabbitSociety.org is a the definitive site for Rescued Rabbits, Rodents and Ferrets in the 
Northwest and other parts of the country. HRS is a non-profit organization.">
<meta NAME="Robots" CONTENT="index all">
<script>
	function lite(whichObj,name){
		whichObj.src = "./images/" + name + "_hi.gif";
	}

	function deLite(){
		var tmpObj;
		for(var i = 0; i < 5; i++){
			tmpObj = eval('document.btn' + i);	
			if (tmpObj.src.indexOf("_hi") > -1) {
				tmpObj.src = tmpObj.src.substring(0,tmpObj.src.indexOf("_hi")) + ".gif";
			}
		}
	}
   </script>

<style>
	a{color:#396bb5; text-decoration:none; font-weight:bold}
	a:visited{color:#555555}
	a:hover{text-decoration:underline; }
   </style>
<!--#include file="google-analytics.js"-->
</head>

<body>
<div align="center"><center>

<table border="0" width="90%">
  <tr>
    <td width="25%" rowspan="3" valign="top"><!--webbot bot="Include" U-Include="HRSMenu.htm" TAG="BODY" startspan -->

<!--#Include file="hrsmenu1.htm">
<!--webbot bot="Include" endspan i-checksum="14527" -->

</td>
    <td width="10%" rowspan="3"></td>
    <td width="50%" valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>Joining
    The House Rabbit Society</b></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040">Time is all it
    takes for a rabbit to be discovered by the right human. When their time is up at the
    animal shelters, rabbits, with your support, can be placed in foster homes until adoptive
    &quot;matches&quot; are made.</font><p><font face="comic sans ms, arial" size="2" color="#400040">Your donations and enrollment in the House Rabbit Society help provide
    needy rabbits with food, housing, veterinary care and enough time to find permanent homes.</font></p>
    <p><font face="comic sans ms, arial" size="2" color="#400040">Our yearly membership
    includes the locally published <em>Washington House Rabbit News,</em>. &nbsp; Additional donations are always
    welcome.</font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="3" color="#400040"><strong>House
    Rabbit Society Membership Enrollment:</strong></font> <p><strong><font face="comic sans ms, arial" size="2" color="#400040">United States: $15/yr. for Washington State membership.</font></strong></p>
    <p><font face="comic sans ms, arial" size="2" color="#400040">Note - All donations to
    House Rabbit Society are <strong><em>tax-deductible</em></strong>.</font></p>
    <p><font face="comic sans ms, arial" size="2" color="#400040">Please send check to: 
    <br>
    <strong>House Rabbit Society - Washington State</strong><br>
    P.O. Box 3242, Redmond, WA 98073. Please state if this is a donation or a membership
    request/renewal.</font></td>
  </tr>
</table>
</center></div>

<hr width="90%">
<!--webbot bot="Include" U-Include="_private/footer.htm" TAG="BODY" startspan -->
<div align="center"><center>

<table border="0" width="90%">
  <tr>
    <td width="100%" align="center"><font face="comic sans ms, arial" size="2" color="#400040">Copyright
      1999 House Rabbit Society - Washington</td>
  </tr>
</table>
</center></div>
<!--webbot bot="Include" endspan i-checksum="29637" -->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>