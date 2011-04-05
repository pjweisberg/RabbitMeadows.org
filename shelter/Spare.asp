<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>HouseRabbit.org - Spare</title>
<meta NAME="Keywords" CONTENT="live online House Rabbit, Rabbitmedia,
	online Rabbits, streaming live videos,Adoptions, Rabbit Facts, Rabbit Sex, Rabbit news,
	online media ">
<meta NAME="Description" CONTENT="WashingtonHouseRabbitSociety.org is a the definitive site for Recued Rabbits in the 
Northwest and other parts of the country. HRS is a non-provit organization.">
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

<table cellspacing="0" cellpadding="0" border="0">
<tr>
	<td nowrap width="1" height="1"><img src="images/NEwspace.gif" width="1" height="1"></td>
	<td nowrap width="58" height="1"><img src="images/NEwspace.gif" width="1" height="1"></td>
	<td nowrap width="90" height="1"><img src="images/NEwspace.gif" width="1" height="1"></td>
	<td nowrap width="44" height="1"><img src="images/NEwspace.gif" width="1" height="1"></td>
</tr>
<tr>
	<td width="1" height="30"><img src="images/NEwspace.gif" width="1" height="1"></td>
	<td valign="top" rowspan="3"><img src="images/PhotoDraw11.gif" border="0" width="58" height="106"></td>
	<td valign="top"><img src="images/PhotoDraw12.gif" border="0" width="90" height="30"></td>
	<td valign="top" rowspan="3"><img src="images/PhotoDraw13.gif" border="0" width="44" height="106"></td>
</tr>
<tr>
	<td width="1" height="44"><img src="images/NEwspace.gif" width="1" height="1"></td>
	<td valign="top"><img src="images/PhotoDraw14.jpg" border="0" width="90" height="44"></td>
</tr>
<tr>
	<td width="1" height="32"><img src="images/NEwspace.gif" width="1" height="1"></td>
	<td valign="top"><img src="images/PhotoDraw15.gif" border="0" width="90" height="32"></td>
</tr>
</table>

<table border="0" width="100%">
  <tr>
    <td width="100%" align="center"></td>
  </tr>
  <tr>
    <td width="100%"><a href="news.asp"><b>What's News</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="AdoptedCurrent.asp"><b>AdoptMe</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="Membership.asp"><b>Membership</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="Supplies.asp"><b>Supplies</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="FAQ.asp"><b>Questions?</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="vets.asp"><b>Vet Referrals</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="sites.asp"><b>Other Links</b></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="HouseBun.asp"><b>HouseBun Online</b></a></td>
  </tr>
  <tr>
    <td width="100%"><b><a href="rodents/Default.htm">Rodents...</a></b></td>
  </tr>
  <tr>
    <td width="100%"><b><a href="ferrets/default.htm">Ferrets...</a></b></td>
  </tr>
  <tr>
    <td width="100%"><a href="AboutHRS.asp"><b>About Us</b></a></td>
  </tr>
</table>
<!--webbot bot="Include" endspan i-checksum="10270" -->

</td>
    <td width="10%" rowspan="3"></td>
    <td width="50%" valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>Title
    Goes Here...</b></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040">Text Goes
    Here...</font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="3" color="#400040"><strong>More
    Text Goes Here...</strong></font></td>
  </tr>
</table>
</center></div>

<hr width="90%">
<!--webbot bot="Include" U-Include="_private/footer.htm" TAG="BODY" startspan -->
<div align="center"><center>

<table border="0" width="90%">
  <tr>
    <td width="100%" align="center"><font face="comic sans ms, arial" size="2" color="#400040">Copyright
      1999 House Rabbit Society - Washington<br>
    </font><b><a href="http://www.connectos.com/" target="_blank"><font size="2" face="comic sans ms, arial">Hosting by
    ConnectOS Corporation</font></a></b> </td>
  </tr>
</table>
</center></div>
<!--webbot bot="Include" endspan i-checksum="29637" -->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>