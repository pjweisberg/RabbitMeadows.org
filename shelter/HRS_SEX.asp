<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>HouseRabbit.org - Rabbit Sex</title>
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

<table border="0" width="100%">
  <tr>
    <td width="100%" align="center"><a href="default.htm"><img src="images/HRSLOGO.gif" alt="Back to HomePage" border="0" width="108" height="89"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="news.asp"><img src="images/btn_News.gif" alt="btn_News.gif (1228 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="AdoptedCurrent.asp"><img src="images/btn_Adoptions.gif" alt="btn_Adoptions.gif (1087 bytes)" border="0" width="144" height="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="Membership.asp"><img src="images/btn_Membership.gif" alt="btn_Membership.gif (1126 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="Supplies.asp"><img src="images/btn_Supplies.gif" alt="btn_Supplies.gif (1013 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="FAQ.asp"><img src="images/btn_FAQ.gif" alt="btn_FAQ.gif (790 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="vets.asp"><img src="images/btn_RabbitVets.gif" alt="btn_RabbitVets.gif (1189 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="sites.asp"><img src="images/btn_RabbitLinks.gif" alt="btn_RabbitLinks.gif (1084 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="HouseBun.asp"><img src="images/btn_HouseBun.gif" alt="btn_HouseBun.gif (1288 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="AboutHRS.asp"><img src="images/btn_AboutUs.gif" alt="btn_AboutUs.gif (1007 bytes)" border="0" WIDTH="144" HEIGHT="19"></a></td>
  </tr>
</table>
<!--webbot bot="Include" endspan i-checksum="17023" -->
</td>
    <td width="10%" rowspan="3"></td>
    <td width="50%" valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>Checking
    the Sex</b></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040">When a rabbit
    is very young, it can be hard to tell if it is a male or a female.&nbsp; Ask your vet to
    check the sex of your rabbit, so you are sure it&nbsp; is&nbsp; the sex you wanted.</font></td>
  </tr>
  <tr>
    <td width="50%"><p align="center"><img src="images/HRS_SEX_Male.jpg" alt="HRS_SEX_Male.jpg (10596 bytes)" WIDTH="137" HEIGHT="116"></p>
    <p align="center"><font face="comic sans ms, arial" size="3" color="#400040"><strong>Male
    Rabbit</strong></font></p>
    <p align="center"><img src="images/HRS_SEX_FEMALE.jpg" alt="HRS_SEX_FEMALE.jpg (10584 bytes)" WIDTH="137" HEIGHT="115"></p>
    <p align="center"><font face="comic sans ms, arial" size="3" color="#400040"><strong>Female
    Rabbit</strong></font></td>
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