<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>HouseRabbit.org - HouseBun</title>
<meta NAME="Keywords"
CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, non-profit, rabbit information, rescue, rabbit rescue, rodent rescue,
         pet rats, pet mice, pet hamsters, pet gerbils, pet ferrets, pet guinea pigs">

<meta NAME="Description"
CONTENT="WashingtonHouseRabbitSociety.org is a the definitive site for Rescued Rabbits, Rodents and Ferrets in the 
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
    <td width="25%" rowspan="3" valign="top"><!--webbot bot="Include" startspan
    U-Include="HRS/HRSMenu.htm" TAG="BODY" -->
<!--#include file="hrsmenu1.htm"-->

<!--webbot bot="Include" i-checksum="17071"   endspan -->
</td>
    <td width="10%" rowspan="3"></td>
    <td width="50%" valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>Welcome
    to HouseBun</b></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040">Welcome to
    HouseBun, the list for people who live with house rabbits! We hope this list will provide
    you with reliable information so that your rabbit will be given the opportunity to live to
    his or her full life potential of 8-10 years or more. This list is moderated by volunteers
    who have lived with hundreds of rabbits. Our purpose is to provide you with reliable
    accurate information, gathered from HRS foster homes around the US. We&#146;ve found that
    each rabbit is a unique individual and because of the large number of foster rabbits that
    we have known and loved, our perspective is much broader than people who live with one or
    two rabbits. We will attempt to obtain answers to medical related question from our very
    experienced rabbit veterinarians and therefore may not be able to respond to questions
    immediately. Please feel free to agreeably disagree or to question responses and of course
    to provide your own input to discussions. How else will we all continue to learn?&nbsp; </font><p><font
    face="comic sans ms, arial" size="2" color="#400040">To read previous archived discussions
    go to: </font><font face="Comic Sans MS" size="2"><a
    href="http://groups.yahoo.com/group/HouseBun/">http://www.groups.yahoo.com/group/HouseBun</a></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="3" color="#400040"><strong>To
    Subscribe to HouseBun</strong></font><form method="GET"
    action="http://groups.yahoo.com/subscribe/HouseBun">
      <table cellspacing="0" cellpadding="2" border="0" bgcolor="#FFFFCC">
        <tr>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td colspan="2" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Subscribe to
          HouseBun</b> &nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="user"
          value="enter email address" size="20"> &nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="image" border="0"
          alt="Click here to join HouseBun" name="Click here to join HouseBun"
          src="http://groups.yahoo.com/img/ui/join.gif"> &nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;&nbsp;</td>
        </tr>
        <tr align="center">
          <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Powered by
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://groups.yahoo.com/">groups.yahoo.com</a>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;&nbsp;</td>
        </tr>
      </table>
    </form>
    <p><font face="comic sans ms, arial" size="3" color="#400040"><strong>List moderators:</strong></font>
    <font face="comic sans ms, arial" size="3" color="#400040"></p>
    <p align="left"></font><font face="comic sans ms, arial" size="2" color="#400040">There
    are currently 9 anonymous moderators running the HouseBun list</font><font
    face="comic sans ms, arial" size="3" color="#400040"></p>
    <p align="left"></font><font face="comic sans ms, arial" size="2" color="#400040"><strong>List
    Owner</strong>: <a href="mailto:sandi@rabbitrodentferret.org">Sandi@RabbitRodentFerret.org</a><br>
    House Rabbit Society Seattle, WA</font></td>
  </tr>
</table>
</center></div>

<hr width="90%">

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>