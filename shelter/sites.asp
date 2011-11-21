<% @LANGUAGE=VBScript %>
<% Option Explicit %>
<!-- #include file="correct-domain.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Best Little Rabbit, Rodent and Ferret House - Rabbit Links</title>
<!--#include file="dropdownmenu.asp"-->
<meta NAME="Keywords" CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, non-profit, rabbit information, rescue, rabbit rescue, rodent rescue,
         pet rats, pet mice, pet hamsters, pet gerbils, pet ferrets, pet guinea pigs">

<meta NAME="Description" CONTENT="Best Little Rabbit, Rodent and Ferret House is a the definitive site for Recued Rabbits in the 
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
<tr><td colspan="3">
<!--#include file="headerfile.asp"-->
</td></tr>
  <tr>
    
    <td width="10%" rowspan="3"></td>
    <td width="50%" valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>House-Rabbit
    Friendly Links</b></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040"><%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------
Dim Conn
Set Conn = Server.CreateObject("ADODB.Connection")
%>

<!--#include file="connstr.asp" -->

<%
Conn.Open sConnect
'--------------------------------------------------------------------
' Display All the Records
'--------------------------------------------------------------------
Dim sql
Dim rsLinks
sql="Select * from links order by [Desc]"
set rsLinks=conn.execute(sql)
do while not rsLinks.eof
%><img src="images/btn_rabbit.gif" alt="[Rabbit Button]" WIDTH="16" HEIGHT="16"> <%
	response.write "<a href=" & chr(34) & rsLinks("Link") & chr(34) & ">" & rsLinks("Desc") & "</a><br><br>"
	rsLinks.movenext
loop

'--------------------------------------------------------------------
' Close everything
'--------------------------------------------------------------------
set rsLinks=nothing
Conn.close
set Conn=nothing

%> </font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040">If You want to
    see a link here send us email at <a href="mailto:info@rabbitmeadows.org.org">info@rabbitmeadows.org</a>
    </font></td>
  </tr>
</table>
</center></div>

<hr width="90%">
<!--webbot bot="Include" U-Include="_private/footer.htm" TAG="BODY" startspan -->
<div align="center"><center>

<table border="0" width="90%">
  <tr>
    <td width="100%" align="center"><!--#include file="footer.asp"--> </td>
  </tr>
</table>
</center></div>
<!--webbot bot="Include" endspan i-checksum="29637" -->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>
