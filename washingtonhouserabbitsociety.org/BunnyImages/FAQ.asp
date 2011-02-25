<html>

<head>
<title>HouseRabbit.org FAQ</title>
<meta NAME="Keywords" CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, non-profit, rabbit information, rescue, rabbit rescue, rodent rescue,
         pet rats, pet mice, pet hamsters, pet gerbils, pet ferrets, pet guinea pigs">

<meta NAME="Description" CONTENT="WashingtonHouseRabbitSociety.org is a the definitive site for Rescued Rabbits, Rodents and Ferrets in the 
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

</head>

<body>
<div align="center"><center>

<table border="0" width="90%">
  <tr>
    <td width="25%" rowspan="4" valign="top" align="left"><!--webbot bot="Include" U-Include="HRSMenu.htm" TAG="BODY" startspan -->

<!--#include file="hrsmenu1.htm"-->
<!--webbot bot="Include" endspan i-checksum="14527" -->

</td>
    <td width="10%" rowspan="4"></td>
    <td width="76%" valign="top" align="center"><font face="comic sans ms, arial" color="#400040" size="5"><b>House Rabbit Society<br>
    F.A.Q.</b></font></td>
  </tr>
  <tr>
    <td width="76%"><font face="comic sans ms, arial" size="2" color="#400040">The Frequently
    Asked Question or FAQ section will help you find answers to the most common questions. If
    you need an answer to something that is not here, please send mail to <a href="mailto:info@RabbitRodentFerret.org">info@RabbitRodentFerret.org</a> and let us know. We will be
    glad to look into your questions. </font></td>
  </tr>
  <tr>
    <td width="76%"><font face="comic sans ms, arial" size="2" color="#400040"><%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
%>

<!--#include file="connstr.asp" -->

<%
Conn.Open sConnect
'--------------------------------------------------------------------
' Display a record
'--------------------------------------------------------------------
If request("id")>0 then
	sql="Select * from FAQ where id=" & request("id")
	on error resume next
	set rsFAQ=conn.execute(sql)
	on error goto 0
	If rsFAQ.eof then
		response.write "No FAQ with that number found"
	else
		response.write "<br><b>"
		response.write rsFAQ("title") & "</b><br><br>"
		response.write rsFAQ("Text")
	end if
else
	sql="Select Id,Title from FAQ order by Title"
	set rsFAQ=conn.execute(sql)
	do while not rsFAQ.eof
%> <img src="images/btn_rabbit.gif" alt="[Rabbit Button]" WIDTH="16" HEIGHT="16"> <%
		response.write "<a href=""faq.asp?id=" & rsFAQ("id") & chr(34) & ">" & rsFAQ("Title") & "</a><br><br>"
		rsfaq.movenext
	loop
end if
'--------------------------------------------------------------------
' Close everything
'--------------------------------------------------------------------
set rsFaq=nothing
Conn.close
set Conn=nothing

%> </font></td>
  </tr>
  <tr>
    <td width="76%"></td>
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