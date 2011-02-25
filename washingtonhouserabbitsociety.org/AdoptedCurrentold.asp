<html>

<head>
<title>HouseRabbit.org - Bunnies</title>
<meta NAME="Keywords"
CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, non-profit, rabbit information, rescue, rabbit rescue, rodent rescue,
         pet rats, pet mice, pet hamsters, pet gerbils, pet ferrets, pet guinea pigs">

<meta NAME="Description"
CONTENT="WashingtonHouseRabbitSociety.org is a the definitive site for Rescued Rabbits, Rodents, and Ferrets in the 
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
<table border=1 bordercolor=black bgcolor=darkblue><tr><td align=center>
<font color="white" face="comic sans ms" size=3>Please help the animals by donating towards spay/neuter and other veterinary services  
<font size=3><b>www.WashingtonHouseRabbitSociety.org </b></font>
Thanks!</font></td></tr></table><br><br>

<table border="0" width="90%">
  <tr>
    <td width="25%" rowspan="3" valign="top" align="left"><!--webbot bot="Include"
    U-Include="HRSMenu.htm" TAG="BODY" startspan -->

<!--#include file="hrsmenu1.htm">
<!--webbot bot="Include" endspan i-checksum="14527" -->
</td>
    <td width="5%" rowspan="3"></td>
    <td valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>House Rabbits 
looking for a good home...</b></font></td> 
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" color="#400040" size="2"><br>
    </font>Please<font face="comic sans ms, arial" color="#400040" size="2"> visit our
    adoption center at: <br>
    <strong>The Best Little Rabbit, Rodent &amp; Ferret House </strong><br>
    14325 Lake city Way NE, Seattle WA 98125 (phone) 206-365-9105 </font><br>
    <font face="comic sans ms, arial" color="#400040" size="2">We only do adoptions within the
    Seattle and Puget Sound area. Companions rabbits must live inside in order to be part of 
the family. If interested in a particular animal, you must come in to
    the shelter to meet one another.

</font><hr>



<%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection")



%>

<!--#include file="connstr.asp" -->

<%


Conn.Open sConnect
'--------------------------------------------------------------------
' Display All the Records
'--------------------------------------------------------------------
sql="Select * from rebeccad.Adopt where archive=0"
set rs=conn.execute(sql)
do while not rs.eof
response.write "<b>"
	response.write rs("Name") & "</b> - Availability: "

	If rs("adopted")=true then
		response.write "<b>Adopted</b><br>"
	else
		response.write "<font color=RED><b>Adopt Me!</font></b><br>"
	end if

	response.write "<br>"
	response.write rs("Desc") & "<br>"
	If not (isnull(rs("Picture1")) or rs("Picture1")="") then
		response.write "<img hspace=4 vspace=4 src=" & chr(34) & "BunnyImages/" & rs("Picture1") & chr(34) & ">"
	end if
	If not (isnull(rs("Picture2")) or rs("Picture2")="") then
		response.write "<br clear=left><img hspace=4 vspace=4 src=" & chr(34) & "BunnyImages/" & rs("Picture2") & chr(34) & ">"
	end if
	If not (isnull(rs("Picture3")) or rs("Picture3")="") then
		response.write "<br clear=left><img hspace=4 vspace=4 src=" & chr(34) & "BunnyImages/" & rs("Picture3") & chr(34) & ">"
	end if
	If not (isnull(rs("Picture4")) or rs("Picture4")="") then
		response.write "<br clear=left><img hspace=4 vspace=4 src=" & chr(34) & "BunnyImages/" & rs("Picture4") & chr(34) & ">"
	end if

	response.write "<br><hr>"


	rs.movenext
loop

'--------------------------------------------------------------------
' Close everything
'--------------------------------------------------------------------
set rs=nothing
Conn.close
set Conn=nothing

%>
    </td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td></td>
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