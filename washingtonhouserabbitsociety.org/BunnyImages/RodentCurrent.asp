<html>

<head>
<title>HouseRabbit.org - Rodents</title>
<meta NAME="Keywords"
CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, rabbit information, rescue, rabbit rescue, rodent rescue,
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
</head>

<body>
<div align="center"><center>
<table border=1 bordercolor=black bgcolor=darkblue><tr><td align=center>
<font color="white" face="comic sans ms" size=3> Soon we will be discontinuing the www.houserabbit.org address.  You will only be able to find us at: 
<font size=3><b>www.WashingtonHouseRabbitSociety.org </b></font>
Please adjust your bookmarks.  Thanks!</font></td></tr></table><br><br>

<table border="0" width="90%">
  <tr>
    <td width="25%" rowspan="2" valign="top" align="left">

<!--#include file="hrsmenu1.htm"-->
<!--webbot bot="Include" endspan i-checksum="14527" -->
</td>
    <td width="10%" rowspan="4"></td>


    <td><b><font face="comic sans ms, arial" color="#400040" size="5">Rodent
    Rescue</font></b>
	<p><strong><font face="comic sans ms, arial" size="2" color="#400040">The
    Best Little Rabbit, Rodent &amp; Ferret House is a non-profit animal welfare organization
    dedicated to:</font></strong><ul>
      <li><font face="comic sans ms, arial" size="2" color="#400040"><strong>The rescue and then
        adoption of rodents who have lost their home, through no fault of their own. </strong></font></li>
      <li><strong><font face="comic sans ms, arial" size="2" color="#400040">The education of
        persons interested in sharing their homes with a rodent.</font></strong></li>
    </ul>
    </td>
  
  <tr>
    <td><font face="comic sans ms, arial" size="2" color="#400040">We rescue
    guinea pigs, rats, mice, gerbils,&nbsp; hamsters, prairie dogs, chinchillas and the
    occasional hedgehog. We work with local animal shelters and accept these animals from them
    as space permits. (Only <em>if</em> a shelter has none of these animals will we consider
    taking in your rodents.) </font><p><font face="comic sans ms, arial" size="2"
    color="#400040">We spay all guinea pigs,&nbsp; chinchillas &amp; prairie dogs&nbsp; and
    neuter these animals as funds permit. This is done to prevent hormonal tumors and early
    deaths, as well as a means to allow these animals to live humanely with their own species.</font><br></p>
</td></tr>

<tr>

    <td valign="top" colspan=3 align=center><hr><br><font face="comic sans ms, arial" color="#400040" size="5"><b>
    Rodents looking for a good home...</b></font></td>
  </tr>
  <tr>
    <td colspan=3><font face="comic sans ms, arial" color="#400040" size="2"><br>
    </font>Please<font face="comic sans ms, arial" color="#400040" size="2"> visit our
    adoption center at: <br>
    <strong>The Best Little Rabbit, Rodent &amp; Ferret House </strong><br>
    14325 Lake city Way NE, Seattle WA 98125 (phone) 206-365-9105 (fax)</font><br>
    <font face="comic sans ms, arial" color="#400040" size="2">We only do adoptions within the
    Seattle and Puget Sound area. If interested in a particular animal, you must come in to
    the shelter to meet one another.</font><p>
<font face="comic sans ms, arial" color=#0400040" size=1>
We include guinea pigs as rodents even though their 
classification has changed, otherwise we'd have to change the name of 
our shelter.    :-)

</font></p><hr>
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
sql="Select * from AdoptRodents where archive=false order by Name, Adopted DESC"
set rs=conn.execute(sql)
do while not rs.eof
response.write "<b>"
	response.write rs("Name") & "</b> - Availability: "

	If rs("adopted")=true then
		response.write "<b>Adopted</b><br>"
	else
		response.write "<font color=RED><b>Adopt Me!</font></b><br>"
	end if

	
	If not (isnull(rs("Picture1")) or rs("Picture1")="") then
		response.write "<p align=left><img align=left vspace=4 hspace=10 src=" & chr(34) & "RodentImages/" & rs("Picture1") & chr(34) & ">"
	end if
	

	
	response.write rs("Desc") & "</p><br clear=left>"

If not (isnull(rs("Picture2")) or rs("Picture2")="") then
		response.write "<br clear=left><p align=left><img vspace=4 hspace=4 align=left vspace=4 hspace=10 src=" & chr(34) & "RodentImages/" & rs("Picture2") & chr(34) & ">"
	end if
	If not (isnull(rs("Picture3")) or rs("Picture3")="") then
		response.write "<br clear=left><img vspace=4 hspace=4 src=" & chr(34) & "RodentImages/" & rs("Picture3") & chr(34) & ">"
	end if
	If not (isnull(rs("Picture4")) or rs("Picture4")="") then
		response.write "<br clear=left><img vspace=4 hspace=4 src=" & chr(34) & "RodentImages/" & rs("Picture4") & chr(34) & ">"
	end if

	response.write "<hr>"


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
      1999 House Rabbit Society - Washington </td>
  </tr>
</table>
</center></div>
<!--webbot bot="Include" endspan i-checksum="29637" -->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>