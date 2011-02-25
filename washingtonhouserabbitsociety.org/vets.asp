<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Best Little Rabbit, Rodent and Ferret Vet Referral</title>
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
</head>

<body>
<div align="center"><center>

<table border="0" width="90%">
<tr><td colspan="3"
<!--#include file="headerfile.asp"-->
</td></tr>
  <tr>
    
    <td width="10%" rowspan="3"></td>
    <td width="50%" valign="top"><font face="comic sans ms, arial" color="#400040" size="5"><b>Vet
    Referrals...</b></font></td>
  </tr>
  <tr>
    <td width="50%"><font face="comic sans ms, arial" size="2" color="#400040"><br>
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
sql="Select * from Vet order by ID"
set rs=conn.execute(sql)
do while not rs.eof
response.write "<b>"
	response.write rs("HEaderLocation") & " - "
	response.write rs("name") & "</b><br>"
	response.write rs("Location") & "<br>"
	response.write rs("Address") & "<br>"
	response.write rs("City") & "," & "WA " & rs("zip") & "<br>"
	response.write "Phone:" & rs("Phone")
	If not isnull(rs("Phone2")) then 
		response.write " , " & rs("Phone2")
	end if
	response.write "<br>"
	response.write "fax:" & rs("Fax") & "<br>"
	response.write "email:" & rs("email") & "<br>"
	response.write "hours:" & rs("hours") & "<br><br>"
	response.write rs("Comments") & "<br>"
	response.write "<hr>"


	rs.movenext
loop

'--------------------------------------------------------------------
' Close everything
'--------------------------------------------------------------------
set rs=nothing
Conn.close
set Conn=nothing

%>    </font></td>
  </tr>
  <tr>
    <td width="50%"></td>
  </tr>
</table>
</center></div>

<hr width="90%">
<!--webbot bot="Include" U-Include="_private/footer.htm" TAG="BODY" startspan -->
<div align="center"><center>

<table border="0" width="90%">
  <tr>
  <td>
  <p>&nbsp;
    <!--#include file="footer.asp"-->
	</td>
  </tr>
</table>
</center></div>
<!--webbot bot="Include" endspan i-checksum="29637" -->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>