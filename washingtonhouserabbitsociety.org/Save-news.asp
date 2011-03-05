<html>

<head>
<title>HouseRabbit.org - News</title>

<meta NAME="Keywords" CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, rabbit information, rescue, rabbit rescue, rodent rescue,
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
    <td width="25%" rowspan="4" valign="top"><!--webbot bot="Include" U-Include="HRSMenu.htm" TAG="BODY" startspan -->

<!--#include file="hrsmenu1.htm"-->
<!--webbot bot="Include" endspan i-checksum="14527" -->

</td>
    <td width="10%" rowspan="4"></td>
    <td width="50%"><font face="comic sans ms, arial" color="#400040" size="5"><b>The Latest
    House Rabbit News</b></font></td>
  </tr>
  <tr>
    <td width="50%" valign="top"><%
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
	sql="Select * from rebeccad.News where id=" & request("id")
	on error resume next
	set rsNews=conn.execute(sql)
	on error goto 0
	If rsNews.eof then
		response.write "No News with that number found"
	else
%>
<p><font face="comic sans ms, arial" color="#400040" size="2"><%
		response.write "<b>"
		response.write rsNews("title") & "</b><br>"
		response.write formatdatetime(rsNews("Date"),2) & "<br>"
		response.write replace(rsNews("Text"),chr(13),"<br>")
%></font>
<%
	end if
else
	sql="Select Id,Title,Date from rebeccad.News order by Date DESC"
	set rsNews=conn.execute(sql)
%>    </p>
    <table border="0" width="100%">
      <tr>
<%
	do while not rsNews.eof
		response.write "<tr><td>"
%>
        <td><img src="images/btn_rabbit.gif" alt="[Rabbit Button]" WIDTH="16" HEIGHT="16"> <font face="comic sans ms, arial" color="#400040" size="2"><%
		response.write "<a href=""news.asp?id=" & rsNews("id") & chr(34) & ">" & formatdatetime(rsNews("Date"),2) & " - " & rsNews("Title") & "</a></td></tr>"
		rsNews.movenext
	loop
%> </font></td>
      </tr>
    </table>
<%
end if
'--------------------------------------------------------------------
' Close everything
'--------------------------------------------------------------------
set rsNews=nothing
Conn.close
set Conn=nothing

%>
    </td>
  </tr>
  <tr>
    <td width="50%"></td>
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
    <td width="100%" align="center"><font face="comic sans ms, arial" size="2" color="#400040">Copyright
      1999 House Rabbit Society - Washington </td>
  </tr>
</table>
</center></div>
<!--webbot bot="Include" endspan i-checksum="29637" -->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>