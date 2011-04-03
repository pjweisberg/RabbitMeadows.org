<%@ LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>
<!-- #include file="correct-domain.asp"-->
<!--#include file="connstr.asp" -->

<%

'---------------------------------------------------
'Read in which animal from link
'----------------------------------------------------

Dim  intAnimal


intAnimal=Request.Querystring("Animal")


If intAnimal="" Then
intAnimal=5
End If

'----------------------------------------------------
'Convert animal type and catagory into text for display
'------------------------------------------------------

Function GetAnimal(intAnimal)
Select Case intAnimal
  Case 1:
	GetAnimal="RABBIT"
  Case 2:
	GetAnimal="RODENT"
  Case 4:
	GetAnimaL="GUINEA PIG"
  Case 5:
	GetAnimal=""
End Select
End Function
%>



<html>

<head>

<title>Best Little Rabbit, Rodent and Ferret Vet Referral</title>
<!--#include file="dropdownmenu.asp"-->
<meta NAME="Keywords" CONTENT="House Rabbit, House Rabbit Society, rabbits, bunnies, pets, pet adoption, shelter, humane,
         pet rabbits, rabbit health, vets, non-profit, rabbit information, rescue, rabbit rescue, rodent rescue,
         pet rats, pet mice, pet hamsters, pet gerbils, pet guinea pigs">

<meta NAME="Description" CONTENT="Best Little Rabbit, Rodent and Ferret House is a the definitive site for Rescued Rabbits in the 
Northwest and other parts of the country. HRS is a non-provit organization.">

<!--#include file="google-analytics.js"-->
</head>


<body>
<div align="center"><center>

<table border="0" width="90%">
<tr><td colspan="3">
<!--#include file="headerfile.asp"-->
</td></tr>
  <tr>
    
    <td width="10%" rowspan="25"></td>
    <td  valign="top"><font face="arial" color="#FF9933"><H1><%=GetAnimal(intAnimal)%> Vet
    Referrals</H1></font></td>
  </tr>
  
    
<%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------



	
	
Dim sql, Conn, rs

Select Case intAnimal
	Case 1:
	sql="Select * from Vet where Rabbits=1 order by City"

	Case 2:
	sql="Select * from Vet where Rodents=1 order by City"

	Case 4:
	sql="Select * from Vet where Guineas=1 order by City"

	Case 5:
	sql="Select * from Vet order by City"

	End Select
Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open sConnect
'--------------------------------------------------------------------
' Display All the Records
'--------------------------------------------------------------------

set rs=conn.execute(sql)



If Not rs.EOF THEN

do while not rs.eof
response.write "<tr><td>"
Response.write "<font face=""arial"" size=""3"" color=""000000"">"
response.write "<b>"
	response.write rs("HEaderLocation") & " - "
	response.write rs("name") & "</b><font size=""2""><br>"
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
	response.write "</font></td><td>"
	if rs("Photo")<>"" then
	response.write "<img src=""Vetphotos/"
	
	response.write rs("Photo") & """" & ">"
	end if
	response.write "</td></tr>"
	response.write "<tr><td colspan=2>"
	response.write rs("Comments") & "<br>"
	response.write "<hr>"
	response.write "</td></tr>"
	


	rs.movenext
loop

else
response.write "No records"

end if

'--------------------------------------------------------------------
' Close everything
'--------------------------------------------------------------------
set rs=nothing
Conn.close
set Conn=nothing

%>    
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
