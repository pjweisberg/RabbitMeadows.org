<% @language=vbscript %>
<% option explicit %>
<% Response.Buffer="true" %>
<!--#include file="connstr.asp"-->

<html>
<head>
	<title>Best Little Rabbit, Rodent and Ferret House Administration </title>

</head>

<body>

<center>
<h2>
Best Little Rabbit, Rodent and Ferret House Administration Pages
</h2><br>
<%
If session("UserLevel")=2 then
	If request("button")="logout" then
		Session("UserLevel")=false
		response.write "<center><h1><b>You are Logged Out</b><br>"
		response.write "<a href=""default.asp"">login</a><br><br>"
		response.write "<a href=""/rabbitrodentferret.org/index.asp"">Back to the Web Page</a></center>"
	else
		Call DisplayMenu()
	end if
else
'begin new code
Dim adminpassword, setright

adminpassword = Request.Form("pwd")

IF adminpassword <> "" then

    dim sqlStringP, objConn, objRSP

    sqlStringP = "SELECT AdminID FROM Admin " &_
    "WHERE adminPassword='" & adminpassword & "'" 

    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionString=sConnect
    objConn.Open 
    SET objRSP = Server.CreateObject("ADODB.Recordset")
    objRSP.Open sqlStringP, objConn

      IF objRSP.EOF THEN 
      setright="no"
	  else
	  setright="yes"
	  end if
     
      	objRSP.Close
	Set objRSP=nothing
	objConn.Close
	set objConn=nothing
end if
'end of new code

	If request("button")="login" then
		If setright="yes" then
			Session("UserLevel")=2
			call DisplayMenu()
		else
			
			response.write "<center><h1><b>Login Incorrect</b><h1></center><br>"
			Call DisplayLogin()
		end if
	else
		call DisplayLogin()
	end if
end if

Sub DisplayMenu()
%>
	<table border="1" cellpadding="5" cellspacing="0">
	
		<td align="center" valign="top">
			<font size="+1">
			<a HREF="adminfaq.asp">FAQ</a><br>
			<a HREF="adminNews.asp">News</a><br>
			<a HREF="adminVets.asp">Vets</a><br>
			<a HREF="adminSites.asp">Other Rabbit Sites</a><br>
			<a HREF="adminAdopt.asp">Adopt Rabbits</a><br>
			<a HREF="adminAdoptRod.asp">Adopt Rodents</a><br>
			<a HREF="adminAdoptPig.asp">Adopt Guinea Pigs</a><br>
			<a HREF="adminAdoptFer.asp">Adopt Ferrets</a><br>
			<a HREF="updatep.asp">Manage Products</a><br>
			<a HREF="updatec.asp">Manage Featured Companions</a><br>
			</font>
		</td>
		<td align="center">
			&nbsp;
		</td>
	</tr>
	</table>
	
	<form action="default.asp" method="POST"><input type="Submit" name="button" value="logout"></form>
<%
End sub

sub DisplayLogin()
%>
	<center>
	<h1>Best Little Rabbit, Rodent and Ferret House Admin login<h1><br>
	<form action="default.asp" method="POST">
	Password:<br>
	<input type="Password" name="pwd" size="25" maxlength="25"><br>
	<input type="hidden" name="button" value="login">
	<input type="Submit" name="NewButton" value="login">
	</form>
	</center>
<%
end sub
%>
</body>
</html>

																																																																																	