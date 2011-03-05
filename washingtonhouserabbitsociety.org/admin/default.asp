<html>
<head>
	<title>HouseRabbit.com Admin Pages</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/sumipntg/sumi1011.css"><meta name="Microsoft Theme" content="sumipntg 1011, default">
</head>

<body>
<center>
<h2>
Welcome to the HouseRabbit.org Admin Pages
</h2><br>
<%
If session("UserLevel")=2 then
	If request("button")="logout" then
		Session("UserLevel")=false
		response.write "<center><h1><b>You are Logged Out</b><br>"
		response.write "<a href=""default.asp"">login</a><br><br>"
		response.write "<a href=""/default.htm"">Back to the Web Page</a></center>"
	else
		Call DisplayMenu()
	end if
else
	If request("button")="login" then
		If request("pwd")="Bungie1" then
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
	<tr>
		<td align="center" valign="top">
			<h2>Admin Pages</h2>
		</td>
		<td align="center" valign="top">
			<h2>Reports</h2>
		</td>
	</tr>
	<tr>
		<td align="center" valign="top">
			<font size="+1">
			<a HREF="adminfaq.asp">FAQ</a><br>
			<a HREF="adminNews.asp">News</a><br>
			<a HREF="adminVets.asp">Vets</a><br>
			<a HREF="adminSites.asp">Other Rabbit Sites</a><br>
			<a HREF="adminAdopt.asp">Adopt Rabbits</a><br>
			<a HREF="adminAdoptRod.asp">Adopt Rodents</a><br>
			<a HREF="adminAdoptFer.asp">Adopt Ferrets</a><br>
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
	<h1>HouseRabbit.org Admin Login<h1><br>
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

																																																																																	