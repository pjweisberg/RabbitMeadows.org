<% Option Explicit
'###########################################################
'## COPYRIGHT (C) 2002-2005, Brinkster Site Statistics Corp.
'## 
'## For licensing details, lease read the license.txt file 
'## included with MetaTraffic or located at:
'## http://www.metasun.com/products/metatraffic/license.asp
'##
'## All copyright notices regarding MetaTraffic 
'## must remain intact in the scripts and in the 
'## outputted HTML. All text and logos with
'## references to Metasun or MetaTraffic must
'## remain visible when the pages are viewed on 
'## the internet or intranet.
'##
'## For support, please visit http://www.metasun.com
'## and use the support forum.
'###########################################################
%>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<%
Dim strUsername : strUsername = Request.Form("username")
Dim strPassword : strPassword = Request.Form("password")
Dim blnRemember : blnRemember = Request.Form("remember")
Dim strAction : strAction = UCase(Request("action"))

Dim strChecked, strError

If blnRemember = "ON" Then
	blnRemember = True
	strChecked = " checked"
Else 
	blnRemember = False
End If

Select Case strAction

Case "LOGIN"

	If strUsername <> "" Then
	
		Response.Cookies("metatraffic")("username")	= strUsername
		Response.Cookies("metatraffic")("password")	= strPassword
		
		Call CreateDatabaseConnection(1)
		Dim blnAdmin : blnAdmin = CInt(Authenticate(False, aryMTDB(5)))
		Call CloseDatabaseConnection()
		
		If blnRemember = True Then
			Response.Cookies("metatraffic").expires = dateadd("d", 365, date())
		End If
		
		' REDIRECT
		Response.Redirect "default.asp"
	End If

Case "LOGOUT"

	Response.Cookies("metatraffic").expires = DateAdd("d", -1, Now())
	strError = "<p class=error>You have been logged out</p>"

Case "FAILURE"
	
	Dim intCode : intCode = CInt(Request.Querystring("code"))
	If intCode = 0 Then
		strError = "<p class=error>Invalid username or password.</p>"
	ElseIf intCode = -1 Then
		strError = "<p class=error>Insufficient priviledges</p>"
	ElseIf intCode = -2 Then
		strError = "<p class=error>Please log in</p>"
	End If

End Select

' RETRIEVE USERNAME / PASSWORD FROM COOKIES
strUsername = Request.Cookies("metatraffic")("username")
strPassword = Request.Cookies("metatraffic")("password")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Site Statistics</title>
	<link rel="stylesheet" href="style.css" type="text/css">
	<style type="text/css">
	body {
		background-color:black;
	}
	</style>
</head>

<body style="padding-top: 50px;">

<form method=post action="login.asp">
<table border=0 cellpadding=0 cellspacing=0 class=login align=center>
<tr class=pgheader>
	<td><img src="images/mt_logo.gif" border=0 alt="MetaTraffic" width=300 height=60></td>
</tr>
<% If strError <> "" Then %>
<tr>
	<td bgcolor="#2e2e2e" align=center style="padding: 10px;"><% Response.Write(strError) %></td>
</tr>
<% End If %>
<tr valign=top>
	<td bgcolor="#2e2e2e" align=center style="padding: 10px;">
	<table border=0 cellpadding=2 cellspacing=0 align=center>
	<tr>
		<td align="right"><span style="color:white;"><p>Username: </p></span></td>
		<td align=left><input type=text name=username value="<% = strUsername %>" maxlength=20 size=15></td>
	</tr>
	<tr>
		<td align=right><span style="color:white;"><p>Password: </p></span></td>
		<td align=left><input type=password name=password value="<% = strPassword %>" maxlength=20 size=15></td>
	</tr>
	<tr valign=top>
		<td align=center colspan=2>
		<span style="color:white;"><p><input type=checkbox name=remember value="ON" class=checkbox <% = strChecked %>>&nbsp;Remember login information</p></span>
		</td>
	</tr>
	<tr>
		<td colspan=2 align=center>
		<input type=image name=login src="images/login_btn.gif" value="Login" border=0>
		<input type=hidden name=Action Value="Login">
		</td>
	</tr>
	</table>
	</td>
</tr>
<tr class=pgfooter>
	<td align=center>
	<div style="padding:5px;">
	<span class=copyright>&copy; Copyright 2007, </span>
	<span style="color:#666666; font-size:10px;">Brinkster Site Statistics</span>
	</div></td>
</tr>
</table>
</form>
</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>