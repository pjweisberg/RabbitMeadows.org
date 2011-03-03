<% OPTION EXPLICIT 
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
<!--#Include File="core.asp"-->
<%
Server.ScriptTimeout = 36000
Response.Buffer = True

Dim blnForm, strError, objConn, objConn2, strSql, strClass

Dim strAction : strAction = Request.Form("action")
Dim intInstall : intInstall = CInt(Request.Form("install"))
Dim strUsername : strUsername = Request.Form("username")
Dim strPassword : strPassword = Request.Form("password")
Dim strPassword2 : strPassword2 = Request.Form("password2")
Dim strDBType : strDBType = Request.Form("dbtype")
Dim strDBLocation : strDBLocation = Request.Form("dblocation")
Dim strDBName : strDBName = Request.Form("dbname")
Dim strDBUsername : strDBUsername = Request.Form("dbusername")
Dim strDBPassword : strDBPassword = Request.Form("dbpassword")
Dim strTablePrefix : strTablePrefix = Request.Form("dbprefix")
Dim intDBCreate : intDBCreate = CInt(Request.Form("dbcreate"))
Dim intDBDefinitions : intDBDefinitions = CInt(Request.Form("dbdefinitions"))
Dim intDBCountries : intDBCountries = CInt(Request.Form("dbcountries"))
Dim intDBConfig : intDBConfig = CInt(Request.Form("dbconfig"))
Dim intUpgradeType : intUpgradeType = CInt(Request.Form("upgradetype"))
Dim strDB2Type : strDB2Type = Request.Form("db2type")
Dim strDB2Location : strDB2Location = Request.Form("db2location")
Dim strDB2Name : strDB2Name = Request.Form("db2name")
Dim strDB2Username : strDB2Username = Request.Form("db2username")
Dim strDB2Password : strDB2Password = Request.Form("db2password")
Dim strTable2Prefix : strTable2Prefix = Request.Form("db2prefix")

If strAction = "" Then
	intInstall 			= 1
	strDBType 			= "MSACCESS"
	strDBLocation 		= "c:\sites\SERVERNAME\USERNAME\database"
	strDBName			= "stats.mdb"
	strTablePrefix		= "mt_"
	intDBCreate 		= 0
	intDBDefinitions 	= 1
	intDBCountries 		= 1
	intDBConfig			= 1
	intUpgradeType 		= 1
	strDB2Type 			= "MSACCESS"
	strDB2Location		= "c:\sites\SERVERNAME\USERNAME\database"
	strDB2Name			= "stats.mdb"
End If

If intInstall = 1 Then
	strClass = "display:none;"
End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Site Statistics Setup</title>
	<link rel="stylesheet" href="style.css" type="text/css">
	<script language="JavaScript" src="javascript.js" type="text/javascript"></script>
</head>
<%
If strAction <> "" Then

	If intDBConfig = 1 And intInstall <> 3 Then
		If Len(strUsername) = 0 Then
			strError = strError & "<li>Username is a required field.</li>"
		End If
		
		If Len(strPassword) = 0 Then
			strError = strError & "<li>Password is a required field.</li>"
		End If
		
		If strPassword <> strPassword2 Then
			strError = strError & "<li>Passwords do not match.</li>"
		End If
	End If
	
	If strError = "" Then
		
		Dim intStep : intStep = 1
		
		With Response
			.Write("<body style=""padding:5px;"">" & vbcrlf)
			.Write("<span class=name>Site Statistics Setup</span>")
			.Write("<p><strong>STEP " & intStep & ":</strong> Connecting to new database...")
			Response.Flush : intStep = intStep + 1
			Call CreateDatabaseConnection(strDBType, strDBLocation, strDBName, strDBUsername, strDBPassword, 1)
			
			If intInstall = 2 Then
				.Write("<p><strong>STEP " & intStep & ":</strong> Connecting to old database...")
				Response.Flush : intStep = intStep + 1
				Call CreateDatabaseConnection(strDB2Type, strDB2Location, strDB2Name, strDB2Username, strDB2Password, 2) : Response.Flush
			End If
			
			.Write("<p><strong>STEP " & intStep & ":</strong> Testing configuration file permissions...")
			Response.Flush : intStep = intStep + 1
			Call TestFile() : Response.Flush
			
			.Write("<p><strong>STEP " & intStep & ":</strong> Creating database connection file...")
			Response.Flush : intStep = intStep + 1
			Call CreateConnectionFile(strDBType, strDBLocation, strDBName, strDBUsername, strDBPassword, strTablePrefix) : Response.Flush
			
			If intInstall < 3 Then
				If strDBType <> "MSACCESS" Then
					If intDBCreate = 1 Then
						.Write("<p><strong>STEP " & intStep & ":</strong> Creating new database tables...")
						Response.Flush : intStep = intStep + 1
						Call SetupDatabase(strDBType, strTablePrefix) : Response.Flush
					End If
				End If
			ElseIf intInstall = 3 Then
				Dim strVersion
				.Write("<p><strong>STEP " & intStep & ":</strong> Checking Site Statistics 2.x Lite database...")
				Response.Flush : intStep = intStep + 1
				Call TestDatabaseLite(strDBType, strTablePrefix) : Response.Flush
				If (strDBType = "MSACCESS") Or (intDBCreate = 1 And strDBType <> "MSACCESS") Then
					.Write("<p><strong>STEP " & intStep & ":</strong> Upgrading database tables...")
					Response.Flush : intStep = intStep + 1
					Call UpgradeDatabase(strDBType, strTablePrefix) : Response.Flush
				End If
			End If
			
			.Write("<p><strong>STEP " & intStep & ":</strong> Testing database permissions...")
			Response.Flush : intStep = intStep + 1
			Call TestDatabase(strDBType, strTablePrefix) : Response.Flush
			
			If intDBConfig = 1 Then
				If intInstall < 3 Then
					.Write("<p><strong>STEP " & intStep & ":</strong> Loading configuration data...")
					Response.Flush : intStep = intStep + 1
					Call SetupConfig(strDBType, strTablePrefix) : Response.Flush
				ElseIf intInstall = 3 Then
					.Write("<p><strong>STEP " & intStep & ":</strong> Upgrading configuration data...")
					Response.Flush : intStep = intStep + 1
					Call UpgradeConfig(strDBType, strTablePrefix) : Response.Flush
					.Write("<p><strong>STEP " & intStep & ":</strong> Writing new configuration data...")
					Response.Flush : intStep = intStep + 1
					Call WriteConfig(strTablePrefix) : Response.Flush
				End If
			End If
			
			If intDBDefinitions = 1 Then
				.Write("<p><strong>STEP " & intStep & ":</strong> Loading definition data...")
				Response.Flush : intStep = intStep + 1
				Call UpdateDefinitions() : Response.Flush
			End If
			
			If intDBCountries = 1 Then
				.Write("<p><strong>STEP " & intStep & ":</strong> Loading country data (this will take a few minutes)...")
				Response.Flush : intStep = intStep + 1
				Call UpdateCountries() : Response.Flush
			End If
			
			If intInstall = 2 Then
				.Write("<p><strong>STEP " & intStep & ":</strong> Upgrading data (this could take a while)...</p>")
				Call UpgradeData(intUpgradeType)
				Response.Flush
			End If

			Call CloseDatabaseConnection(1)
			Call CloseDatabaseConnection(2)

			.Write("<br><table border=0 cellpadding=5 cellspacing=0 class=settings width=300>")
			.Write("<tr><th align=left>Congratulations!</th></tr>")
			.Write("<tr><td><p>Setup is complete. You can now <a href=""default.asp"">login to Site Statistics</a>. ")
			.Write("Once you have logged in, you should check your Settings.</p>")
			.Write("<p><strong>WARNING!</strong> Delete this file (setup.asp) ")
			.Write("to prevent unauthorized access.</p></td></tr></table>")
			
		End With
	Else
		blnForm = True
	End If

Else
	blnForm = True
End If
%>

<% If blnForm Then %>
<body onLoad="setupform();">
<table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
<tr id="header" class=pgheader height=44>
	<td colspan=2><img src="images/metatraffic_logo.gif" width="267" height="44"></td>
</tr>
<tr height=22>
	<td colspan=2>
		<table border=0 cellpadding=0 cellspacing=0 width="100%">
		<tr>
			<td width=61><img src="images/subnav_pointer.gif" width="61" height="22"></td>
			<td background="images/subnav_scale.gif" width="100%"><span class=sitename>&nbsp;Version 2.23&nbsp;</span></td>
		</tr>
		</table>
	</td>
</tr>
<form action="setup.asp" method=post onSubmit="return agreesetup(document.setup.install.options[document.setup.install.selectedIndex].value);" name=setup>
<tr valign=top>
	<td style="padding: 5px;" width=180>
		<table border=0 cellpadding=0 cellspacing=0 class=select width=180>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0 width="100%">
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td width="20"><img src="images/grey_arrow.gif"></td>
							<td class=header>Welcome</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class=chooser style="padding:5px;">
					<p class=about>Welcome to Site Statistics setup. This will install or upgrade 
					to Site Statistics v2.23!</p>
					<p class=about>
					<a href="javascript:showhelp('setup','')"><img src="images/help.gif" alt="Help" border=0 align=left></a>
					<strong>Tip:</strong> For help, click on the 
					help image.</p>
					</td>
				</tr>
				</table>
		</table>
	</td>
	<td style="padding:5px;" width="100%">
		<table border=0 cellpadding=3 cellspacing=0>
		<tr><td>
			<table border=0 cellpadding=0 cellspacing=0>
			<tr><td width=22><img src="images/white_arrow.gif"></td>
			<td width="100%"><span class=name>Site Statistics Setup</td>
			<td align=right width=24><a href="javascript:showhelp('setup','')"><img src="images/help.gif" alt="Help" border=0></a></td>
			</tr></table>
		</td></tr>
		<tr><td>
			<table border=0 cellpadding=5 cellspacing=0 class=settings width=400>
			<tr><td>
				<table border=0 cellpadding=5 cellspacing=0 width="100%" class=text>
					<%
					If strError <> "" Then
						With Response
							.Write("<tr><td colspan=2><p class=error>The following errors occurred: </p>")
							.Write("<ul>" & strError & "</ul></td></tr>")
						End With
					End If
					%>
					<tr><th colspan=2 align=left>General</th></tr>
					<tr>
						<td>Installation Type: </td>
						<td>
							<select name=install onChange="installform();">
							<option value=1<% Call SetSelect(1, intInstall) %>>New Install</option>
							<option value=2<% Call SetSelect(2, intInstall) %>>Upgrade (From Site Statistics 1.2 or 1.3x)</option>
							<option value=3<% Call SetSelect(2, intInstall) %>>Upgrade (From Site Statistics 2.x Lite)</option>
							</select>
						</td>
					</tr>
					<tr id=username>
						<td>Username: </td>
						<td><input type=text name=username value="<% = strUsername %>" size=20 maxlength=20></td>
					</tr>
					<tr id=password>
						<td>Password: </td>
						<td><input type=password name=password value="<% = strPassword %>" size=20 maxlength=20></td>
					</tr>
					<tr id=password2>
						<td>Confirm Password: </td>
						<td><input type=password name=password2 value="<% = strPassword2 %>" size=20 maxlength=20></td>
					</tr>
					<tr><th colspan=2 align=left>Database Information</th></tr>
					<tr>
						<td>Database Type: </td>
						<td>
							<select name=dbtype onChange="setupform();">
							<option value="MSACCESS"<% Call SetSelect("MSACCESS", strDBType) %>>MS Access</option>
							<option value="MSSQL"<% Call SetSelect("MSSQL", strDBType) %>>Microsoft SQL Server</option>
							<option value="MYSQL"<% Call SetSelect("MYSQL", strDBType) %>>mySql v4.1 or Greater</option>
							</select>
						</td>
					</tr>
					<tr>
						<td>Database Location: </td>
						<td><input type=text name=dblocation value="<% = strDBLocation %>" size=40 maxlength=255></td>
					</tr>
					<tr>
						<td>Database Name: </td>
						<td><input type=text name=dbname value="<% = strDBName %>" size=40 maxlength=255></td>
					</tr>
					<tr id=dbusername>
						<td>Database Username: </td>
						<td><input type=text name=dbusername value="<% = strDBUsername %>" size=20 maxlength=50></td>
					</tr>
					<tr id=dbpassword>
						<td>Database Password: </td>
						<td><input type=password name=dbpassword value="<% = strDBPassword %>" size=20 maxlength=50></td>
					</tr>
					<tr id=dbprefix>
						<td>Table Prefix: </td>
						<td><input type=text name=dbprefix value="<% = strTablePrefix%>" size=20 maxlength=20></td>
					</tr>
				</table>
			</td></tr>
			<tr id=upgrade style="<%=strClass%>"><td>
				<table border=0 cellpadding=5 cellspacing=0 width="100%" class=text>
					<tr><th colspan=2 align=left>Upgrade Information</th></tr>
					<tr><td colspan=2><p>If you are upgrading from a previous version of Site Statistics, enter your old database configuration data 
					below. This will be used to move the data from the previous Site Statistics installation to the new one. </p></td></tr>
					<tr>
						<td>Upgrade Type: </td>
						<td>
							<select name=upgradetype>
							<option value=1<% Call SetSelect(1, intUpgradeType) %>>Upgrade All Data</option>
							<option value=2<% Call SetSelect(2, intUpgradeType) %>>Upgrade Last Week of Data</option>
							<option value=3<% Call SetSelect(3, intUpgradeType) %>>Upgrade Last Month of Data</option>
							<option value=4<% Call SetSelect(4, intUpgradeType) %>>Upgrade Last 3 Months of Data</option>
							<option value=5<% Call SetSelect(5, intUpgradeType) %>>Upgrade Last 6 Month of Data</option>
							<option value=6<% Call SetSelect(6, intUpgradeType) %>>Upgrade Last Year of Data</option>
							</select>
						</td>
					</tr>
					<tr>
						<td>Database Type: </td>
						<td>
							<select name=db2type onChange="setupform();">
							<option value="MSACCESS"<% Call SetSelect("MSACCESS", strDB2Type) %>>MS Access</option>
							<option value="MSSQL"<% Call SetSelect("MSSQL", strDB2Type) %>>Microsoft SQL Server</option>
							<option value="MYSQL"<% Call SetSelect("MYSQL", strDB2Type) %>>mySql v4.1 or Greater</option>
							</select>
						</td>
					</tr>
					<tr>
						<td>Database Location: </td>
						<td><input type=text name=db2location value="<% = strDB2Location %>" size=40 maxlength=255></td>
					</tr>
					<tr>
						<td>Database Name: </td>
						<td><input type=text name=db2name value="<% = strDB2Name %>" size=40 maxlength=255></td>
					</tr>
					<tr id=db2username>
						<td>Database Username: </td>
						<td><input type=text name=db2username value="<% = strDB2Username %>" size=20 maxlength=50></td>
					</tr>
					<tr id=db2password>
						<td>Database Password: </td>
						<td><input type=password name=db2password value="<% = strDB2Password %>" size=20 maxlength=50></td>
					</tr>
					<tr id=db2prefix>
						<td>Table Prefix: </td>
						<td><input type=text name=db2prefix value="<% = strTable2Prefix %>" size=20 maxlength=20></td>
					</tr>
				</table>
			</td></tr>
			<tr id=options><td>
				<table border=0 cellpadding=5 cellspacing=0 width="100%" class=text>
					<tr><th colspan=2 align=left>Options</th></tr>
					<tr><td colspan=2 align=left>The options shown below are for advanced usage only. They do not 
					normally need to be changed.</td></tr>
					<tr id=dbcreate>
						<td align=right><input type=checkbox name=dbcreate value=1 <% Call SetCheckbox(1, intDBCreate) %> class=checkbox>&nbsp;</td>
						<td>Create database tables</td>
					</tr>
					<tr id=dbconfig>
						<td align=right><input type=checkbox name=dbconfig value=1 <% Call SetCheckbox(1, intDBConfig) %> class=checkbox>&nbsp;</td>
						<td>Load Configuration Data</td>
					</tr>
					<tr>
						<td align=right><input type=checkbox name=dbdefinitions value=1 <% Call SetCheckbox(1, intDBDefinitions) %> class=checkbox>&nbsp;</td>
						<td>Load Definition Data</td>
					</tr>
					<tr>
						<td align=right><input type=checkbox name=dbcountries value=1 <% Call SetCheckbox(1, intDBCountries) %> class=checkbox>&nbsp;</td>
						<td>Load Country Data</td>
					</tr>
				</table>
			</td></tr>
			<tr>
				<td align=center style="padding:5px;"><input type=submit name=action value="Setup Site Statistics"></td>
			</tr>
			</table>
		</td></tr>
	</table>
</td>
</tr>
</form>
<tr class=pgfooter height=22>
	<td valign=middle align=right colspan=3><span class=copyright>&copy; Copyright 2007, </span>
	<a href="http://www.metasun.com/" target="_new">Brinkster Site Statistics</a><span style="font-size:10px; color:#737373;">, powered by MetaTraffic</span> </td>
</tr>
<tr height=4 class=pgbottom>
	<td colspan=3>&nbsp;</td>
</tr>
</table>
<% End If %>
</body>
</html>
<%
Sub SetSelect(strActual, strValue)
	If strActual = strValue Then
		Response.Write " selected"
	End If
End Sub

Sub SetCheckbox(strActual, strValue)
	If strActual = strValue Then
		Response.Write " checked"
	End If
End Sub

Sub CreateDatabaseConnection(strType, byVal strLocation, strName, strUsername, strPassword, intConn)

	Dim strSql, strConn, strLocationType, strTemp, intPort, aryServer
	Dim blnPort : blnPort = False

	If InStr(strLocation, ":") > 0 And strType <> "MSACCESS" Then
		aryServer = Split(strLocation, ":")
		strLocation = aryServer(0)
		If UBound(aryServer) > 0 Then
			If IsNumeric(aryServer(1)) Then
				intPort = Int(aryServer(1))
				If intPort > 0 Then
					blnPort = True
				End If
			End If
		End If
	End If
	
	If strType = "MSSQL" Then
	
		strConn = "DRIVER={SQL Server};" &_
			"SERVER=" & strLocation & ";"
			If blnPort = True Then
				strConn = strConn & "PORT=" & intPort & ";"
			End If
			strConn = strConn & "DATABASE=" & strName & ";" &_
			"UID=" & strUsername & ";" &_
			"PWD=" & strPassword & ";" &_
			"Provider=MSDASQL.1"
			
	ElseIf strType = "MYSQL" Then
	
		strConn = "DRIVER={MySQL ODBC 3.51 Driver};" &_
			"SERVER=" & strLocation & ";"
			If blnPort = True Then
				strConn = strConn & "PORT=" & intPort & ";"
			Else
				strConn = strConn & "PORT=3306;"
			End If
			strConn = strConn & "DATABASE=" & strName & ";" &_
			"UID=" & strUsername & ";" &_
			"PWD=" & strPassword & ";Option=16387"
			
	Else ' MSACCESS
	
		If Len(strLocation) > 2 Then
		If Mid(strLocation, 2, 1) = ":" Or Mid(strLocation, 1, 2) = "\\" Then
				strLocationType = "ABSOLUTE"
			Else
				strLocationType = "VIRTUAL"
			End If
		Else
			strLocationType = "VIRTUAL"
		End If
		
		If strLocationType = "ABSOLUTE" Then
			strTemp = strLocation & "\" & strName
		Else
			strTemp = Server.MapPath(strLocation & "/" & strName)
		End If
		
		strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTemp
	End If
	
	On Error Resume Next
	
	If intConn = 1 Then
	
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open strConn
		
	ElseIf intConn = 2 Then
	
		Set objConn2 = Server.CreateObject("ADODB.Connection")
		objConn2.Open strConn
		
	End If
	
	If Err.Number <> 0 Then
		Response.Write("<span class=failed>FAILED</span></p>")
		Call DisplayDBError(Err, strType)
	Else
		Response.Write("<span class=success><span class=success>SUCCESS</span></span></p>")
	End If
	On Error Goto 0
	
	Err.Clear

End Sub

Sub CloseDatabaseConnection(intConn)
	If intConn = 1 Then
		If IsObject(objConn) Then
			objConn.Close : Set objConn = Nothing
		End If
	ElseIf intConn = 2 Then
		If IsObject(objConn2) Then
			objConn2.Close : Set objConn2 = Nothing
		End If
	End If
End Sub

Sub CreateConnectionFile(strType, strLocation, strName, strUsername, strPassword, strTablePrefix)

	Dim strError

	Dim strPath : strPath = Request.Servervariables("Script_Name")
	strPath = Left(strPath, InStrRev(strPath, "/") - 1)
	
	Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	On Error Resume Next
	Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strPath  & "/conn.asp"), 2)
	
	strError = CheckErrors(Err.Number, Err.Description)
	
	If strError = "" Then
		With objTS
			.WriteLine("<" & Chr(37))
			.WriteLine("Dim aryMTDB(5), objConn")
			.WriteLine("aryMTDB(0) = """ & strType & """")
			.WriteLine("aryMTDB(1) = """ & strLocation & """")
			.WriteLine("aryMTDB(2) = """ & strName & """")
			.WriteLine("aryMTDB(3) = """ & strUsername & """")
			.WriteLine("aryMTDB(4) = """ & strPassword & """")
			.WriteLine("aryMTDB(5) = """ & strTablePrefix & """")
			.WriteLine(Chr(37) & ">")
		End With
		Response.Write("<span class=success>SUCCESS</span></p>")
	Else
		Call DisplayFatalError("Could not write to conn.asp file. Check the file's permissions.")
	End If
	
	On Error Goto 0
	Set objTS = Nothing : Set objFSO = Nothing
	
End Sub

Sub TestDatabase(strType, strTablePrefix)

	On Error Resume Next
	
	Call InsertConfig("Test", "Test", "Test", 3, 0, "")
	
	If Err.Number <> 0 Then
		Response.Write("<span class=failed>FAILED</span></p>")
		Call DisplayDBError(Err, strType)
	End If
	
	strSql = "DELETE FROM " & strTablePrefix & "Config WHERE c_name = 'Test'"
	objConn.Execute(strSql)
	
	If Err.Number <> 0 Then
		Response.Write("<span class=failed>FAILED</span></p>")
		Call DisplayDBError(Err, strType)
	End If
	
	Response.Write("<span class=success>SUCCESS</span></p>")
	
	On Error Goto 0
	
End Sub

Sub TestDatabaseLite(strType, strTablePrefix)

	On Error Resume Next
	
	strSql = "SELECT c_value FROM " & strTablePrefix & "Config " &_
		"WHERE c_name = 'Site Statistics_Version' AND c_value LIKE '2.% Lite'"
	Dim rsCheck : Set rsCheck = objConn.Execute(strSql)
	
	If Err.Number <> 0 Then
		Response.Write("<span class=failed>FAILED</span></p>")
		Call DisplayDBError(Err, strType)
	ElseIf rsCheck.Eof Then
		Call DisplayFatalError("Could not find Site Statistics Lite installation.")
	End If

	strVersion = rsCheck(0)
	rsCheck.Close : Set rsCheck = Nothing

	Response.Write("<span class=success>SUCCESS</span></p>")
	
	On Error Goto 0
	
End Sub

Sub TestFile()

	Dim strPath : strPath = Request.Servervariables("Script_Name")
	strPath = Left(strPath, InStrRev(strPath, "/") - 1)
	
	Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	On Error Resume Next
	Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strPath & "/config.asp"), 8)
	
	strError = CheckErrors(Err.Number, Err.Description)
	
	If strError = "" Then
		Response.Write("<span class=success>SUCCESS</span></p>")
	Else
		Call DisplayFatalError("Could not write to config.asp file. Check the file's permissions.")
	End If
	
	On Error Goto 0
	Set objTS = Nothing : Set objFSO = Nothing
	
End Sub

Sub SetupDatabase(strDBType, strTablePrefix)

	Select Case strDBType
	
	Case "MSSQL"
	
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Config] (" &_
			"[c_name] [varchar] (255) NOT NULL ," &_
			"[c_value] [varchar] (255) NOT NULL ," &_
			"[c_group] [varchar] (255) NOT NULL ," &_
			"[c_type] [tinyint] NOT NULL ," &_
			"[c_order] [smallint] NOT NULL ," &_
			"[c_extra] [varchar] (255) NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Definitions] ( " &_
			"[d_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[d_name] [varchar] (255) NOT NULL ," &_
			"[d_regexp] [varchar] (255) NOT NULL ," &_
			"[d_extra] [varchar] (255) NULL ," &_
			"[d_url] [varchar] (255) NULL ," &_
			"[d_type] [tinyint] NOT NULL" &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "IPCountry] (" &_
			"[ic_ipstart] [int] NOT NULL ," &_
			"[ic_ipend] [int] NOT NULL ," &_
			"[ic_code] [varchar] (2) NOT NULL" &_
			") ON [PRIMARY]"
		objConn.Execute(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Keywords] (" &_
			"[k_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[k_value] [varchar] (255) NOT NULL ," &_
			"[k_site] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Names] (" &_
			"[n_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[n_value] [varchar] (255) NOT NULL ," &_
			"[n_type] [tinyint] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "PageLog] (" &_
			"[pl_datetime] [datetime] NOT NULL ," &_
			"[pl_pn_id] [int] NOT NULL ," &_
			"[pl_r_id] [int] NOT NULL ," &_
			"[pl_s_id] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "PageNames] (" &_
			"[pn_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[pn_url] [varchar] (255) NOT NULL ," &_
			"[pn_page] [varchar] (255) NOT NULL ," &_
			"[pn_path] [varchar] (255) NOT NULL ," &_
			"[pn_label] [varchar] (100) NULL ," &_
			"[pn_extension] [varchar] (10) NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)

		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "ReferrerNames] (" &_
			"[rn_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[rn_page] [varchar] (255) NULL ," &_
			"[rn_host] [varchar] (255) NULL ," &_
			"[rn_domain] [varchar] (100) NULL ," &_
			"[rn_extension] [varchar] (10) NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)

		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Referrers] (" &_
			"[r_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[r_url] [varchar] (255) NOT NULL ," &_
			"[r_rn_id] [int] NOT NULL ," &_
			"[r_k_id] [int] NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "RobotLog] (" &_
			"[rl_datetime] [datetime] NOT NULL ," &_
			"[rl_pn_id] [int] NOT NULL ," &_
			"[rl_useragent] [int] NOT NULL ," &_
			"[rl_robot] [int] NOT NULL ," &_
			"[rl_ip] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Sessions] (" &_
			"[s_id] [int] NOT NULL ," &_
			"[s_ip] [int] NOT NULL ," &_
			"[s_hostname] [int] NOT NULL ," &_
			"[s_language] [varchar] (5) NOT NULL ," &_
			"[s_country] [varchar] (2) NOT NULL ," &_
			"[s_useragent] [int] NOT NULL ," &_
			"[s_browser] [int] NOT NULL ," &_
			"[s_os] [int] NOT NULL ," &_
			"[s_screenarea] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Users] (" &_
			"[u_id] [int] IDENTITY (1, 1) NOT NULL ," &_
			"[u_username] [varchar] (20) NOT NULL ," &_
			"[u_password] [varchar] (20) NULL ," &_
			"[u_admin] [bit] NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Actions] (" &_
			"[a_code] [varchar] (12) NOT NULL ," &_
			"[a_name] [varchar] (255) NOT NULL ," &_
			"[a_display] [bit] NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Campaigns] (" &_
			"[ca_code] [varchar] (12) NOT NULL ," &_
			"[ca_name] [varchar] (255) NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "ActionLog] (" &_
			"[al_datetime] [datetime] NOT NULL ," &_
			"[al_unique] [varchar] (255) NOT NULL ," &_
			"[al_amount] [money] NOT NULL ," &_
			"[al_ca_code] [varchar] (12) NOT NULL ," &_
			"[al_a_code] [varchar] (12) NOT NULL ," &_
			"[al_s_id] [int] NOT NULL ," &_
			"[al_r_id] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "CampaignLog] (" &_
			"[cl_datetime] [datetime] NOT NULL ," &_
			"[cl_ca_code] [varchar] (12) NOT NULL ," &_
			"[cl_s_id] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Actions] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Actions] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[a_code]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Campaigns] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Campaigns] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[ca_code]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Config] ADD " &_
			"CONSTRAINT [DF_" & strTablePrefix & "Config_c_order] DEFAULT (0) FOR [c_order]"
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Definitions] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Definitions] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[d_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Keywords] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Keywords] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[k_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Names] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Names] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[n_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "PageNames] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "PageNames] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[pn_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "ReferrerNames] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "ReferrerNames] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[rn_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Referrers] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Referrers] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[r_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)

		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Sessions] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Sessions] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[s_id]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE UNIQUE INDEX ix_" & strTablePrefix & "PageNames ON " & strTablePrefix & "PageNames(pn_id)"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE UNIQUE INDEX ix_" & strTablePrefix & "Referrers ON " & strTablePrefix & "Referrers(r_id)"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE UNIQUE INDEX ix_" & strTablePrefix & "ReferrerNames ON " & strTablePrefix & "ReferrerNames(rn_id)"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE UNIQUE INDEX ix_" & strTablePrefix & "Keywords ON " & strTablePrefix & "Keywords(k_id)"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE UNIQUE INDEX ix_" & strTablePrefix & "Names ON " & strTablePrefix & "Names(n_id)"
		Call ExecuteQuery(strSql)
		
	Case "MYSQL"
	
		strSql = "CREATE TABLE " & strTablePrefix & "config (" &_
			"c_name varchar(255) NOT NULL default ''," &_
			"c_value varchar(255) NOT NULL default ''," &_
			"c_group varchar(255) NOT NULL default ''," &_
			"c_type tinyint(4) NOT NULL default '0'," &_
			"c_order smallint(6) NOT NULL default '0'," &_
			"c_extra varchar(255) default NULL" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "definitions (" &_
			"d_id int(11) NOT NULL auto_increment," &_
			"d_name varchar(255) NOT NULL default ''," &_
			"d_regexp varchar(255) NOT NULL default ''," &_
			"d_extra varchar(255) NOT NULL default ''," &_
			"d_url varchar(255) NOT NULL default ''," &_
			"d_type tinyint(4) NOT NULL default '0'," &_
			"KEY d_id (d_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=40 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "ipcountry (" &_
			"ic_ipstart int(11) NOT NULL default '0'," &_
			"ic_ipend int(11) NOT NULL default '0'," &_
			"ic_code char(2) NOT NULL default ''" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "keywords (" &_
			"k_id int(11) NOT NULL auto_increment," &_
			"k_value varchar(255) NOT NULL default ''," &_
			"k_site int(11) NOT NULL default '0'," &_
			"KEY k_id (k_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=5 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "names (" &_
			"n_id int(11) NOT NULL auto_increment," &_
			"n_value varchar(255) NOT NULL default ''," &_
			"n_type tinyint(4) NOT NULL default '0'," &_
			"KEY n_id (n_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=72 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "pagelog (" &_
			"pl_datetime datetime NOT NULL default '0000-00-00 00:00:00'," &_
			"pl_pn_id int(11) NOT NULL default '0'," &_
			"pl_r_id int(11) NOT NULL default '0'," &_
			"pl_s_id int(11) NOT NULL default '0'" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "pagenames (" &_
			"pn_id int(11) NOT NULL auto_increment," &_
			"pn_url varchar(255) NOT NULL default ''," &_
			"pn_page varchar(255) NOT NULL default ''," &_
			"pn_path varchar(255) NOT NULL default ''," &_
			"pn_label varchar(100) default NULL," &_
			"pn_extension varchar(10) NOT NULL default ''," &_
			"KEY pn_id (pn_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=45 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "referrernames (" &_
			"rn_id int(11) NOT NULL auto_increment," &_
			"rn_page varchar(255) NOT NULL default ''," &_
			"rn_host varchar(255) NOT NULL default ''," &_
			"rn_domain varchar(100) NOT NULL default ''," &_
			"rn_extension varchar(10) NOT NULL default ''," &_
			"KEY rn_id (rn_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=14 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "referrers (" &_
			"r_id int(11) NOT NULL auto_increment," &_
			"r_url varchar(255) NOT NULL default ''," &_
			"r_rn_id int(11) NOT NULL default '0'," &_
			"r_k_id int(11) NOT NULL default '0'," &_
			"KEY r_id (r_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=14 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "robotlog (" &_
			"rl_datetime datetime NOT NULL default '0000-00-00 00:00:00'," &_
			"rl_pn_id int(11) NOT NULL default '0'," &_
			"rl_useragent int(11) NOT NULL default '0'," &_
			"rl_robot int(11) NOT NULL default '0'," &_
			"rl_ip int(11) NOT NULL default '0'" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "sessions (" &_
			"s_id int(11) NOT NULL default '0'," &_
			"s_ip int(11) NOT NULL default '0'," &_
			"s_hostname int(11) NOT NULL default '0'," &_
			"s_language varchar(5) NOT NULL default ''," &_
			"s_country char(2) NOT NULL default ''," &_
			"s_useragent int(11) NOT NULL default '0'," &_
			"s_browser int(11) NOT NULL default '0'," &_
			"s_os int(11) NOT NULL default '0'," &_
			"s_screenarea int(11) NOT NULL default '0'," &_
			"UNIQUE KEY s_id (s_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "users (" &_
			"u_id int(11) NOT NULL auto_increment," &_
			"u_username varchar(20) NOT NULL default ''," &_
			"u_password varchar(20) NOT NULL default ''," &_
			"u_admin tinyint(4) NOT NULL default '0'," &_
			"PRIMARY KEY  (u_id)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "actions (" &_
			"a_code varchar(12) NOT NULL default ''," &_
			"a_name varchar(255) NOT NULL default ''," &_
			"a_display tinyint(4) NOT NULL default '0'," &_
			"PRIMARY KEY (a_code)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;"
		Call ExecuteQuery(strSql)

		strSql = "CREATE TABLE " & strTablePrefix & "campaigns (" &_
			"ca_code varchar(12) NOT NULL default ''," &_
			"ca_name varchar(255) NOT NULL default ''," &_
			"PRIMARY KEY (ca_code)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "actionlog (" &_
			"al_datetime datetime NOT NULL default '0000-00-00 00:00:00'," &_
			"al_unique varchar(100) NOT NULL default '0'," &_
			"al_amount decimal(12,2) NOT NULL default '0' ," &_
			"al_ca_code varchar(12) NOT NULL default ''," &_
			"al_a_code varchar(12) NOT NULL default ''," &_
			"al_s_id int(11) NOT NULL default '0'," &_
			"al_r_id int(11) NOT NULL default '0'" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "campaignlog (" &_
			"cl_datetime datetime NOT NULL default '0000-00-00 00:00:00'," &_
			"cl_ca_code varchar(12) NOT NULL default ''," &_
			"cl_s_id int(11) NOT NULL default '0'" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
	End Select
	
	Response.Write("<span class=success>SUCCESS</span></p>")

End Sub

Sub UpgradeDatabase(strDBType, strTablePrefix)

	Select Case strDBType
	
	Case "MSSQL"
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Actions] (" &_
			"[a_code] [varchar] (12) NOT NULL ," &_
			"[a_name] [varchar] (255) NOT NULL ," &_
			"[a_display] [bit] NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "Campaigns] (" &_
			"[ca_code] [varchar] (12) NOT NULL ," &_
			"[ca_name] [varchar] (255) NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "ActionLog] (" &_
			"[al_datetime] [datetime] NOT NULL ," &_
			"[al_unique] [varchar] (100) NOT NULL ," &_
			"[al_amount] [money] NOT NULL ," &_
			"[al_ca_code] [varchar] (12) NOT NULL ," &_
			"[al_a_code] [varchar] (12) NOT NULL ," &_
			"[al_s_id] [int] NOT NULL ," &_
			"[al_r_id] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE [dbo].[" & strTablePrefix & "CampaignLog] (" &_
			"[cl_datetime] [datetime] NOT NULL ," &_
			"[cl_ca_code] [varchar] (12) NOT NULL ," &_
			"[cl_s_id] [int] NOT NULL " &_
			") ON [PRIMARY]"
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Actions] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Actions] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[a_code]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
		strSql = "ALTER TABLE [dbo].[" & strTablePrefix & "Campaigns] ADD " &_
			"CONSTRAINT [PK_" & strTablePrefix & "Campaigns] PRIMARY KEY  CLUSTERED " &_
			"(" &_
			"[ca_code]" &_
			")  ON [PRIMARY] "
		Call ExecuteQuery(strSql)
		
	Case "MSACCESS"
	
		strSql = "CREATE TABLE " & strTablePrefix & "Actions (" &_
			"a_code text (12) CONSTRAINT PK_a_code PRIMARY KEY," &_
			"a_name text (255)," &_
			"a_display yesno" &_
			")"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "Campaigns (" &_
			"ca_code text (12) CONSTRAINT PK_ca_code PRIMARY KEY," &_
			"ca_name text (255)" &_
			")"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "ActionLog (" &_
			"al_datetime date," &_
			"al_unique text (100)," &_
			"al_amount currency," &_
			"al_ca_code text (12)," &_
			"al_a_code text (12)," &_
			"al_s_id long," &_
			"al_r_id long" &_
			")"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "CampaignLog (" &_
			"cl_datetime date," &_
			"cl_ca_code text (12)," &_
			"cl_s_id long " &_
			")"
		Call ExecuteQuery(strSql)
		
	Case "MYSQL"
		
		strSql = "CREATE TABLE " & strTablePrefix & "actions (" &_
			"a_code varchar(12) NOT NULL default ''," &_
			"a_name varchar(255) NOT NULL default ''," &_
			"a_display tinyint(4) NOT NULL default '0'," &_
			"PRIMARY KEY (a_code)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;"
		Call ExecuteQuery(strSql)

		strSql = "CREATE TABLE " & strTablePrefix & "campaigns (" &_
			"ca_code varchar(12) NOT NULL default ''," &_
			"ca_name varchar(255) NOT NULL default ''," &_
			"PRIMARY KEY (ca_code)" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1 AUTO_INCREMENT=4 ;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "actionlog (" &_
			"al_datetime datetime NOT NULL default '0000-00-00 00:00:00'," &_
			"al_unique varchar(100) NOT NULL default '0'," &_
			"al_amount decimal (12,2) NOT NULL default '0' ," &_
			"al_ca_code varchar(12) NOT NULL default ''," &_
			"al_a_code varchar(12) NOT NULL default ''," &_
			"al_s_id int(11) NOT NULL default '0'," &_
			"al_r_id int(11) NOT NULL default '0'" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
		strSql = "CREATE TABLE " & strTablePrefix & "campaignlog (" &_
			"cl_datetime datetime NOT NULL default '0000-00-00 00:00:00'," &_
			"cl_ca_code varchar(12) NOT NULL default ''," &_
			"cl_s_id int(11) NOT NULL default '0'" &_
			") ENGINE=MyISAM DEFAULT CHARSET=latin1;"
		Call ExecuteQuery(strSql)
		
	End Select
	
	Response.Write("<span class=success>SUCCESS</span></p>")

End Sub

Sub SetupConfig(strType, strTablePrefix)

	If strType = "MSACCESS" Then
		strSql = "DELETE FROM " & strTablePrefix & "Users"
	Else
		strSql = "TRUNCATE TABLE " & strTablePrefix & "Users"
	End If
	objConn.Execute(strSql)
	
	If strType = "MSACCESS" Then
		strSql = "DELETE FROM " & strTablePrefix & "Config"
	Else
		strSql = "TRUNCATE TABLE " & strTablePrefix & "Config"
	End If
	objConn.Execute(strSql)

	Dim datSerial : datSerial = Year(Date()) & "-" & FormatDatePart(Month(Date())) & "-" & FormatDatePart(Day(Date()))
	
	Call InsertConfig("Site Statistics_Version", "2.23 Pro", "System", 0, 0, "")
	Call InsertConfig("Script_Version", "ASP 3.0", "System", 0, 1, "")
	Call InsertConfig("Install", datSerial, "System", 0, 2, "")
	Call InsertConfig("Definitions", "", "Maintenance", 1, 0, "")
	Call InsertConfig("Countries", "", "Maintenance", 1, 1, "")
	Call InsertConfig("Compact", "", "Maintenance", 1, 2, "")
	Call InsertConfig("Delete_Log", "", "Maintenance", 1, 3, "")
	Call InsertConfig("Delete_Robot_Log", "", "Maintenance", 1, 4, "")
	Call InsertConfig("Site_Name", "Insert Site Name Here", "General", 2, 0, "text||str||30||255||")
	Call InsertConfig("Site_Url", "http://www.yourdomain.com", "General", 2, 1, "text||str||30||255||")
	Call InsertConfig("Enable_Log", "-1", "Logging", 2, 2, "select||bln||Yes,No||-1,0||")
	Call InsertConfig("IP_Exclude", "", "Logging", 2, 3, "textarea||str||40||3||")
	Call InsertConfig("Querystring_Filter", "", "Logging", 2, 4, "textarea||str||40||3||")
	Call InsertConfig("Default_Doc", "", "Logging", 2, 5, "text||str||20||255||")
	Call InsertConfig("Querystring_Name", "mtc", "Logging", 2, 6, "text||str||20||255||")
	Call InsertConfig("Site_Aliases", "", "Reports", 2, 7, "textarea||str||40||3||")
	If strDBType = "MYSQL" Then
		Call InsertConfig("Session_Duration", "60", "Reports", 2, 8, "text||int||5||5||^\\d{1,5}$")
	Else
		Call InsertConfig("Session_Duration", "60", "Reports", 2, 8, "text||int||5||5||^\d{1,5}$")
	End If
	Call InsertConfig("Show_Graph", "-1", "Reports", 2, 9, "select||bln||Yes,No||-1,0||")
	Call InsertConfig("Truncate_Urls", "-1", "Reports", 2, 10, "select||bln||Yes,No||-1,0||")
	Call InsertConfig("Short_Date_Format", "mm/dd/yyyy", "Date / Time", 2, 11, "text||str||20||10||")
	Call InsertConfig("Long_Date_Format", "mmmm dd yyyy", "Date / Time", 2, 12, "text||str||30||20||")
	Call InsertConfig("Time_Offset", "0", "Date / Time", 2, 13, "select||int||-23,-22,-21,-20,-19,-18,-17,-16,-15,-14,-13,-12,-11,-10,-9,-8,-7,-6,-5,-4,-3,-2,-1,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23||||")
	
	strSql = "INSERT INTO " & strTablePrefix & "Users (u_username, u_password, u_admin) VALUES(" &_
		FormatString(strUsername, 20) & ", " &_
		FormatString(strPassword, 20) & ", -1)"
	objConn.Execute(strSql)

	Response.Write("<span class=success>SUCCESS</span></p>")
	
End Sub

Sub UpgradeConfig(strType, strTablePrefix)

	Dim datSerial : datSerial = Year(Date()) & "-" & FormatDatePart(Month(Date())) & "-" & FormatDatePart(Day(Date()))
	
	Call UpdateConfigValue("Site Statistics_Version", "2.23 Pro")
	Call UpdateConfigValue("Install", datSerial)
	Call InsertConfig("Querystring_Name", "mtc", "Logging", 2, 6, "text||str||20||255||")
	Call UpdateConfigOrder("Site_Aliases", 7)
	Call UpdateConfigOrder("Session_Duration", 8)
	Call UpdateConfigOrder("Show_Graph", 9)
	Call UpdateConfigOrder("Truncate_Urls", 10)
	Call UpdateConfigOrder("Short_Date_Format", 11)
	Call UpdateConfigOrder("Long_Date_Format", 12)

	Dim aryVersion : aryVersion = Split(strVersion, " ")
	If UBound(aryVersion) = 1 And IsNumeric(aryVersion(0)) Then
		If CSng(aryVersion(0)) < 2.20 Then
			Call InsertConfig("Time_Offset", "0", "Reports", 2, 13, "select||int||-23,-22,-21,-20,-19,-18,-17,-16,-15,-14,-13,-12,-11,-10,-9,-8,-7,-6,-5,-4,-3,-2,-1,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23||||")
		Else
			Call UpdateConfigOrder("Time_Offset", 13)
		End If
	End If

	Response.Write("<span class=success>SUCCESS</span></p>")
	
End Sub

Sub ExecuteQuery(strSql)

	On Error Resume Next
	objConn.Execute(strSql)
	If Err.Number <> 0 Then
		Response.Write("<span class=failed>FAILED</span></p>")
		Call DisplayDBError(Err, strDBType)
	End If
	On Error Goto 0
	
End Sub

Sub WriteConfig(strTablePrefix)

	Dim strName, strValue, intConfig

	strSql = "SELECT c_name, c_value, c_order, c_extra FROM " & strTablePrefix & "Config " &_
		"WHERE c_type = 2 ORDER BY c_order ASC"
	Dim rsConfig : Set rsConfig = Server.CreateObject("ADODB.RecordSet")
	rsConfig.Open strSql, objConn, 1, 2, &H0000

	If Not rsConfig.Eof Then
	
		Dim strPath : strPath = Request.Servervariables("Script_Name")
		strPath = Left(strPath, InStrRev(strPath, "/") - 1)
		
		Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
		On Error Resume Next
		Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strPath  & "/config.asp"), 2)
		
		strError = CheckErrors(Err.Number, Err.Description)
		
		If strError = "" Then
			With objTS
				.WriteLine("<" & Chr(37))
				Do While Not rsConfig.Eof
					strName = Replace(rsConfig(0), "_", "")
					objTS.WriteLine("Dim intMT" & strName & " : " & "intMT" & strName & " = " & rsConfig(2))
					intConfig = rsConfig(2)
					rsConfig.Movenext
				Loop
				objTS.WriteLine()
				objTS.WriteLine("Dim aryMTConfig(" & intConfig & ")")
				rsConfig.Movefirst
				Do While Not rsConfig.Eof
					strName = Replace(rsConfig(0), "_", "")
					aryExtra = Split(rsConfig(3), "||")
					If aryExtra(1) = "str" Then
						strValue = """" & rsConfig(1) & """"
					Else
						strValue = rsConfig(1)
					End If
					objTS.WriteLine("aryMTConfig(" & "intMT" & strName & ") = " & strValue)
					rsConfig.Movenext
				Loop
				.WriteLine(Chr(37) & ">")
			End With
			Response.Write("<span class=success>SUCCESS</span></p>")
		Else
			Call DisplayFatalError("Could not write to config.asp file. Check the file's permissions.")
		End If
		
		On Error Goto 0
		Set objTS = Nothing : Set objFSO = Nothing
	
	End If
	
	rsConfig.Close : Set rsConfig = Nothing

End Sub

Sub DisplayDBError(Err, strType)

	With Response
		.Write("<table border=0 cellpadding=5 cellspacing=0 class=settings><tr><td>")
		.Write("<p class=error>There was a database error: </p>")
		.Write("<p>Number: " & Err.Number & "<br>")
		.Write("Source: " & Err.Source & "<br>")
		.Write("Description: " & Err.Description & "</p>")
		If Err.Number = -2147467259 And strType = "MSACCESS" Then
			.Write("<p>The problem might be that there are no write permissions on the database ")
			.Write("file or that the database could not be found because of an incorrect Database Location or Name.</p>")
		End If
		.Write("<p><a href=""javascript:history.back();"">Go back</a> and try again.</p>")
		.Write("</td></tr></table>")
	End With
	
	Response.End

End Sub

Sub DisplayFatalError(strErrorMsg)

	With Response
		.Write("<span class=failed>FAILED</span></p>")
		.Write("<table border=0 cellpadding=5 cellspacing=0 class=settings><tr><td>")
		.Write("<p class=error>" & strErrorMsg & "</p>")
		.Write("<p><a href=""javascript:history.back();"">Go back</a> and try again.</p>")
		.Write("</td></tr></table>")
		.End
	End With

End Sub

Function CheckErrors(intNumber, strDescription)
	
	Dim strError

	If intNumber <> 0 Then
		strError = strError & "<li>" & strDescription & "</li>"
		Err.Clear
	End If
	
	CheckErrors = strError
		
End Function

Sub UpdateDefinitions()

	Dim strError, strLine, aryLine, intLine, strResult, strFileName
	
	Dim strInstallPath : strInstallPath = Request.Servervariables("Script_Name")
	strInstallPath = Left(strInstallPath, InStrRev(strInstallPath, "/") - 1) & "/data"

	Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	On Error Resume Next

	strFileName = "definitions.txt"
	
	Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strInstallPath) & "\" & strFileName, 1)
	
	strError = strError & CheckErrors(Err.Number, Err.Description)
	If strError <> "" Then
		Response.Write("<span class=failed>FAILED</span></p><ul>" & strError & "</ul>")
		Exit Sub
	End If
	
	If strDBType = "MSACCESS" Then
		strSql = "DELETE FROM " & strTablePrefix & "Definitions"
	Else
		strSql = "TRUNCATE TABLE " & strTablePrefix & "Definitions"
	End If
	
	Dim rsTruncate : Set rsTruncate = Server.CreateObject("ADODB.Recordset")
	rsTruncate.Open strSql, objConn, 1, 2, &H0000
	Set rsTruncate = Nothing
	
	strError = strError & CheckErrors(Err.Number, Err.Description)
	If strError <> "" Then
		Response.Write("<span class=failed>FAILED</span></p><ul>" & strError & "</ul>")
		Exit Sub
	End If
	
	On Error Goto 0
	
	If Not objTS.AtEndOfStream Then
	
		Dim strFirstLine : strFirstLine = objTS.ReadLine
		
		If Not CheckFileHeader("Definitions", strFirstLine) Then
			Response.Write("<span class=failed>FAILED</span></p><ul><li>The definitions file has does not have the correct header. Check the file and try again.</li></ul>")
			ExitSub
		End If
		
		Dim datSerial : datSerial = Right(strFirstLine, 10)
		
		Do While Not objTS.AtEndOfStream

			intLine = objTS.Line
			strLine = objTS.Readline
			aryLine = Split(strLine,"||")
			If UBound(aryLine) = 4 Then
				strSql = "INSERT INTO " & strTablePrefix & "Definitions " &_
					"(d_name, d_regexp, d_extra, d_url, d_type) VALUES(" &_
					FormatDatabaseString(Trim(aryLine(0)), 255) & ", " &_
					FormatDatabaseString(Trim(aryLine(1)), 255) & ", " &_
					FormatDatabaseString(Trim(aryLine(2)), 255) & ", " &_
					FormatDatabaseString(Trim(aryLine(3)), 255) & ", " &_
					Trim(aryLine(4)) & ")"
				objConn.Execute(strSql)
			Else
				strError = strError & "<li>Error on line " & intLine & ".</li>"
			End If
		Loop
	
	End If
	
	If strError <> "" Then
		Response.Write("OK. Update definitions partially completed, some lines had errors:</p><ul>" & strError & "</ul>")
	Else
		Response.Write("<span class=success>SUCCESS</span></p>")
		Call UpdateConfigValue("Definitions", datSerial)
	End If
	
	objTS.Close : Set objTS = Nothing
	
End Sub

Sub UpdateCountries()

	Dim strError, strLine, aryLine, intLine, strResult, strFileName
	
	Dim strInstallPath : strInstallPath = Request.Servervariables("Script_Name")
	strInstallPath = Left(strInstallPath, InStrRev(strInstallPath, "/") - 1) & "/data"

	Dim objFSO : Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	strFileName = "countries.txt"
	
	On Error Resume Next

	Dim objTS : Set objTS = objFSO.OpenTextFile(Server.MapPath(strInstallPath) & "\" & strFileName, 1)
	
	strError = strError & CheckErrors(Err.Number, Err.Description)
	If strError <> "" Then
		Response.Write("<span class=failed>FAILED</span></p><ul>" & strError & "</ul>")
		Exit Sub
	End If
	
	If strDBType = "MSACCESS" Then
		strSql = "DELETE FROM " & strTablePrefix & "IPCountry"
	Else
		strSql = "TRUNCATE TABLE " & strTablePrefix & "IPCountry"
	End If
	
	Dim rsTruncate : Set rsTruncate = Server.CreateObject("ADODB.Recordset")
	rsTruncate.Open strSql, objConn, 1, 2, &H0000
	Set rsTruncate = Nothing
	
	strError = strError & CheckErrors(Err.Number, Err.Description)
	If strError <> "" Then
		Response.Write("<span class=failed>FAILED</span></p><ul>" & strError & "</ul>")
		Exit Sub
	End If
	
	On Error Goto 0
	
	If Not objTS.AtEndOfStream Then
	
		Dim strFirstLine : strFirstLine = objTS.ReadLine
		
		If Not CheckFileHeader("Countries", strFirstLine) Then
			Response.Write("<span class=failed>FAILED</span></p><ul><li>The countries file has does not have the correct header. Check the file and try again.</li></ul>")
			Exit Sub
		End If
		
		Dim datSerial : datSerial = Right(strFirstLine, 10)
		
		Do While Not objTS.AtEndOfStream
			
			intLine = objTS.Line
			strLine = objTS.Readline
			aryLine = Split(strLine,"||")
			If UBound(aryLine) = 2 Then
				strSql = "INSERT INTO " & strTablePrefix & "IPCountry " &_
					"(ic_ipstart, ic_ipend, ic_code) VALUES(" &_
					Trim(aryLine(0)) & ", " &_
					Trim(aryLine(1)) & ", " &_
					FormatDatabaseString(Trim(aryLine(2)), 2) & ")"
				objConn.Execute(strSql)
			Else
				strError = strError & "<li>Error on line " & intLine & ".</li>"
			End If
		Loop
		
	End If
		
	If strError <> "" Then
		Response.Write("OK. Update countries partially completed, some lines had errors:</p><ul>" & strError & "</ul>")
	Else
		Response.Write("<span class=success>SUCCESS</span></p>")
		Call UpdateConfigValue("Countries", datSerial)
	End If
	
	objTS.Close : Set objTS = Nothing
	
End Sub

Function CheckFileHeader(strFileType, strHeader)

	Dim blnResult
	
	If Left(strHeader, Len(strFileType) + 3) <> "##" & strFileType & ":" Or Not IsDate(Right(strHeader, 10)) Then
		blnResult = False
	Else
		blnResult = True
	End If
	
	CheckFileHeader = blnResult

End Function

Sub InsertConfig(strName, strValue, strGroup, intType, intOrder, strExtra)

	strSql = "INSERT INTO " & strTablePrefix & "Config (c_name, c_value, c_group, c_type, c_order, c_extra) VALUES" &_
		"(" & FormatString(strName, 255) & ", " &_
		FormatString(strValue, 255) & ", " &_
		FormatString(strGroup, 255) & ", " &_
		intType & ", " &_
		intOrder & ", " &_
		FormatString(strExtra, 255) & ")"
	objConn.Execute(strSql)

End Sub

Sub UpdateConfigValue(strName, strValue)
	
	strSql = "UPDATE " & strTablePrefix & "Config " &_
		"SET c_value = " & FormatDatabaseString(strValue, 255) & " " &_
		"WHERE c_name = " & FormatDatabaseString(strName, 255)
	
	Dim rsUpdate : Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
	
	rsUpdate.Open strSql, objConn, 1, 2, &H0000

	Set rsUpdate = Nothing
	
End Sub

Sub UpdateConfigOrder(strName, intOrder)
	
	strSql = "UPDATE " & strTablePrefix & "Config " &_
		"SET c_order = " & intOrder & " " &_
		"WHERE c_name = " & FormatDatabaseString(strName, 255)
	
	Dim rsUpdate : Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
	
	rsUpdate.Open strSql, objConn, 1, 2, &H0000

	Set rsUpdate = Nothing
	
End Sub

Sub UpgradeData(intUpgradeType)

	Dim datBegin
	
	If intUpgradeType = 2 Then
		datBegin = FormatDatabaseDate(DateAdd("d", -7, Date()), strDB2Type)
	ElseIf intUpgradeType = 3 Then
		datBegin = FormatDatabaseDate(DateAdd("m", -1, Date()), strDB2Type)
	ElseIf intUpgradeType = 4 Then
		datBegin = FormatDatabaseDate(DateAdd("m", -3, Date()), strDB2Type)
	ElseIf intUpgradeType = 5 Then
		datBegin = FormatDatabaseDate(DateAdd("m", -6, Date()), strDB2Type)
	ElseIf intUpgradeType = 6 Then
		datBegin = FormatDatabaseDate(DateAdd("m", -12, Date()), strDB2Type)
	End If
	
	strSql = "SELECT COUNT(*) FROM " & strTable2Prefix & "PageLog "

	If datBegin <> "" Then
		strSql = strSql & "WHERE pl_datetime > " & datBegin
	End If

	Dim rsCount : Set rsCount = objConn2.Execute(strSql)
	Dim intRecords : intRecords = rsCount(0)
	rsCount.Close : Set rsCount = Nothing
	
	With Response
		.Write("<table border=0 cellpadding=5 cellspacing=0 class=text>")
		.Write("<form name=progress>")
		.Write("<tr><td width=20>&nbsp;</td><td><strong>Total Records:</strong>&nbsp;</td>")
		.Write("<td><input type=text value=""" & intRecords & """ name=records size=12></td></tr>")
		.Write("<tr><td width=20>&nbsp;</td><td><strong>Completed Records:</strong>&nbsp;</td>")
		.Write("<td><input type=text value=""0"" name=counter size=12></td></tr>")
		.Write("<tr><td width=20>&nbsp;</td><td><strong>Percent Complete:</strong>&nbsp;</td>")
		.Write("<td><input type=text value=""0.0"" name=percent size=8></td></tr>")
		.Write("</form></table>")
		.Flush
	End With
		
	
	Dim intCounter : intCounter = 0

	Dim rsDefinitions : Set rsDefinitions = CreateObject("ADODB.Recordset")
	
	strSql = "SELECT d_id, d_name, d_regexp, d_extra, d_type " &_
		"FROM " & strTablePrefix & "Definitions " &_
		"ORDER BY d_id ASC"
	
	rsDefinitions.Open strSql, objConn, 0, 1, &H0001
	
	strSql = "SELECT pl_datetime, pl_scriptname, pl_scripturl, pl_referrer, pl_referrerurl, pl_referrerhost,  " &_
		"pl_referrerdomain, pl_referrerextension, pl_keywords, pl_sessionid, pl_useragent, " &_
		"pl_ipaddress, pl_remotehost, pl_browser, pl_browsertype, pl_screenarea, pl_os, pl_language " &_
		"FROM " & strTable2Prefix & "PageLog "
		
	If datBegin <> "" Then
		strSql = strSql & "WHERE pl_datetime > " & datBegin & " "
	End If
		
	strSql = strSql & "ORDER BY pl_datetime ASC"
		
	Dim rsPageLog : Set rsPageLog = objConn2.Execute(strSql)
	
	Do While Not rsPageLog.Eof
	
		intCounter = intCounter + 1
	
		Dim datHit : datHit	= rsPageLog(0)
		Dim strScriptName : strScriptName = rsPageLog(1)
		Dim strScriptUrl : strScriptUrl	= rsPageLog(2)
		Dim strReferrer : strReferrer = rsPageLog(3)
		Dim strReferrerPage : strReferrerPage = rsPageLog(4)
		Dim strReferrerHost : strReferrerHost = rsPageLog(5)
		Dim strReferrerDomain : strReferrerDomain = rsPageLog(6)
		Dim strReferrerExtension : strReferrerExtension = rsPageLog(7)
		Dim strKeywords : strKeywords = rsPageLog(8)
		Dim strSession : strSession = rsPageLog(9)
		Dim strUserAgent : strUseragent = rsPageLog(10)
		Dim strIPAddress : strIPAddress = rsPageLog(11)
		Dim strHost : strHost = rsPageLog(12)
		Dim strBrowser : strBrowser = rsPageLog(13)
		Dim strBrowserType : strBrowserType = rsPageLog(14)
		Dim strResolution : strResolution = rsPageLog(15)
		Dim strOS : strOS = rsPageLog(16)
		Dim strLanguage : strLanguage = rsPageLog(17)

		If strHost = strIPAddress Then
			strHost = ""
		End If
		
		Dim intSessionID : If Len(strIPAddress) > 0 Then
			intSessionID = Left(strSession, Len(strSession) - Len(Replace(strIPAddress, ".", "")))
		Else
			intSessionID = strSession
		End if

		Dim intIPNumber : intIPNumber = ConvertIPAddressToLong(strIPAddress)
		Dim strPath : strPath = ExtractPath(strScriptName)
		Dim strExtension : strExtension = ExtractFileType(strScriptName)

		strSql = "SELECT pn_id, pn_url, pn_page, pn_path, pn_extension " &_
			"FROM " & strTablePrefix & "PageNames " &_
			"WHERE pn_url = " & FormatDatabaseString(strScriptUrl, 255)
			
		Dim rsUrl : Set rsUrl = Server.CreateObject("ADODB.Recordset")
		
		If strDBType = "MYSQL" Then
			rsUrl.CursorLocation = 3
		End If
		
		rsUrl.Open strSql, objConn, 1, 2, &H0001

		If rsUrl.Eof Then
			rsUrl.AddNew
			rsUrl(1) = ProtectInsert(strScriptUrl, 255)
			rsUrl(2) = ProtectInsert(strScriptName, 255)
			rsUrl(3) = ProtectInsert(strPath, 255)
			rsUrl(4) = ProtectInsert(strExtension, 10)
			rsUrl.Update
		End If
		Dim intPage : intPage = rsUrl("pn_id")

		rsUrl.Close : Set rsUrl = Nothing
		
		Dim intUserAgent
		
		If strBrowserType <> "Robot" Then
			
			strSql = "SELECT s_id, s_ip, s_hostname, s_useragent, s_browser, " &_
				"s_os, s_language, s_country, s_screenarea " &_
				"FROM " & strTablePrefix & "Sessions " &_
				"WHERE s_id = " & intSessionID

			Dim rsSession : Set rsSession = Server.CreateObject("ADODB.Recordset")
			rsSession.Open strSql, objConn, 1, 2, &H0001
			
			If rsSession.Eof Then
				
				Dim strCountry
				
				If CheckPrivateIP(strIPAddress) = True Then
					strCountry = "00"
				Else
					strCountry = GetCountry(intIPNumber)
				End If
				
				strLanguage = CleanLanguage(strLanguage)
				intUserAgent = CheckName(2, strUserAgent)
				
				Dim intHost : intHost = CheckName(1, strHost)
				Dim intResolution : intResolution = CheckName(3, strResolution)
				Dim intBrowser : intBrowser = CheckName(4, strBrowser)
				Dim intOs : intOs = CheckName(5, strOs)

				rsSession.Addnew
				rsSession(0) = intSessionID
				rsSession(1) = intIPNumber
				rsSession(2) = intHost
				rsSession(3) = intUserAgent
				rsSession(4) = intBrowser
				rsSession(5) = intOs
				rsSession(6) = ProtectInsert(strLanguage, 5)
				rsSession(7) = ProtectInsert(strCountry, 2)
				rsSession(8) = intResolution
				rsSession.Update
			End If
			
			rsSession.Close : Set rsSession = Nothing
			
			Dim intReferrer : intReferrer = 0
			
			If strReferrer <> "" And strReferrerPage = "" And strReferrerDomain = "" Then
				strReferrer = ""
			End If
			
			If strReferrer <> "" Then
				
				strSql = "SELECT r_id, r_url, r_rn_id, r_k_id " &_
					"FROM " & strTablePrefix & "Referrers " &_
					"WHERE r_url = " & FormatDatabaseString(strReferrer, 255)

				Dim rsReferrer : Set rsReferrer = Server.CreateObject("ADODB.Recordset")
				
				If strDBType = "MYSQL" Then
					rsReferrer.CursorLocation = 3
				End If
				
				rsReferrer.Open strSql, objConn, 1, 2, &H0001
				
				If rsReferrer.Eof Then
				
					strSql = "SELECT rn_id, rn_page, rn_host, rn_domain, rn_extension " &_
						"FROM " & strTablePrefix & "ReferrerNames " &_
						"WHERE rn_page = " & FormatDatabaseString(strReferrerPage, 255)

					Dim rsReferrerName : Set rsReferrerName = Server.CreateObject("ADODB.Recordset")
					
					If strDBType = "MYSQL" Then
						rsReferrerName.CursorLocation = 3
					End If
					
					rsReferrerName.Open strSql, objConn, 1, 2, &H0001
					
					If rsReferrerName.Eof Then
						rsReferrerName.AddNew
						rsReferrerName(1) = ProtectInsert(strReferrerPage, 255)
						rsReferrerName(2) = ProtectInsert(strReferrerHost, 255)
						rsReferrerName(3) = ProtectInsert(strReferrerDomain, 100)
						rsReferrerName(4) = ProtectInsert(strReferrerExtension, 10)
						rsReferrerName.Update
					End If
					
					Dim intReferrerName : intReferrerName = rsReferrerName(0)
					
					rsReferrerName.Close : Set rsReferrerName = Nothing
					
					Dim intKeywords : intKeywords = 0
					
					If strKeywords <> "" Then
						
						Dim strSite : strSite = MatchDefinition(rsDefinitions, strReferrer, 4)
						Dim intSite : intSite = CheckName(7, strSite)
						
						strSql = "SELECT k_id, k_value, k_site " &_
							"FROM " & strTablePrefix & "Keywords " &_
							"WHERE k_value = " & FormatDatabaseString(strKeywords, 255) & " " &_
							"AND k_site = " & intSite

						Dim rsKeywords : Set rsKeywords = Server.CreateObject("ADODB.Recordset")

						If strDBType = "MYSQL" Then
							rsKeywords.CursorLocation = 3
						End If

						rsKeywords.Open strSql, objConn, 1, 2, &H0001

						If rsKeywords.Eof Then
							rsKeywords.AddNew
							rsKeywords(1) = ProtectInsert(strKeywords, 255)
							rsKeywords(2) = intSite
							rsKeywords.Update
						End If
						intKeywords = rsKeywords("k_id")
						rsKeywords.Close : Set rsKeywords = Nothing
						
					End If
				
					rsReferrer.Addnew
					rsReferrer(1) = ProtectInsert(strReferrer, 255)
					rsReferrer(2) = intReferrerName
					rsReferrer(3) = intKeywords
					rsReferrer.Update
				End If
				intReferrer = rsReferrer(0)
				
				rsReferrer.Close : Set rsReferrer = Nothing

			End If

			strSql = "INSERT INTO " & strTablePrefix & "PageLog (pl_datetime, pl_pn_id, pl_r_id, pl_s_id) VALUES(" &_
				FormatDatabaseDate(datHit, strDBType) & ", " &_
				intPage & ", " &_
				intReferrer & ", " &_
				intSessionID & ")"

			objConn.Execute(strSql)
		
		Else

			intUserAgent = CheckName(2, strUserAgent)
			Dim intRobot : intRobot = CheckName(6, strBrowser)

			strSql = "INSERT INTO " & strTablePrefix & "RobotLog (rl_datetime, rl_pn_id, rl_useragent, rl_robot, rl_ip) VALUES(" &_
				FormatDatabaseDate(datHit, strDBType) & ", " &_
				intPage & ", " &_
				intUserAgent & ", " &_
				intRobot & ", " &_
				intIPNumber & ")"

			objConn.Execute(strSql)
		
		End If
		
		If (intCounter Mod 100 = 0 Or intCounter = intRecords) Then
			Response.Write("<script language=JavaScript>" & vbcrlf)
			Response.Write("document.progress.counter.value='" & intCounter & "';" & vbcrlf)
			Response.Write("document.progress.percent.value='" & FormatPercent(intCounter / intRecords, 2) & "';" & vbcrlf)
			Response.Write("</script>" & vbcrlf)
			Response.Flush
		End If
		
		rsPageLog.Movenext
	Loop
	rsDefinitions.Close : Set rsDefinitions = Nothing
	rsPageLog.Close : Set rsPagelog = Nothing

End Sub

Function FormatDatabaseDate(datDate, strType)

	Dim datDateTemp, datTimeTemp, strDateFormat, strTimeFormat
	Dim datTemp, strSeparator, datDatabaseDate, datDatabaseTime, datFull

	If strType = "MSSQL" Then
		strDateFormat = "YYYYMMDD"
	Else
		strDateFormat = "YYYY-MM-DD"
	End If
	
	strTimeFormat = "HH:MM:SS"
	
	datDateTemp = UCase(strDateFormat)
	datTimeTemp = UCase(strTimeFormat)

	datDateTemp = Replace(datDateTemp, "DD", FormatDatePart(Day(datDate)))
	datDateTemp = Replace(datDateTemp, "MMMM", MonthName(Month(datDate), False))
	datDateTemp = Replace(datDateTemp, "MMM", MonthName(Month(datDate), True))
	datDateTemp = Replace(datDateTemp, "MM", FormatDatePart(Month(datDate)))
	datDateTemp = Replace(datDateTemp, "YYYY", Year(datDate))
	datDateTemp = Replace(datDateTemp, "YY", Right(Year(datDate), 2))
	
	datTimeTemp = Replace(datTimeTemp, "HH", FormatDatePart(DatePart("h", datDate)))
	datTimeTemp = Replace(datTimeTemp, "MM", FormatDatePart(DatePart("n", datDate)))
	datTimeTemp = Replace(datTimeTemp, "SS", FormatDatePart(DatePart("s", datDate)))
	
	If strType = "MSACCESS" Then
		strSeparator = "#"
	Else
		strSeparator = "'"
	End If

	datTemp = strSeparator & datDateTemp & " " & datTimeTemp & strSeparator

	FormatDatabaseDate = datTemp

End Function

Function FormatDatePart(datPart)
	Dim datTemp
	
		If Len(datPart) = 1 Then
			datTemp = "0" & datPart
		Else
			datTemp = datPart
		End If

	FormatDatePart = datTemp
End Function

Function FormatDatabaseString(strString, intLength)

	Dim strTemp
	
	If strDBType = "MYSQL" Then 
		strTemp = "'" & Replace(Replace(Left(strString, intLength), "\", "\\"), "'", "''")  & "'"
	Else
		strTemp = "'" & Replace(Left(strString, intLength), "'", "''") & "'"
	End If

	FormatDatabaseString = strTemp

End Function

Function ConvertIPAddressToLong(strIPAddress)

	Dim strTemp : strTemp = strIPAddress
	Dim aryIP : aryIP = Split(strTemp, ".")
	Dim intNumber : intNumber = (CInt(aryIP(0)) * 16777216) + (CInt(aryIP(1)) * 65536) + (CInt(aryIP(2)) * 256) + CInt(aryIP(3))

	intNumber = intNumber - 2147483647
	
	ConvertIPAddressToLong = intNumber

End Function

Function ExtractPath(strScriptName)

	Dim strTemp : strTemp = Left(strScriptName, InStrRev(strScriptName, "/"))

	ExtractPath = strTemp

End Function

Function ExtractFileType(strScriptName)

	Dim strTemp
	If InstrRev(strScriptName, ".") > 0 Then
		strTemp = Mid(strScriptName, InStrRev(strScriptName, ".") + 1)
	Else
		strTemp = ""
	End If

	ExtractFileType = strTemp

End Function

Function GetCountry(intIPNumber)

	Dim strValue

	If Not IsNumeric(intIPNumber) Then
		strValue = ""
	Else
		strSql = "SELECT ic_code FROM " & strTablePrefix & "IPCountry " &_
			"WHERE " & intIPNumber & " BETWEEN ic_ipstart and ic_ipend"

		Dim rsCountry : Set rsCountry = Server.CreateObject("ADODB.Recordset")
		rsCountry.Open strSql, objConn, 1, 2, 1

		If Not rsCountry.Eof Then
			strValue = rsCountry(0)
		Else
			strValue = ""
		End If

		rsCountry.Close
		Set rsCountry = Nothing
	End If

	GetCountry = strValue

End Function

Function CheckName(intType, strName)

	Dim intValue

	If strName = "" Then
		intValue = 0
	Else
		strSql = "SELECT n_id, n_value, n_type FROM " & strTablePrefix & "Names WHERE n_value = " & FormatDatabaseString(strName, 255)

		Dim rsName : Set rsName = Server.CreateObject("ADODB.Recordset")
		
		If strDBType = "MYSQL" Then
			rsName.CursorLocation = 3
		End If
		
		rsName.Open strSql, objConn, 1, 2, &H0001

		If rsName.Eof Then
			rsName.AddNew
			rsName("n_value")	= ProtectInsert(strName, 255)
			rsName("n_type")	= intType
			rsName.Update
		End If
		intValue = rsName("n_id")

		rsName.Close
		Set rsName = Nothing
		
	End If
	
	CheckName = intValue

End Function

Public Function ExtractHost(strReferrer)
	
	Dim strTemp : strTemp = strReferrer
	
	strTemp = Replace(strTemp, "http://", "")
	strTemp = Replace(strTemp, "https://", "")

	If InStr(strTemp, "/") > 0 Then
		strTemp = Mid(strTemp, 1, InStr(strTemp, "/") - 1)
	End If
	
	ExtractHost = strTemp

End Function

Public Function ExtractDomain(strHost)
		
	Dim strDomain, strExtension
	
	Dim strTemp : strTemp = strHost
	
	If InStr(strTemp, ".") > 0 Then
	
		Dim strEnd : strEnd = Mid(strTemp, InStrRev(strTemp, "."))
	
		If InStr(".com.net.org.edu.gov.mil.int.aero.biz.coop.info.museum.name.pro", strEnd) > 0 Then
			strExtension = strEnd
		Else
			If Len(strEnd) = 3 And Not IsNumeric(Right(strEnd, 2)) Then 
				Dim strRemainder : strRemainder = Left(strTemp, Len(strTemp) - Len(strEnd))
				Dim strPart : strPart = Right(strRemainder, Len(strRemainder) - InStrRev(strRemainder, ".") + 1)
				Dim strGeneric : strGeneric = ".ac.com.co.edu.go.gv.gov.govt.int.ltd.mi.mil.net.or.org.plc"

				Select Case strEnd
				Case ".ca"
					strExtension = CheckExtension(".ab.bc.mb.nb.nf.ns.nt.nu.on.pe.qc.sk.yk", strPart, strEnd)
				Case Else
					strExtension = CheckExtension(strGeneric, strPart, strEnd)
				End Select
				
				If strExtension = "" Then
					strExtension = strEnd
				End If
				
			End If
		End If
		
	End If

	If strExtension <> "" Then
	
		Dim objSearch : Set objSearch	= New RegExp
		
		Dim strPattern : strPattern = "[\w|\-]+" & Replace(strExtension, ".", "\.") & "$"
		
		With objSearch
			.Pattern 	= strPattern
			.IgnoreCase = True
			.Global 	= False
		End With

		Dim objResults : Set objResults = objSearch.Execute(strTemp)

		If objResults.Count > 0 Then
			Dim colItem
			For Each colItem In objResults
				strDomain = colItem.Value
				Exit For
			Next
		End If
		
		Set objSearch = Nothing : Set objResults = Nothing
	Else
		strDomain = ""
	End If

	ExtractDomain = strDomain

End Function

Function CheckExtension(strCompare, strPart, strEnd)

	Dim strTemp

	If InStr(strCompare, strPart) > 0 Then
		strTemp = strPart & strEnd
	End If
	
	CheckExtension = strTemp

End Function

Public Function ExtractExtension(strDomain)

	Dim strTemp : strTemp = strDomain
	
	If strDomain <> "" Then
		strTemp = Mid(strTemp, InStr(strTemp, "."))
	Else
		strTemp = ""
	End If

	ExtractExtension = strTemp

End Function

Function CleanLanguage(strLanguage)

	Dim strTemp : strTemp = strLanguage
	
	If strTemp <> "" Then
		If InStr(strTemp, ",") > 0 Then
			strTemp = Trim(Left(strTemp, InStr(strTemp, ",") - 1))
		Else
			strTemp = Trim(strTemp)
		End If
		If InStr(strTemp, ";") > 0 Then
			strTemp = Trim(Left(strTemp, InStr(strTemp, ";") - 1))
		End If
	End If

	CleanLanguage = strTemp

End Function

Function MatchDefinition(rsDefinition, strCompare, intType)

	Dim strMatch

	rsDefinition.Filter = "d_type = " & intType

	Do While Not rsDefinition.Eof
	
		Dim objSearch : Set objSearch = New RegExp
		With objSearch
			.Pattern 	= rsDefinition(2)
			.IgnoreCase = True
			.Global 	= False
		End With
		
		If objSearch.Test(strCompare) = True Then
			strMatch = rsDefinition(1)
			Exit Do
		End If
		
		Set objSearch = Nothing
		rsDefinition.Movenext
	Loop
	
	MatchDefinition = strMatch

End Function

Function CheckPrivateIP(strIPAddress)

	Dim blnCheck : blnCheck = False
	
	If Left(strIPAddress, 3) = "10." Then
		blnCheck = True
	ElseIf Left(strIPAddress, 7) = "192.168" Then
		blnCheck = True
	ElseIf Left(strIPAddress, 4) = "172." Then
		Dim aryIP : aryIP = Split(strIPAddress, ".")
		If UBound(aryIP) = 3 Then
			If CInt(aryIP(1)) => 16 And CInt(aryIP(1)) =< 31 Then
				blnCheck = True
			End If
		End If
	End If

	CheckPrivateIP = blnCheck

End Function

Function ProtectInsert(strValue, intLength)

	ProtectInsert = Left(strValue, intLength)

End Function

Function FormatString(strValue, intLength)

	FormatString = "'" & Replace(Left(strValue, intLength), "'", "''") & "'"

End Function

%>
