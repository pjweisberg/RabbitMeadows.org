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
<!--#Include File="clsConfig.asp"-->
<!--#Include File="clsUpload.asp"-->
<%
Server.ScriptTimeout = 9000

Dim intSetting : intSetting = Request.Querystring("s")

Dim objConfig : Set objConfig = New MTConfig
With objConfig
	.Database		= aryMTDB
	.Setting		= intSetting
	.Config			= aryMTConfig
End With

Call CreateDatabaseConnection(1)
Dim blnAdmin : blnAdmin = Authenticate(True, aryMTDB(5))
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Site Statistics</title>
	<link rel="stylesheet" href="style.css" type="text/css">
	<script language="JavaScript" src="javascript.js" type="text/javascript"></script>
</head>

<body>
<table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
  <tr id="header" class=pgheader>
  	<td colspan="2" height="100">
	  <table height="100" width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
    		<td rowspan="3" width="267"><img src="images/metatraffic_logo.gif" width="267" height="44"></td>
   		  <td colspan="7"></td>
   		</tr>
			<tr>
				<td width="81"><a href="default.asp" onMouseOver="MM_swapImage('Image1','','images/reports_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/reports.gif" name="Image1" width="81" height="21" border="0" id="Image1"></a></td>
				<td width="1"><img src="images/spacer.gif" height="" width="1" border="0"></td>
				<% If blnAdmin = True Then %><td width="81"><a href="settings.asp" onMouseOver="MM_swapImage('Image2','','images/settings_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/settings.gif" name="Image2" width="81" height="21" border="0" id="Image2"></a></td><% End If %>
				<td width="1"><img src="images/spacer.gif" height="" width="1" border="0"></td>
				<td width="81"><a href="tracking.asp" onMouseOver="MM_swapImage('Image3','','images/tracking_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/tracking.gif" name="Image3" width="81" height="21" border="0" id="Image3"></a></td>
				<td width="1"><img src="images/spacer.gif" height="" width="1" border="0"></td>
				<td width="81"><a href="login.asp?action=logout" onMouseOver="MM_swapImage('Image4','','images/logout_ovr.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/logout.gif" name="Image4" width="81" height="21" border="0" id="Image4"></a></td>
				<td width="100%"></td>
	  </tr>
		<tr>
			<td colspan="7" height="20"></td>
		</tr>
      </table></td>
  </tr>
  <tr>
    <td colspan=2>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="61"><img src="images/subnav_pointer.gif" width="61" height="22"></td>
        <td width=258 background="images/subnav_scale.gif" valign=middle>
		<span class=sitename><% Response.Write objConfig.SiteName %></span></td>
        <td background="images/subnav_scale.gif" valign=middle align=right>
		<span class=version><% Response.Write objConfig.Version %></span></td>
		</td>
      </tr>
    </table>
  </td>
</tr>
<tr valign=top height="100%">
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
							<td class=header>Options</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class=chooser>
						<table cellpadding=3 cellspacing=0 border=0>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=0" class=chtitle>General</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=1" class=chtitle>Configuration</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=2" class=chtitle>Users</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=5" class=chtitle>Campaigns</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=4" class=chtitle>Actions</a></td>
						</tr>
						<tr>
							<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
							<td><a href="settings.asp?s=3" class=chtitle>Maintenance</a></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
		</table>
	</td>
	<td align=left style="padding: 5px;" width="100%">
	<%
	' GENERATE REPORT AND CALCULATE EXECUTION TIME
	Call objConfig.GenerateSettings(intSetting)
	%>
	</td>
</tr>
<tr class=pgfooter>
	<td><img src="images/blue_scale.gif" width="2" height="23"></td>
	<td valign=middle align=right><span class=copyright>&copy; Copyright 2007, </span>
	<a href="http://www.metasun.com/" target="_new">Brinkster Site Statistics</a><span style="font-size:10px; color:#737373;">, powered by MetaTraffic</span> </td>
</tr>
<tr class=pgbottom>
	<td colspan=2 height=4>&nbsp;</td>
</tr>
</table>
</body>
</html>
<%
Set objConfig = Nothing
Call CloseDatabaseConnection() 
%>