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
<!--#Include File="clsReport.asp"-->
<%
Server.ScriptTimeout = 900

Dim strScriptName, datScriptStart, datScriptEnd, datExecutionTime

Dim datStart : datStart = Request.Querystring("sd")
Dim datEnd : datEnd = Request.Querystring("ed")
Dim intReport : intReport = Request.Querystring("r")
Dim intDetail : intDetail = Request.Querystring("d")
Dim strExtra : strExtra = Request.Querystring("e")
Dim intExtraType : intExtraType = Request.Querystring("et")
Dim strExtra2 : strExtra2 = Request.Querystring("e2")
Dim intExtraType2 : intExtraType2 = Request.Querystring("et2")
Dim intItems : intItems = Request.Querystring("i")
Dim intGroup : intGroup = Request.Querystring("g")

' SET REPORTING DEFAULTS
If datStart = "" Then
	datStart = Request.Cookies("report")("start")
	If datStart = "" Then
		datStart = FormatDateTime(DateAdd("h", aryMTConfig(intMTTimeOffset), Now()), 2)
	End If
Else
	Response.Cookies("report")("start") = datStart
End If

If datEnd = "" Then
	datEnd = Request.Cookies("report")("end")
	If datEnd = "" Then
		datEnd = FormatDateTime(DateAdd("h", aryMTConfig(intMTTimeOffset), Now()), 2)
	End If
Else
	Response.Cookies("report")("end") = datEnd
End If

If intReport = "" Then
	intReport = 1
	intGroup = 1
End If

If intItems = "" Then
	intItems = Request.Cookies("report")("items")
	If intItems = "" Then
		intItems = 100
	End If
Else
	Response.Cookies("report")("items") = intItems
End If

If intExtraType = "" Then
	intExtraType = 0
End If

Dim objReport : Set objReport = New MTReport
With ObjReport
	.Database			= aryMTDB
	.Config				= aryMTConfig
	.Report				= intReport
	.StartDate			= datStart
	.EndDate			= datEnd
	.Items				= intItems
	.Detail				= intDetail
	.Extra				= strExtra
	.ExtraType			= intExtraType
	.Extra2				= strExtra2
	.ExtraType2			= intExtraType2
End With

Call CreateDatabaseConnection(1)
Dim blnAdmin : blnAdmin = Authenticate(False, aryMTDB(5))

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Site Statistics</title>
	<link rel="stylesheet" href="style.css" type="text/css">
	<script language="JavaScript" src="javascript.js" type="text/javascript"></script>
	<script language="JavaScript">
	<% Call objReport.GenerateReportJS %>
	<% Call objReport.GeneratePresetDatesJS() %>
	</script>
	
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
    <td colspan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width=61><img src="images/subnav_pointer.gif" width="61" height="22"></td>
		<td width=258 background="images/subnav_scale.gif" valign=middle>
		<span class=sitename><% Response.Write objReport.SiteName %></span></td>
        <td background="images/subnav_scale.gif" valign=middle align=right>
		<span class=version><% Response.Write objReport.Version %></span></td>
		</td>
      </tr>
    </table>
  </td>
</tr>
<tr valign=top height="100%">
	<td id="chooser" style="padding: 5px;" width=160>
	<form name=report>
		<table border=0 cellpadding=0 cellspacing=0 class=select width=160>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0 width="100%">
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td width="20"><img src="images/grey_arrow.gif"></td>
							<td class=header>Report Period</td>
						</tr>
						</table>					</td>
				</tr>
				<tr>
					<td class=chooser>
						<table cellpadding=3 cellspacing=0 border=0 width="100%">
						<tr>
							<td colspan=2 align=center>
							<select name=preset onChange="presetdate();">
								<option value="CUSTOM">Custom</option>
								<option value="TODAY">Today</option>
								<option value="YESTERDAY">Yesterday</option>
								<option value="LASTWEEKROLL">Last 7 Days</option>
								<option value="THISMONTH">Current Month</option>
								<option value="LASTMONTH">Last Month</option>
								<option value="LASTMONTHROLL">Last Month (Rolling)</option>
							</select>							</td>
						</tr>
						<tr>
							<td align=right><span class=chooser>Start: </span></td>
							<td><% Call objReport.DisplayDateChooser(datStart, "start") %></td>
						</tr>
						<tr>
							<td align=right><span class=chooser>End: </span></td>
							<td><% Call objReport.DisplayDateChooser(datEnd, "end") %></td>
						</tr>
						</table>					</td>
				</tr>
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td width="20"><img src="images/grey_arrow.gif"></td>
							<td class=header>Report Type</td>
						</tr>
						</table>					</td>
				</tr>
				<tr>
					<td class=chooser>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<% Call objReport.DisplayReportList(intReport) %>
						</table>					</td>
				</tr>
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td width="20"><img src="images/grey_arrow.gif"></td>
							<td class=header>Options</td>
						</tr>
						</table>					</td>
				</tr>
				<tr>
					<td class=chooser style="padding: 5px;" align=center>
						<span class=chooser>Show&nbsp;</span>
						<% Call objReport.DisplayItemsChooser(intItems) %>					</td>
				</tr>
				<tr>
					<td class=chooser style="padding: 5px;" align=center>
					<a href="javascript:GenerateReport(<%=intReport%>)">
					<img src="images/generate_btn.gif" border=0 vpspace=10></a>					</td>
				</tr>
				</table>			</td>
		</tr>
		</table>
	</form>	</td>
	<td align=left style="padding: 5px;" width="100%">
	<%
	' GENERATE REPORT AND CALCULATE EXECUTION TIME
	datScriptStart = Now()
	Call objReport.GenerateReport
	datScriptEnd = Now()
	datExecutionTime = DateDiff("s", datScriptStart, datScriptEnd)
	%>
	<br>
	<img src="images/spacer.gif" width=300 height=1></td>
</tr>
<tr class=pgfooter id=pgfooter>
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
Set objReport = Nothing
Call CloseDatabaseConnection() 
%>
