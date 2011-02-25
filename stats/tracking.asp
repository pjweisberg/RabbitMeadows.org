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
Dim strScriptDir, strAction, blnExclude, strTarget

Dim intTracking : intTracking = Request.Querystring("t")
Dim strActionCode : strActionCode = Request.Form("action")
Dim strCampaignCode : strCampaignCode = Request.Form("campaign")

strScriptDir = Request.Servervariables("script_name")
strScriptDir = Left(strScriptDir, InStrRev(strScriptDir, "/"))

Dim objReport : Set objReport = New MTReport
With ObjReport
	.Database			= aryMTDB
	.Config				= aryMTConfig
End With

Call CreateDatabaseConnection(1)
Dim blnAdmin : blnAdmin = Authenticate(False, aryMTDB(5))

strAction = Request.Form("action")

Select Case strAction
Case "Exclude Visits"
	Response.Cookies("mt_exclude") = 1
	Response.Cookies("mt_exclude").Path = "/"
	Response.Cookies("mt_exclude").Expires = DateAdd("m", 24, Date())
Case "Include Visits"
	Response.Cookies("mt_exclude") = ""
	Response.Cookies("mt_exclude").Path = "/"
End Select

If Request.Cookies("mt_exclude") <> "" Then
	blnExclude = False
Else
	blnExclude = True
End If
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
	<td>
		<table border=0 cellpadding=0 cellspacing=0>
		<tr>
			<td id="chooser" style="padding: 5px;" width=180>
				<table border=0 cellpadding=0 cellspacing=0 class=select width=180>
				<tr>
					<td>
						<table border=0 cellpadding=0 cellspacing=0 width="100%">
						<tr>
							<td>
								<table border=0 cellpadding=0 cellspacing=0 width="100%">
								<tr>
									<td width="20"><img src="images/grey_arrow.gif"></td>
									<td class=header>Tracking</td>
								</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td class=chooser>
								<table cellpadding=3 cellspacing=0 border=0>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=0" class=chtitle>Overview</a></td>
								</tr>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=1" class=chtitle>Javascript</a></td>
								</tr>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=2" class=chtitle>Active Server Pages</a></td>
								</tr>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=4" class=chtitle>Campaigns</a></td>
								</tr>
								<tr>
									<td width=20><img src="images/lt_grey_arrow_dn.gif" border=0></td>
									<td><a href="tracking.asp?t=3" class=chtitle>Redirects</a></td>
								</tr>
								</table>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td></td>
		</tr>
		<tr>
			<td id="chooser" style="padding: 5px;" width=180>
				<table border=0 cellpadding=0 cellspacing=0 class=select width=180>
				<tr>
					<td>
					<table border=0 cellpadding=0 cellspacing=0 width="100%">
					<tr>
						<td>
							<table border=0 cellpadding=0 cellspacing=0 width="100%">
							<tr>
								<td width="20"><img src="images/grey_arrow.gif"></td>
								<td class=header>Exclude Visits</td>
							</tr>
							</table>
						</td>
					</tr>
					<tr>
						<form method=post>
						<td class=chooser style="padding:5px;">
						<% If blnExclude = True Then %>
						<p>Your visits are being tracked.</p>
						<p class=about>If you have a dynamic IP address, click the button below to set 
						a cookie that will exclude your visits from being tracked with this 
						browser.</p>
						<div align=center><input type=submit name=action value="Exclude Visits"></div>
						<% Else %>
						<p><span class=chooser>Your visits are not being tracked.</span></p>
						<p class=about>This browser is currently being excluded from tracking. Click the button 
						below to remove this.</p>
						<div align=center><input type=submit name=action value="Include Visits"></div><br>
						<% End If %>
						</td>
						</form>
					</tr>
					</table>
			</table></td>
		</tr>
		</table>
	</td>
	<td align=left style="padding: 5px;" width="100%">
		<% Select Case intTracking %>
		<% Case 1 %>
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Javascript</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<th align=left>Overview</th>
				</tr>
				<tr>
					<td><p>Javascript tracking provides a way to track any type of web page including .htm, .html, .php, .aspx, .cfm, etc. 
					Robots cannot be tracked using javascript because they will not execute the javascript code.</p>
					</td>
				</tr>
				<tr>
					<th align=left>Select Action</th>
				</tr>
				<tr>
					<td><p>Actions allow you to track special events on your web site such as a product sale, 
					member signup, etc. If you wish to track an action, make a selection from the drop down menu 
					below and click the Update button.</p>
						<table border=0 cellpadding=4 cellspacing=0 align=center>
						<form method=post>
						<tr>
						<td><% Call DisplayActionsSelect(strActionCode) %></td>
						<td><input type=image name=submit src="images/update_btn.gif"></td>
						</tr>
						</form>
						</table>
					</td>
				</tr>
				<tr>
					<th align=left>Tracking Code</th>
				</tr>
				<tr>
					<td>
						<p>Copy and paste the following code into each web page that you would like to track:</p>
						<div align=center>
						<textarea cols=100 rows=18 readonly nowrap>
&lt;script language=&quot;JavaScript&quot;&gt;
// METATRAFFIC -- COPYRIGHT (C) 2002-2005, Brinkster Site Statistics Corp.

var pagetitle = document.title; //INSERT CUSTOM PAGE NAME IN QUOTES
var action = &quot;<% = strActionCode %>&quot;; //ACTION CODE
var amount = &quot;0&quot;; //ACTION AMOUNT (LEAVE BLANK OR 0 IF NO AMOUNT)
var order = &quot;&quot;; //INSERT UNIQUE ORDER NUMBER

var scriptlocation = &quot;<% = strScriptDir %>track.asp&quot;;

var pagedata = 'mtpt=' + escape(pagetitle) + '&mtac=' + escape(action) + '&mta=' + amount + '&mto=' + escape(order) + '&mtr=' + escape(document.referrer) + '&mtt=2&mts=' + window.screen.width + 'x' + window.screen.height + '&mti=1&mtz=' + Math.random(); 
document.write ('&lt;img height=1 width=1 ');
document.write ('src=&quot;' + scriptlocation + '?' + pagedata + '&quot;&gt;');
&lt;/script&gt;
&lt;noscript&gt;&lt;a href=&quot;http://www.metasun.com/&quot; target=&quot;_blank&quot;&gt;
&lt;img src=&quot;<% = strScriptDir %>track.asp?mtt=2&mti=1&quot; alt=&quot;website analytics software&quot; border=0&gt;&lt;/a&gt;
&lt;/noscript&gt;</textarea></div>
					</td>
				</tr>
				<tr>
					<th align=left>More Information</th>
				</tr>
				<tr>
					<td>
					<p>The ideal spot to place your tracking code is at the bottom of each web page before the closing html 
					tag (&lt;/html&gt;). For the page title to be automatically inserted, you must place the tracking code 
					after the &lt;/title&gt; tag and there must be text in the title tag.</p>
					<p>Actions can be configured in the Settings section. When using actions to track events on your web site, enter an amount and order number if available 
					into the tracking code in the designated spots. Entering an amount will allow you to track the sales for each 
					action. Entering an order number ensures that duplicate actions don't occur as Site Statistics will 
					only count a distinct order number for each action. The order number should be unique and can be up to 
					100 characters long.</p>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
		
		<% Case 2 %>
		
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Active Server Pages (ASP)</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<th align=left>Overview</th>
				</tr>
				<tr>
					<td><p>Using ASP is the optimal way to track your web pages. It allows you to track robots accessing your 
					web pages which javascript will not. The only limitation with using ASP tracking is that you will not 
					collect screen area data from your visitors.</p></td>
				</tr>
				<tr>
					<th align=left>Select Action</th>
				</tr>
				<tr>
					<td><p>Actions allow you to track special events on your web site such as a product sale, 
					member signup, etc. If you wish to track an action, make a selection from the drop down menu 
					below and click the Update button. </p>
						<table border=0 cellpadding=4 cellspacing=0 align=center>
						<form method=post>
						<tr>
						<td><% Call DisplayActionsSelect(strActionCode) %></td>
						<td><input type=image name=submit src="images/update_btn.gif"></td>
						</tr>
						</form>
						</table>
					</td>
				</tr>
				<tr>
					<th align=left>Tracking Code</th>
				</tr>
				<% 
				Dim strCode
				If strAction <> "" Then
					strCode = "&lt;%" & vbcrlf &_
						"Response.Cookies(""mt"")(""pagetitle"") = """" 'INSERT PAGE NAME (OR LEAVE BLANK)" & vbcrlf &_
						"Response.Cookies(""mt"")(""action"") = """ & strAction & """ 'ACTION CODE" & vbcrlf &_
						"Response.Cookies(""mt"")(""amount"") = """" 'ACTION AMOUNT (LEAVE BLANK OR 0 IF NO AMOUNT)" & vbcrlf &_
						"Response.Cookies(""mt"")(""order"") = """" 'INSERT UNIQUE ORDER NUMBER" & vbcrlf &_
						"Server.Execute(&quot;" & strScriptDir & "track.asp&quot;)" & vbcrlf &_
						"%&gt;"
				Else
					strCode = "&lt;%" & vbcrlf &_ 
						"Response.Cookies(""mt"")(""pagetitle"") = """" 'INSERT PAGE NAME (OR LEAVE BLANK)" & vbcrlf &_
						"Server.Execute(&quot;" & strScriptDir & "track.asp&quot;)" & vbcrlf &_
						"%&gt;"
				End If
				%>
				<tr>
					<td>
						<p>To track your .ASP files, add the following code to any .ASP page that you want to track:</p>
						<div align=center><textarea cols=110 rows=7 readonly><% = strCode %></textarea></div>
					</td>
				</tr>
				<% 
				If strAction <> "" Then
					strCode = "&lt;%" & vbcrlf &_ 
						"Response.Cookies(""mt"")(""pagetitle"") = """" 'INSERT PAGE NAME (OR LEAVE BLANK)" & vbcrlf &_
						"Response.Cookies(""mt"")(""action"") = """ & strAction & """ 'ACTION CODE" & vbcrlf &_
						"Response.Cookies(""mt"")(""amount"") = """" 'ACTION AMOUNT (LEAVE BLANK OR 0 IF NO AMOUNT)" & vbcrlf &_
						"Response.Cookies(""mt"")(""order"") = """" 'INSERT UNIQUE ORDER NUMBER" & vbcrlf &_
						"%&gt;" & vbcrlf & _
						"&lt;!-- #Include Virtual=&quot;" & strScriptDir & "track.asp&quot; --&gt;"
				Else
					strCode = "&lt;%" & vbcrlf &_ 
						"Response.Cookies(""mt"")(""pagetitle"") = """" 'INSERT PAGE NAME (OR LEAVE BLANK)" & vbcrlf &_
						"%&gt;" & vbcrlf & _
						"&lt;!-- #Include Virtual=&quot;" & strScriptDir & "track.asp&quot; --&gt;"
				End If
				%>
				<tr>
					<td>
						<p>You can also use a standard include in your asp pages like this:</p>
						<div align=center><textarea cols=110 rows=7 readonly><% = strCode %></textarea></div>
					</td>
				</tr>
				<tr>
					<th align=left>More Information</th>
				</tr>
				<tr>
					<td>
					<p>The ideal spot to place your tracking code is at the bottom of each web page.
					<p>Actions can be configured in the Settings section. When using actions to track events on your web site, enter an amount and order number if available 
					into the tracking code in the designated spots. Entering an amount will allow you to track the sales for each 
					action. Entering an order number ensures that duplicate actions don't occur as Site Statistics will 
					only count a distinct order number for each action. The order number should be unique and can be up to 
					100 characters long.</p>
					</td>
				</tr>	
				</table>
			</td>
		</tr>
		</table>
		
		<% Case 3 %>
		
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Redirects</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<td><p>You can track files that aren't web pages by linking to them in your web 
					page with HTML. This is useful for tracking downloads, media files, etc.</p></td>
				</tr>
				<tr>
					<td>
						<p>Here is an example of how to link to a file and track it:</p>
						<div align=center><textarea cols=80 rows=3 readonly>&lt;a href="<% = strScriptDir %>track.asp?mtr=/downloads/somefile.zip"&gt;Download somefile.zip&lt;/a&gt;</textarea></div>
					</td>
				</tr>				
				</table>
			</td>
		</tr>
		</table>
		
		<% Case 4 %>
		
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Campaigns</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<th align=left>Overview</th>
				</tr>
				<tr>
					<td><p>This will help you to generate links to track campaigns you have setup in the 
					Settings section. Select the campaign you want from the drop down menu below and press 
					the update button. </p>
						<table border=0 cellpadding=4 cellspacing=0 align=center>
						<form method=post>
						<tr>
						<td><% Call DisplayCampaignsSelect(strCampaignCode) %></td>
						<td><input type=image name=submit src="images/update_btn.gif"></td>
						</tr>
						</form>
						</table>
					</td>
				</tr>
				<% If strCampaignCode <> "" Then %>
				<tr>
					<th align=left>Campaign Tracking Links</th>
				</tr>
				<tr>
					<td>
						<p>Here are some examples of how to link to your site and track a campaign:</p>
						<p>Linking to your home page: </p>
						<div align=center><textarea cols=100 rows=3 readonly><% = aryMTConfig(1) %>/?<% = aryMTConfig(6) %>=<% = strCampaignCode %></textarea></div>
						<p>Linking to a specific page on your web site: </p>
						<div align=center><textarea cols=100 rows=3 readonly><% = aryMTConfig(1) %>/somedirectory/somepage.html?<% = aryMTConfig(6) %>=<% = strCampaignCode %></textarea></div>
						<p>In the example above, replace the "somedirectory" and "somepage.html" with the actual directory and filename of where you want to send the visitor.</p>
					</td>
				</tr>
				<% End If %>			
				</table>
			</td>
		</tr>
		</table>
		
		<% Case Else '0 %>
		<table border=0 cellpadding=3 cellspacing=0>
		<tr>
			<td>
				<table border=0 cellpadding=0 cellspacing=0>
				<tr>
					<td width=22><img src="images/white_arrow.gif"></td>
					<td width="100%"><span class=name>Overview</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding=5 cellspacing=0 class=settings>
				<tr>
					<td><p class=about>To track activity on your web site, you will need to insert 
					tracking code on each web page. There are three different types of tracking code:</p> 
					<ul>
						<li><a href="?t=1">Javascript</a> - Track any web page</li>
						<li><a href="?t=2">ASP (Active Server Pages)</a> - Track ASP pages and robots</li>
						<li><a href="?t=4">Campaigns</a> - Track your ad campaigns</li>
						<li><a href="?t=3">Redirects</a> - Track downloads or multimedia files</li>
					</ul>
					<p class=about>To setup web site tracking, follow these steps:</p>
						<ol>
							<li>Choose a tracking method</li>
							<li>Copy the code and paste it into your web pages</li>
						</ol>
					</td>
				</tr>		
				</table>
			</td>
		</tr>
		</table>
		<% End Select %>
	</td>
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

Sub DisplayActionsSelect(strCode)
	
	Dim strSelected
	
	Dim strSql : strSql = "SELECT a_code, a_name " &_
		"FROM " & aryMTDB(5) & "Actions "
	If aryMTDB(0) = "MSSQL" Then
		strSql = strSql & "WHERE a_display = 1"
	Else
		strSql = strSql & "WHERE a_display = -1"
	End If
	
	Dim rsAction : Set rsAction = Server.CreateObject("ADODB.Recordset")
	rsAction.Open strSql, objConn, 1, 2, &H0001

	With Response
		.Write ("<select name=action>")
		.Write ("<option value="""">No Action</option>")
		Do While Not rsAction.Eof
			If strCode = rsAction(0) Then
				strSelected = " selected"
			Else
				strSelected = ""
			End If
			.Write ("<option value=""" & rsAction(0) & """" & strSelected & ">" & rsAction(1) & "</option>")
			rsAction.Movenext
		Loop
	End With
	
	rsAction.Close : Set rsAction = Nothing
	
End Sub

Sub DisplayCampaignsSelect(strCode)
	
	Dim strSelected, strName
	
	Dim strSql : strSql = "SELECT ca_code, ca_name " &_
		"FROM " & aryMTDB(5) & "Campaigns "
	
	Dim rsCampaign : Set rsCampaign = Server.CreateObject("ADODB.Recordset")
	rsCampaign.Open strSql, objConn, 1, 2, &H0001

	With Response
		.Write ("<select name=campaign>")
		.Write ("<option value="""">Select a Campaign...</option>")
		Do While Not rsCampaign.Eof
			If strCode = rsCampaign(0) Then
				strSelected = " selected"
			Else
				strSelected = ""
			End If
			
			strName = rsCampaign(1)
			If strName = "" Then
				strName = rsCampaign(0)
			End If
			
			.Write ("<option value=""" & rsCampaign(0) & """" & strSelected & ">" & strName & "</option>")
			rsCampaign.Movenext
		Loop
	End With
	
	rsCampaign.Close : Set rsCampaign = Nothing
	
End Sub


Set objReport = Nothing
Call CloseDatabaseConnection()
%>
