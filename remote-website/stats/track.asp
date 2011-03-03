<%
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
<% 'Option Explicit %>
<!--#Include File="config.asp"-->
<!--#Include File="conn.asp"-->
<!--#Include File="core.asp"-->
<!--#Include File="clsLog.asp"-->
<%

' GET URL FOR LOGGING / REDIRECTION OR REFERRER
Dim strUrl : strUrl	= Request.Querystring("mtr")
Dim blnImage : blnImage = CBool(Request.Querystring("mti"))

' CHECK WHAT TYPE OF LOGGING METHOD IS BEING USED
' 0 - ASP EXECUTE METHOD
' 1 - REDIRECT FILE METHOD
' 2 - JAVASCRIPT METHOD

' SET LOGGING TYPE IN CASE UNSPECIFIED
Dim intType
If strUrl <> "" Then
	intType		= 1
Else
	intType		= 0
End If

' GET LOGGING TYPE IF SPECIFIED
If Request.Querystring("mtt") <> "" Then 
	intType = Request.Querystring("mtt")
End If

' GET SCREENAREA IF AVAILABLE
Dim strScreenarea
If Request.Querystring("mts") <> "x" Then
	strScreenArea = Request.Querystring("mts")
End If

Dim blnExclude
If Request.Cookies("mt_exclude") <> "" Then
	blnExclude = True
Else
	blnExclude = False
End If

' GET ACTION DATA
Dim strAction, strAmount, strOrder, strPageTitle

If intType = 0 Then
	strPageTitle	= Request.Cookies("mt")("pagetitle")
	strAction		= Request.Cookies("mt")("action")
	strAmount 		= Request.Cookies("mt")("amount")
	strOrder 		= Request.Cookies("mt")("order")
ElseIf intType = 2 Then
	strPageTitle	= Request.Querystring("mtpt")
	strAction 		= Request.Querystring("mtac")
	strAmount 		= Request.Querystring("mta")
	strOrder 		= Request.Querystring("mto")
End If

' LOG REQUEST IF LOGGING IS ENABLED
If (aryMTConfig(2) = True Or aryMTConfig(2) = "") And blnExclude = False Then

	' INSTANTIATE OBJECT FROM CLASS.ASP FILE
	Dim objTrack : Set objTrack = New MTLog
	
	Call CreateDatabaseConnection(0)
	
	' SET SOME PROPERTIES
	With ObjTrack
		.Database		= aryMTDB
		.Config			= aryMTConfig
		.Action			= strAction
		.Amount			= strAmount
		.Order			= strOrder
		.PageTitle		= strPageTitle
	End With

	' CHECK TO SEE IF IP MATCHES LOG EXCLUSION LIST
	If Not objTrack.MatchIPAddress(aryMTConfig(3)) Then
		' PERFORM LOGGING OPERATION
		Call objTrack.LogFile(strUrl, intType, strScreenArea)
	End If
	
	Set objTrack = Nothing
	
	Call CloseDatabaseConnection()

End If

' REDIRECT TO PAGE IF USING REDIRECT FILE METHOD (intType = 1)
If CInt(intType) = 1 Then
	Response.Redirect strUrl
End If

If blnImage = True Then
	Response.ContentType ="image/gif"
%>
<!--#Include File="images/spacer.gif"-->
<% End If %>
