<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")
end if
response.buffer=true
%>
<html>
<head><title>HouseRabbit.org Sites Admin</title><!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/sumipntg/sumi1011.css"><meta name="Microsoft Theme" content="sumipntg 1011, default">
</head>
<body>
<center>
<h1>Sites For HouseRabbit</h1>
<hr>
<%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection")
%>

<!--#include file="connstr.asp" -->

<%
Conn.Open sConnect
'--------------------------------------------------------------------
' Do what we have to do
'--------------------------------------------------------------------
Select Case request("Button")
case "New"
	Call DisplayRecord(0,NULL,"http://","Adding")
case "Edit"
	sql="Select * from links Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Desc"),rsJokes("Link"),"Updating")
case "Delete"
	sql="Select * from links Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Desc"),rsJokes("Link"),"Deleting")
case "Submit New"
	sql="Select Max(id) as MaxId From links"
	set rsMax=Conn.execute(sql)
	iMax=rsMax("MaxId")+1
	set rsMax=nothing
	If isnull(iMax) then iMax=1
	sql="INSERT INTO links ([id],[Desc],[Link]) VALUES("
	sql=sql & iMax & ","
	sql=sql & "'" & replace(request("Desc"),"'","''") & "',"
	sql=sql & "'" & replace(request("Link"),"'","''") & "')"
	'response.write sql
	On error resume next
	Conn.execute(sql)
	If err<>0 then
		On error goto 0
		response.write "<b>Error Adding New Record</b><br>" & err.description & "<hr>"
	else
		On error goto 0
		response.write "Record Added Successfully<hr>"
	end if
	On error goto 0
	Call DisplayMenu()
case "Submit Changes"
	sql="UPDATE links SET "
	sql=sql & "[Desc]='" & replace(request("Desc"),"'","''") & "',"
	sql=sql & "[Link]='" & replace(request("Link"),"'","''") & "'"
	sql=sql & " WHERE Id=" & request("Index")
	'response.write sql
	On error resume next
	Conn.execute(sql)
	If err<>0 then
		response.write "<b>Error Editing Record</b><br>" & err.description & "<hr>"
	else
		response.write "Record Updated Successfully<hr>"
	end if
	On error goto 0

	Call DisplayMenu()
case "Delete It"
	sql="DELETE FROM links WHERE id=" & request("Index")
	'response.write sql
	On error resume next
	Conn.execute(sql)
	If err<>0 then
		response.write "<b>Error Deleting Record</b><br>" & err.description & "<hr>"
	else
		response.write "Record Deleted Successfully<hr>"
	end if
	On error goto 0

	Call DisplayMenu()
case else
	Call DisplayMenu()
end select
'--------------------------------------------------------------------
' Close the connection
'--------------------------------------------------------------------
Conn.close
Set Conn = nothing
%>
</center>
</body>
</html>

<%
'--------------------------------------------------------------------
' Display the main menu
'--------------------------------------------------------------------
Sub DisplayMenu()
	sql="Select * From links ORDER BY [Desc]"
	set rsWhatsNew=Conn.execute(sql)
%>
	<form name="FAQ" action="AdminSites.asp" method="post">
	<b>Select a Site Entry</b> <select name="Index">
<%
	Do while not rsWhatsNew.eof
		response.write "<option value=""" & rsWhatsNew("Id") & """>" & rsWhatsNew("Desc") & "</option>"
		rsWhatsNew.movenext
	loop
	set rsWhatsNew=nothing
%>
	</select><hr>
	<input type="submit" name="Button" value="New">&nbsp;&nbsp;&nbsp;
	<input type="submit" name="Button" value="Edit">&nbsp;&nbsp;&nbsp;
	<input type="submit" name="Button" value="Delete"><br>
	</form>
	<hr>
	<form action="default.asp" method="post">
	<input type="Submit" value="Back to the main Admin Page">
	</form>
<%
end sub
'--------------------------------------------------------------------
' Display a record
'--------------------------------------------------------------------
Sub DisplayRecord(iIndex,sDesc,sLink,sMode)
%>
	<font size="+2"><%=sMode%> A Record</font><br>
	<form name="FAQ" action="AdminSites.asp" method="post">
	<input type="hidden" name="mode" value="<%=sMode%>">
	<input type="hidden" name="Index" value="<%=iIndex%>">
	<table border="1" cellpadding="5" cellspacing="0">
	<tr>
		<th align="right">Desc</th>
		<td><input name="Desc" type="text" size="30" value="<%=sDesc%>"></td>
	</tr>
	<tr>
		<th align="right">Link</th>
		<td><input name="Link" type="text" size="30" value="<%=sLink%>"></td>
	</tr>
	<tr><td colspan="2" align="center">
<%
	Select case sMode
	Case "Adding"
		response.write "<input type=""submit"" name=""Button"" value=""Submit New""><br>"
		response.write "<input type=""submit"" name="""" value=""Cancel"">"
	Case "Updating"
		response.write "<input type=""submit"" name=""Button"" value=""Submit Changes""><br>"
		response.write "<input type=""submit"" name="""" value=""Cancel"">"
	Case "Deleting"
		response.write "<input type=""submit"" name=""Button"" value=""Delete It""><br>"
		response.write "<input type=""submit"" name="""" value=""Cancel"">"
	end select
%>
	</td></tr>
	</table><br>
	</form>
<%
end sub
%>
