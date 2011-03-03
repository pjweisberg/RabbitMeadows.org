<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")
end if
response.buffer=true
%>
<html>
<head><title>HouseRabbit.org FAQ Admin</title><!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/sumipntg/sumi1011.css"><meta name="Microsoft Theme" content="sumipntg 1011, default">
</head>
<body>
<center>
<h1>FAQ For HouseRabbit</h1>
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
	Call DisplayRecord(0,NULL,NULL,"Adding")
case "Edit"
	sql="Select id,title,Text from FAQ Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Title"),rsJokes("Text"),"Updating")
case "Delete"
	sql="Select id,Title,Text from FAQ Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Title"),rsJokes("Text"),"Deleting")
case "Submit New"
	sql="Select Max(id) as MaxId From FAQ"
	set rsMax=Conn.execute(sql)
	iMax=rsMax("MaxId")+1
	set rsMax=nothing
	If isnull(iMax) then iMax=1
	sql="INSERT INTO FAQ ([id],[Title],[Text]) VALUES("
	sql=sql & iMax & ","
	sql=sql & "'" & replace(request("Title"),"'","''") & "',"
	sql=sql & "'" & replace(request("Text"),"'","''") & "')"
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
	sql="UPDATE FAQ SET "
	sql=sql & "[Title]='" & replace(request("Title"),"'","''") & "',"
	sql=sql & "[Text]='" & replace(request("Text"),"'","''") & "'"
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
	sql="DELETE FROM FAQ WHERE id=" & request("Index")
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
	sql="Select Title,Id From FAQ ORDER BY title"
	set rsWhatsNew=Conn.execute(sql)
%>
	<form name="FAQ" action="AdminFAQ.asp" method="post">
	<b>Select a FAQ Entry</b> <select name="Index">
<%
	Do while not rsWhatsNew.eof
		response.write "<option value=""" & rsWhatsNew("Id") & """>" & rsWhatsNew("Title") & "</option>"
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
Sub DisplayRecord(iIndex,sTitle,sText,sMode)
%>
	<font size="+2"><%=sMode%> A Record</font><br>
	<form name="FAQ" action="AdminFAQ.asp" method="post">
	<input type="hidden" name="mode" value="<%=sMode%>">
	<input type="hidden" name="Index" value="<%=iIndex%>">
	<table border="1" cellpadding="5" cellspacing="0">
	<tr>
		<th align="right">Title</th>
		<td><input name="Title" type="text" size="30" value="<%=sTitle%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Text</th>
		<td><textarea name="Text" rows="10" cols="80"><%=sText%></textarea></td>
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
