<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")
end if
response.buffer=true
%>
<html>
<head><title>HouseRabbit.org Rabbits Admin</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/sumipntg/sumi1011.css"><meta name="Microsoft Theme" content="sumipntg 1011, default">
</head>
<body>
<center>
<h1>Rabbits For HouseRabbit</h1>
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
	Call DisplayRecord(0,NULL,NULL,NULL,NULL,NULL,false,false,NULL,"Adding")
case "Edit"
	sql="Select * from adopt Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Name"),rsJokes("Picture1"),rsJokes("Picture2"),rsJokes("Picture3"),rsJokes("Picture4"),rsJokes("adopted"),rsJokes("archive"),rsJokes("Desc"),"Updating")
case "Delete"
	sql="Select * from adopt Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Name"),rsJokes("Picture1"),rsJokes("Picture2"),rsJokes("Picture3"),rsJokes("Picture4"),rsJokes("adopted"),rsJokes("archive"),rsJokes("Desc"),"Deleting")
case "Submit New"
	sql="Select Max(Id) as MaxId From adopt"
	set rsMax=Conn.execute(sql)
	iMax=rsMax("MaxId")+1
	set rsMax=nothing
	If isnull(iMax) then iMax=1
	sql="INSERT INTO [adopt] ([id],[name],[Picture1],[Picture2],[Picture3],[Picture4],[Adopted],[Archive],[Desc]) VALUES("
	sql=sql & iMax & ","
	sql=sql & "'" & replace(request("Name"),"'","''") & "',"
	sql=sql & "'" & replace(request("Picture1"),"'","''") & "',"
	sql=sql & "'" & replace(request("Picture2"),"'","''") & "',"
	sql=sql & "'" & replace(request("Picture3"),"'","''") & "',"
	sql=sql & "'" & replace(request("Picture4"),"'","''") & "',"
	If request("Adopted") then
		sql=sql & "1,"
	else
		sql=sql & "0,"
	end if
	If request("Archive") then
		sql=sql & "1,"
	else
		sql=sql & "0,"
	end if
	sql=sql & "'" & replace(request("Desc"),"'","''") & "')"
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
	sql="UPDATE adopt SET "
	sql=sql & "[Name]='" & replace(request("Name"),"'","''") & "',"
	sql=sql & "[Picture1]='" & replace(request("Picture1"),"'","''") & "',"
	sql=sql & "[Picture2]='" & replace(request("Picture2"),"'","''") & "',"
	sql=sql & "[Picture3]='" & replace(request("Picture3"),"'","''") & "',"
	sql=sql & "[Picture4]='" & replace(request("Picture4"),"'","''") & "',"
	If request("Adopted") then
		sql=sql & "[adopted]=1,"
	else
		sql=sql & "[adopted]=0,"
	end if
	If request("Archive") then
		sql=sql & "[Archive]=1,"
	else
		sql=sql & "[Archive]=0,"
	end if
	sql=sql & "[Desc]='" & replace(request("Desc"),"'","''") & "'"
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
	sql="DELETE FROM adopt WHERE id=" & request("Index")
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
	sql="Select ID,Name From adopt ORDER BY Name"
	set rsWhatsNew=Conn.execute(sql)
%>
	<form name="FAQ" action="Adminadopt.asp" method="post">
	<b>Select a rabbit Entry</b> <select name="Index">
<%
	Do while not rsWhatsNew.eof
		response.write "<option value=""" & rsWhatsNew("Id") & """>" & rsWhatsNew("Name") 
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
Sub DisplayRecord(iIndex,sName,sPicture1,sPicture2,sPicture3,sPicture4,iAdopted,iArchive,sDesc,sMode)
%>
	<font size="+2"><%=sMode%> A Record</font><br>
	<form name="FAQ" action="Adminadopt.asp" method="post">
	<input type="hidden" name="mode" value="<%=sMode%>">
	<input type="hidden" name="Index" value="<%=iIndex%>">
	<table border="1" cellpadding="5" cellspacing="0">
	<tr>
		<th align="right">Name</th>
		<td><input name="Name" type="Name" size="30" value="<%=sName%>"></td>
	</tr>
	<tr>
		<th align="right">Picture 1</th>
		<td><input name="Picture1" type="text" size="30" value="<%=sPicture1%>"></td>
	</tr>
	<tr>
		<th align="right">Picture 2</th>
		<td><input name="Picture2" type="text" size="30" value="<%=sPicture2%>"></td>
	</tr>
	<tr>
		<th align="right">Picture 3</th>
		<td><input name="Picture3" type="text" size="30" value="<%=sPicture3%>"></td>
	</tr>
	<tr>
		<th align="right">Picture 4</th>
		<td><input name="Picture4" type="text" size="30" value="<%=sPicture4%>"></td>
	</tr>
	<tr>
		<th align="right">Adopted</th>
<%
		if iAdopted then
			sChecked=" CHECKED"
		else
			sChecked=""
		end if
%>
		<td><input name="Adopted" type="checkbox" value="1" <%=sChecked%>></td>

	</tr>
	<tr>
		<th align="right">Archive</th>
<%
		if iArchive then
			sChecked=" CHECKED"
		else
			sChecked=""
		end if
%>
		<td><input name="Archive" type="checkbox" value="1" <%=sChecked%>></td>
	</tr>
	<tr>
		<th align="right">Description</th>
		<td><textarea name="desc" cols="50" rows="10"><%=sDesc%></textarea></td>
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
