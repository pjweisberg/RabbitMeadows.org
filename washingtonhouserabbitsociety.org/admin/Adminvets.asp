<%
If not Session("UserLevel")=2 then
	response.redirect("default.asp")
end if
response.buffer=true
%>
<html>
<head><title>HouseRabbit.org vets Admin</title><!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/sumipntg/sumi1011.css"><meta name="Microsoft Theme" content="sumipntg 1011, default">
</head>
<body>
<center>
<h1>Vets For HouseRabbit</h1>
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
	Call DisplayRecord(0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"Adding")
case "Edit"
	sql="Select * from vet Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Name"),rsJokes("HeaderLocation"),rsJokes("Location"),rsJokes("Address"),rsJokes("City"),rsJokes("Zip"),rsJokes("Phone"),rsJokes("Phone2"),rsJokes("Fax"),rsJokes("email"),rsJokes("hours"),rsJokes("comments"),"Updating")
case "Delete"
	sql="Select * from vet Where Id=" & request("Index")
	set rsJokes=Conn.execute(sql)
	Call DisplayRecord(request("Index"),rsJokes("Name"),rsJokes("HeaderLocation"),rsJokes("Location"),rsJokes("Address"),rsJokes("City"),rsJokes("Zip"),rsJokes("Phone"),rsJokes("Phone2"),rsJokes("Fax"),rsJokes("email"),rsJokes("hours"),rsJokes("comments"),"Deleting")
case "Submit New"
	sql="Select Max(id) as MaxId From vet"
	set rsMax=Conn.execute(sql)
	iMax=rsMax("MaxId")+1
	set rsMax=nothing
	If isnull(iMax) then iMax=1
	sql="INSERT INTO vet ([id],[Name],[HeaderLocation],[Location],[Address],[City],[zip],[phone],[phone2],[fax],[email],[hours],[comments]) VALUES("
	sql=sql & iMax & ","
	sql=sql & "'" & replace(request("Name"),"'","''") & "',"
	sql=sql & "'" & replace(request("HeaderLocation"),"'","''") & "',"
	sql=sql & "'" & replace(request("Location"),"'","''") & "',"
	sql=sql & "'" & replace(request("Address"),"'","''") & "',"
	sql=sql & "'" & replace(request("City"),"'","''") & "',"
	sql=sql & "'" & replace(request("zip"),"'","''") & "',"
	sql=sql & "'" & replace(request("phone"),"'","''") & "',"
	sql=sql & "'" & replace(request("phone2"),"'","''") & "',"
	sql=sql & "'" & replace(request("fax"),"'","''") & "',"
	sql=sql & "'" & replace(request("email"),"'","''") & "',"
	sql=sql & "'" & replace(request("hours"),"'","''") & "',"
	sql=sql & "'" & replace(request("comments"),"'","''") & "')"
	'response.write sql
	'On error resume next
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
	sql="UPDATE vet SET "
	sql=sql & "[Name]='" & replace(request("Name"),"'","''") & "',"
	sql=sql & "[HeaderLocation]='" & replace(request("HeaderLocation"),"'","''") & "',"
	sql=sql & "[Location]='" & replace(request("Location"),"'","''") & "',"
	sql=sql & "[Address]='" & replace(request("Address"),"'","''") & "',"
	sql=sql & "[City]='" & replace(request("City"),"'","''") & "',"
	sql=sql & "[zip]='" & replace(request("zip"),"'","''") & "',"
	sql=sql & "[phone]='" & replace(request("phone"),"'","''") & "',"
	sql=sql & "[phone2]='" & replace(request("phone2"),"'","''") & "',"
	sql=sql & "[fax]='" & replace(request("fax"),"'","''") & "',"
	sql=sql & "[email]='" & replace(request("email"),"'","''") & "',"
	sql=sql & "[hours]='" & replace(request("hours"),"'","''") & "',"
	sql=sql & "[comments]='" & replace(request("comments"),"'","''") & "'"
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
	sql="DELETE FROM vet WHERE id=" & request("Index")
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
	sql="Select Name,Location,Id From vet ORDER BY Name"
	set rsWhatsNew=Conn.execute(sql)
%>
	<form name="FAQ" action="Adminvets.asp" method="post">
	<b>Select a Vet Entry</b> <select name="Index">
<%
	Do while not rsWhatsNew.eof
		response.write "<option value=""" & rsWhatsNew("Id") & """>" & rsWhatsNew("Name") & "-" & rsWhatsNew("Location") & "</option>"
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
Sub DisplayRecord(iIndex,sName,sHeaderLocation, sLocation,sAddress,sCity,sZip,sPhone,sPhone2,sFax,semail,shours,sComments,sMode)
%>
	<font size="+2"><%=sMode%> A Record</font><br>
	<form name="FAQ" action="Adminvets.asp" method="post">
	<input type="hidden" name="mode" value="<%=sMode%>">
	<input type="hidden" name="Index" value="<%=iIndex%>">
	<table border="1" cellpadding="5" cellspacing="0">
	<tr>
		<th align="right">Name</th>
		<td><input name="Name" type="text" size="30" value="<%=sName%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Header Location</th>
		<td><input name="HeaderLocation" type="text" size="30" value="<%=sHeaderLocation%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Vet Store Name</th>
		<td><input name="Location" type="text" size="30" value="<%=sLocation%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Address</th>
		<td><input name="Address" type="text" size="30" value="<%=sAddress%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">City</th>
		<td><input name="City" type="text" size="30" value="<%=sCity%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Zip</th>
		<td><input name="Zip" type="text" size="30" value="<%=sZip%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Phone</th>
		<td><input name="Phone" type="text" size="30" value="<%=sPhone%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Phone 2</th>
		<td><input name="Phone2" type="text" size="30" value="<%=sPhone2%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Fax</th>
		<td><input name="Fax" type="text" size="30" value="<%=sFax%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Email</th>
		<td><input name="Email" type="text" size="30" value="<%=sEmail%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Hours</th>
		<td><input name="Hours" type="text" size="30" value="<%=sHours%>"></td>
	</tr>
	<tr>
		<th align="right" valign="top">Comments</th>
		<td><textarea name="comments" cols="50" rows="10"><%=sComments%></textarea></td>
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
