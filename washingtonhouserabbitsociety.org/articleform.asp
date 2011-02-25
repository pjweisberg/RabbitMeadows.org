		<META NAME="GENERATOR" Content="ASP Express">
	<META HTTP-EQUIV="Content-Type"CONTENT="text/html;CHARSET=iso-8859-1">

<% @Language=VBScript %>
<%
dim varArt
varArt=Request("articletext")
If varArt<>"" then
Const fsoForReading=1
Const fsoForWriting = 2

Dim objFSO, strPath
strPath="\\brink-premfs1\sites\premium4\washhrs\database\NewFile.txt"
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream
Set objTextStream = objFSO.OpenTextFile(strPath, fsoForWriting, True)
'Display the contents of the text file
objTextStream.WriteLine varArt

'Close the file and clean up
objTextStream.Close
Set objTextStream = Nothing
Set objFSO = Nothing

Dim objOpenFile,  strText
Set objFSO=Server.CreateObject("Scripting.FileSystemObject")
Set objOpenFile=objFSO.OpenTextFile(strPath, fsoForReading)
Response.Write "<html><body>"

Do While Not objOpenFile.AtEndOfStream
 strText = objOpenFile.ReadLine
Response.Write strText
Response.Write "<br>"
Loop
objOpenFile.Close
Set objOpenFile = Nothing
Set objFSO=Nothing


end if
%>

	
<form method="post" action="articleform.asp">
<textarea name="articletext" rows="20" cols="50" wrap="hard"> Type the text of the article here</textarea>
<input type="submit" value=Submit>
</form>
</BODY>
</HTML>