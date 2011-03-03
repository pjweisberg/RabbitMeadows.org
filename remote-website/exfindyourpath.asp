<HTML>
<HEAD>
<TITLE>Path Help For Premium Members</TITLE>
<STYLE TYPE="text/css">
<!--
.Font8				{font-family:arial; font-size:8pt;}
.Font10			 	{font-family:arial;   font-size:10pt; }
.Font12Bold			 	{font-family:arial;   font-size:12pt; font-weight:bold;}
.Font12BoldBlue 	{font-family:arial;   font-size:12pt; font-weight:bold; color:#336699;}
.Font14BoldBlue 	{font-family:arial;   font-size:14pt; font-weight:bold; color:#336699;}
-->
</STYLE>
</HEAD>
<BODY>
<CENTER>
<FONT CLASS='Font14BoldBlue'>Brinkster Working Examples</FONT>
<HR COLOR=#336699 WIDTH=260>
</CENTER>

<BR><BR>
<FONT CLASS=Font12BoldBlue>Server.MapPath<BR></FONT><FONT CLASS=Font10>
<HR COLOR=#336699>
<%
'-------------------------------------------------------------------------Sever.MapPath
Dim sCurrentLocation, sRootDirectory

sCurrentLocation = Server.MapPath ("exFindYourPath.asp")'-----------------here we set a variable equal to the path of this file.

Response.Write ("This function will produce a physical path based on the information you pass it (''within the quotation marks'') and the current path of the file you are working out of. ")
Response.Write ("It will not give you an error if the information you pass it is incorrect.  It just assumes the information is correct.<BR><BR>")
Response.Write ("So, double check the result with your directory structure.<BR><BR>")


Response.Write ("Using <b>&lt;&#37;Server.MapPath (''FileInUse.asp'')&#37;&gt;</b> you can find the path of the file you are currently using.<BR><BR>")


Response.Write ("Using Server.MapPath you can see that this file is located at: <BR></FONT><FONT CLASS=Font12Bold>")
Response.Write (sCurrentLocation & "<BR><BR><BR>")



'-------------------------------------------------------------------------Referencing Other Files
Response.Write ("</FONT><FONT CLASS=Font12BoldBlue>Referencing Other Files<BR></FONT><FONT CLASS=Font10>")
Response.Write ("<HR COLOR=#336699>") 
Response.Write ("Let's say we want to reference a file called <b>''Test.asp''</b>.  We know this file is located at <b>''WebRoot\TestFiles\Test.asp''. </b>")
Response.Write ("The file we a currently working in is located at <b>''Webroot\Production\Working.asp''</b>.  So, how do we reference <b>''Test.asp''</b> from the file <b>''Working.asp''</b>?<BR><BR>")

Response.Write ("Just like this. <b>&lt;&#37;Server.MapPath (''\TestFiles\Test.asp'')&#37;&gt;</b><BR><BR>")

Response.Write ("By putting the ''\'' at the beginning we are telling ASP to start at the WebRoot directory and move into the TestFiles directory looking for the Test.asp file.<BR><BR>")

Response.Write ("Now let's say we want to reference a file called <b>''FindMe.asp''</b>.  We know this file is located at <b>''WebRoot\Production\Search\FindMe.asp''</b>. ")
Response.Write ("We are still working at <b>''Webroot\Production\Working.asp''</b>.  So how do we get to <b>''FindMe.asp''</b>.<BR><BR>")

Response.Write ("With a small change. <b>&lt;&#37;Server.MapPath (''Search\FindMe.asp'')&#37;&gt;</b><BR><BR>")

Response.Write ("This time we did not use the ''\''.  We are starting from our current directory (Production) going up one (to Search) and then referencing FindMe.asp.<BR><BR><BR>")



'-------------------------------------------------------------------------Referencing Your DataBase
Response.Write ("</FONT><FONT CLASS=Font12BoldBlue>Referencing Your DataBase<BR></FONT><FONT CLASS=Font10>")
Response.Write ("<HR COLOR=#336699>") 
Response.Write ("As a Premium Member your database can be located outside of you webroot directory.  This is done so that it can not be downloaded when a person references it in the URL.<BR><BR>")



'--------------------------------------------------- some parsing code that will display 6 "\" from the left.
Dim iCounter, iLocation, sTemp, iTemp, iLength		
sTemp = sCurrentLocation								
iLength = Len(sCurrentLocation)						
iCounter = 0										
While iCounter <= 3									
	iTemp = InStr(sTemp, "\")						
	iLocation = iTemp + iLocation					
	iLength = iLength - iTemp						
	sTemp = Right(sTemp, iLength)					
	iCounter = iCounter + 1							
Wend												
sRootDirectory = Left(sCurrentLocation, instr(lcase(sCurrentLocation),"\webroot")-1)
'--------------------------------------------------- end of parseing


Response.Write ("The path should look like this:<BR>")
Response.Write ("</FONT><FONT CLASS=Font12Bold>")
Response.Write (sRootDirectory) & "\database\YourDB.mdb<BR><BR>"

Response.Write ("</FONT><FONT CLASS=Font10>Make sure you put your database in the Database directory.  The Database directory is the only directory with write permissions.<BR><BR>")


Response.Write ("</FONT><FONT CLASS=Font12BoldBlue>Referencing Include Files<BR></FONT><FONT CLASS=Font10>")
Response.Write ("<HR COLOR=#336699>") 
Response.Write ("There are two different ways to reference an Include.<BR><BR>")

Response.Write ("<b>&lt;!-- #Include Virtual=''/includes/FileName.asp --&gt;</b> or <b>&lt;!-- #Include File=''includes\FileName.asp'' --&gt;</b><BR><BR>")

Response.Write ("With <b>Include Virtual</b> you can only reference files within your virtual directories (everything back to the WebRoot). ")
Response.Write ("The same principles apply to <b>Include Virtual</b> as they did with <b>Server.MapPath</b>.  Again, if you put a ''/'' at the beginning of the reference you will start at the WebRoot directory and so on. <BR><BR>")

Response.Write ("With <b>Include File</b> you can use the entire path to reference the file.<BR>")
Response.Write ("<b>&lt;!-- #Include File=''" & sRootDirectory & "webroot\includes\FileName.asp'' --&gt;</b><BR><BR>")
Response.Write ("Or, you can start at the directory you are in.  In this case we are starting in the WebRoot directory, then moving into the Includes directory, referencing FileName.asp.<BR>")
Response.Write ("<b>&lt;!-- #Include File=''includes\FileName.asp'' --&gt;</b><BR><BR>")
Response.Write ("You can not put a ''\'' at the beginning to start at the WebRoot directory when using <b>Include File</b>.")
Response.Write ("<BR><BR>")




%>

</FONT>
<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>
</BODY>
</HTML>
