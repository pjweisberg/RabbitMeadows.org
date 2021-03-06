﻿<%@ LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>
<!-- #include file="correct-domain.asp"-->
<!--#include file="adovbs.inc"-->


<%



	' Read in the value of the option selected and assign a variable for page titles
'----------------------------------------------------

Dim Status
Status=Request("Status")
If Status = "" then
Status=1
end if

Dim title, blurb
Select Case Status
  Case 1:
  title="Rodents Looking for a Good Home"  


  Case 2:
  title="Rodents Adopted to a Home of Their Own"  

  

End Select

 
%>
<HTML>

<HEAD>
<TITLE>Rabbit Meadows Sanctuary & Adoption Center - <%=title%> </TITLE>

<!--#include file="dropdownmenu.asp"-->
<!--#include file="google-analytics.js"-->
</head>

<BODY BGCOLOR="#ffffff" LINK="#000000"  ALINK="#8FBC8F" VLINK="#2E8B57">
<CENTER>

	<!--Main, all encompassing table-->

<TABLE WIDTH=95% HEIGHT=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor="#ffffff">

<TR BGCOLOR="#ffffff"><TD COLSPAN=2 ALIGN=CENTER VALIGN=BOTTOM>

</TD></TR>

<tr><td>
<!--#include file="headerfile.asp"-->
</td></tr>
	<!--end of page header section-->

<TR><TD VALIGN=TOP bgcolor="#ffb563" ALIGN=CENTER width=100%>

	<!--inner table with info and rodents-->
	
	<TABLE CELLPADDING=5 cellspacing=4 width=100%>
		<TR><TD ALIGN=left BGCOLOR="#FFFFFF" width=100%><br><p align=center>
		<FONT FACE="ARIAL" SIZE=5><%=title%></FONT> <p ALIGN=LEFT>

<%
'---------------------------------------------------------------------------
'Custom Heading to category
'---------------------------------------------------------------------------
If Status=1 or Status=2 then
%>
<font size 3 face="Arial"><p align="center">
<FONT SIZE=2 FACE="ARIAL">Please visit our adoption center at: <br>
<b>Rabbit Meadows<br> 
8311 252nd Ave NE, Redmond, WA 98053 - 425-836-8925<br>Open without appointment: Sat,Sun noon-5pm </b><br><p align=left>



	<font face="arial" size="2">The
    Rabbit Meadows Sanctuary & Adoption Center is a non-profit animal welfare organization
    dedicated to:
	<ul>
      <li><font face="arial" size="2" >The rescue and then
        adoption of rodents who have lost their home, through no fault of their own. </li>
      <li><font face="arial" size="2" >The education of
        persons interested in sharing their homes with a rodent.</li>
    </ul>
    
	<p>
  
  <font face="arial" size="2" >We rescue
    guinea pigs, rats, mice, gerbils,&nbsp; hamsters, prairie dogs, chinchillas and the
    occasional hedgehog. We work with local animal shelters and accept these animals from them
    as space permits. (Only <b>if</b> a shelter does not have rodents needing our help will we consider
    taking in your rodents.) 
	
	<p><font face="arial" size="2">We recommend that all rats be spayed or neutered to prevent hormonal tumors and early
    deaths, as well as a means to allow these animals to live humanely with their own species. Adoption Fee for spayed/neutered rats
	
	<p>



We only do adoptions within the Seattle and Puget Sound area. Companion rodents must live inside in order to be part of the family. If interested in a particular animal, you must come in to the shelter to meet one another. 

<p align=left>
<%
end if
%>

</FONT>
<P>

	<!--Populate selection drop-down box-->

<FORM METHOD="get" ACTION="RodentCurrent.asp"><CENTER>
<FONT SIZE=2 COLOR=#b8860b FACE=ARIAL><B>Our Rodents</B></FONT>

<SELECT NAME="Status" LENGTH="15">
<OPTION VALUE=1 selected>Looking for a home
<OPTION VALUE=2>Recently Adopted
</SELECT>

<INPUT TYPE="submit" VALUE="Go"></CENTER>
</FORM>
<P align=center>

<%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")


%>

<!--#include file="connstr.asp" -->

<%


objConn.Open sConnect

'---------------------------------------------------------------
'Set max number of records to display per page 
'---------------------------------------------------------------

Dim pg

pg=Trim(Request("pg"))
If pg="" Then pg=1


Dim objRS
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorType=adOpenStatic
objRS.PageSize=12


'-------------------------------------------------------------------------------------------
'Open source table, determine total number of pages required and print links to all pages
'------------------------------------------------------------------------------------------
Dim strSQL
If Status=1 then
strSQL="Select * from AdoptRodents where adopted=False and archive=0 Order by Name" 
'and status=1"
elseif Status=2 then
strSQL="Select * from AdoptRodents where adopted=True" 
'and status=2"
elseif Status=3 then
strSQL="Select * from AdoptRodents where archive=0" 
'and status=3"
end if

objRS.Open strSQL, objConn


If Not objRS.EOF THEN

	objRS.AbsolutePage=pg
	Dim i, pgcount
	If objRS.PageCount>0 then
	pgcount=objRS.PageCount
        else
	pgcount=0
	end if
	Response.Write "<B>Page " & pg & " of " & pgcount & "<BR>"
		If pgcount>1 Then
		Response.Write "<b>Go to page "		
		 For i=1 to objRS.PageCount
	 		If i <> cInt(pg) THEN 
%>
			<a href="RodentCurrent.asp?Status=<%=Status%>&pg=<%=i%>"> <%= i %> </a> |
<%
			
			Else Response.Write i & "|"

			End if
		Next
		end if
end if
%>
<a href="/shelter/index.asp">Home</a>|<p>

<%

'-----------------------------------------------------------
'If no records to display, print message to screen
'----------------------------------------------------------

Dim norec
norec=false
If objRS.EOF then
Response.Write "<FONT SIZE=4 FACE=ARIAL COLOR=darkblue>There are no records in this category at the moment.<BR> Please check back!</FONT><BR>"
norec=true
End if
%>

</TD></TR></TABLE>
</TD></TR>

<!--End title and general info section, begin rabbit display-->
<TR><TD>
<%

Do While Not objRS.EOF and rowCount < objRS.PageSize
	Dim rowCount
	rowCount=rowCount+1
%>


<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#ffffff">

<%

'--------------------------------------------------------------------------------
'Read and display info from rabbits table into individual nested tables, 
'alternate printing image first, then text, and text first then image
'------------------------------------------------------------------------------------

dim varImg1, varWidth, varHeight
varImg1=objRS("Picture1")
if varImg1="" or IsNull(objRS("Picture1")) then
varImg1="noimage.gif"
varWidth=181
varHeight=67
end if

If rowcount mod 2 = 0 then
%>
<tr><td width=100%><table bgcolor="#ffb563" cellpadding=8 CELLSPACING=4 width=100%>
<tr>
<TD BGCOLOR="ivory" align="center">
<IMG SRC="RodentImages/<%=varImg1%>"  ALT="<%=objRS("Name")%>" border=1>
<%
if not isNull(objRS("Picture2")) and objRS("Picture2")<>""  then
%>
<IMG SRC="RodentImages/<%=objRS("Picture2")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture3")) and objRS("Picture3")<>""  then
%>
<IMG SRC="RodentImages/<%=objRS("Picture3")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture4")) and objRS("Picture4")<>""  then
%>
<IMG SRC="RodentImages/<%=objRS("Picture4")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>

</TD>

<TD BGCOLOR="ivory"><FONT FACE="ARIAL" color="#000000">
<B><%=objRS("Name")%></B><P>
<%=objRS("Desc")%>
</FONT></TD></tr></table>
</td></tr>

<%
else
%>
<tr><td width=100%><table bgcolor="9c9c42" cellpadding=4 CELLSPACING=4 width=100%>
<tr>

<TD BGCOLOR="ivory"><FONT FACE="ARIAL" color="#000000">
<B><%=objRS("Name")%></B><P>
<%=objRS("Desc")%>
</FONT></TD>

<TD BGCOLOR="ivory" ALIGN="center">
<IMG SRC="RodentImages/<%=varImg1%>" ALT="<%=objRS("Name")%>" border=1>
<%
if not isNull(objRS("Picture2")) and objRS("Picture2")<> "" then
%>
<IMG SRC="RodentImages/<%=objRS("Picture2")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture3")) and objRS("Picture3")<>""  then
%>
<IMG SRC="RodentImages/<%=objRS("Picture3")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture4")) and objRS("Picture4")<>""  then
%>
<IMG SRC="RodentImages/<%=objRS("Picture4")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>


</TD></tr></table>
</td></tr>

<%
end if
%>

</TABLE><p>

<%
objRS.MoveNext

Loop

%>

</TD></TR>

<TR><TD ALIGN=CENTER>

<%

If norec=false then
 Response.Write "<B>Page " & pg & " of " & pgcount & "<BR>"
		If pgcount>1 Then
		Response.Write "<b>Go to page "		
		 For i=1 to objRS.PageCount
	 		If i <> cInt(pg) THEN 
%>
			<a href="RodentCurrent.asp?Status=<%=Status%>&pg=<%=i%>"> <%= i %> </a> |
<%
			
			Else Response.Write i & "|"

			End if
		Next
		end if
 end if
If norec = false then
%>
<a href="/shelter/index.asp">Home</a>|
<%
end if
objRS.Close


Set objRS = Nothing
objConn.Close

Set objConn=Nothing
%>
</TD></TR>


<TR BGCOLOR="#EEE8AA"><TD COLSPAN=2 ALIGN=CENTER VALIGN=BOTTOM>

</TD></TR>
<tr><td colspan=2 align=center>
<p>&nbsp;<br>
<!--#include file="footer.asp"-->
</td></tr>
</TABLE></center>

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>