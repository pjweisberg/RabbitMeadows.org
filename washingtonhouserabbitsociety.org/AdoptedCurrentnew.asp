<%@ LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>

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
  title="Ferrets Looking for a Good Home"  


  Case 2:
  title="Ferrets Adopted to a Home of Their Own"  

  

End Select

 
%>
<HTML>

<HEAD>
<TITLE>Ferret Adoption - <%=title%> </TITLE>
<STYLE TYPE="text/css"><!--A { text-decoration: none }A:hover { text-decoration: underline }--></STYLE>
</HEAD>

<BODY BGCOLOR="#ffffff" LINK="#B8860B"  ALINK="#8FBC8F" VLINK="#2E8B57">
<CENTER>

	<!--Main, all encompassing table-->

<TABLE WIDTH=95% HEIGHT=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor="#ffffff">

<TR BGCOLOR="#ffffff"><TD COLSPAN=2 ALIGN=CENTER VALIGN=BOTTOM>

</TD></TR>

<tr><td>
<!--#include file="washlinks.asp"-->
</td></tr>
	<!--end of page header section-->

<TR><TD VALIGN=TOP bgcolor="#ffb563" ALIGN=CENTER width=100%>


<table border="0">
  <tr>
    <td align="center"></td>
  </tr>

 
    <td><a href="news.asp"><b>What's News</b></a></td>

  <tr>
    <td><a href="AdoptedCurrent.asp"><b>AdoptMe Rabbits</b></a></td>
  </tr>
  <tr>
    <td><a href="Membership.asp"><b>Membership</b></a></td>
  </tr>
  <tr>
    <td><a href="http://www.rabbitrodentferret.org"><b>On-line Store</b></a></td>
  </tr>
  <tr>
    <td><a href="FAQ.asp"><b>Questions?</b></a></td>
  </tr>
  <tr>
    <td><a href="vets.asp"><b>Vet Referrals</b></a></td>
  </tr>
<tr>
    <td><a href="Youth/default.htm"><b>Youth House Rabbit Club</b></a></td>
  </tr>
  
  <tr>
    <td><a href="sites.asp"><b>Other Links</b></a></td>
  </tr>
  <tr>
    <td><a href="HouseBun.asp"><b>HouseBun Online</b></a></td>
  </tr>
  <tr>
    <td><b><a href="RodentCurrent.asp">AdoptMe Rodents</a></b></td>
  </tr>
  <tr>
    <td><b><a href="FerretCurrent.asp">AdoptMe Ferrets</a></b></td>
  </tr>
  <tr>
    <td><a href="AboutHRS.asp"><b>About Us</b></a></td>
  </tr>
</table>

<!--webbot bot="Include" endspan i-checksum="14527" -->
</td>
    <td width="10%" rowspan="4"></td>
    <td width="76%" colspan="6" valign="top" align="center"><font face="comic sans ms, arial"
    size="5" color="#400040"><b>Welcome to the </b></font><p><font face="comic sans ms, arial"
    size="2" color="#400040"><span style="line-height: 5px; vertical-align: baseline"><b>Best
    Little Rabbit,Rodent &amp; Ferret House </b></span></font></p>
    <p><b><font face="comic sans ms, arial" size="5" color="#400040">House Rabbit Society<br>
    Washington State Chapter</font><font face="comic sans ms, arial" size="4" color="#400040"><br>
    </font></b></td>
  </tr>
  <tr>
    <td width="76%" colspan="6"><font face="comic sans ms, arial" size="2" color="#400040"><strong>T</strong>he
    House Rabbit Society is an all-volunteer, non-profit organization with two primary goals:</font>
    <ul>
      <li><font face="comic sans ms, arial" size="2" color="#400040">To rescue abandoned rabbits
        and find permanent homes for them and </font></li>
      <li><font face="comic sans ms, arial" size="2" color="#400040">To educate the public and
        assist animal shelters, through publications on rabbit care, phone consultation, and
        classes upon request</font> </li>
    </ul>
    <p><font face="comic sans ms, arial" size="3" color="#400040"><strong>S</strong></font><font
    face="comic sans ms, arial" size="2" color="#400040">ince 1988 over 8,120 rabbits have
    been rescued through House Rabbit Society foster homes across the country. The House
    Rabbit Society has been granted a tax-exempt status under the Internal Revenue code for
    prevention of cruelty to animals.</font></p>
    <p><font face="comic sans ms, arial" size="2" color="#FF0000"><strong>For the latest news,
    please visit the &quot;<a href="news.asp">What's News</a>&quot; section.</strong></font></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="12%"></td>
    <td width="13%"></td>
    <td width="13%"></td>
    <td width="13%"></td>
    <td width="13%"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="12%"></td>
    <td width="13%"></td>
    <td width="13%"></td>
    <td width="13%"></td>
    <td width="13%"></td>
  </tr>
</table>


<hr width="90%">

	<!--inner table with info and ferrets-->
	
	<TABLE CELLPADDING=5 cellspacing=4 width=100%>
		<TR><TD ALIGN=CENTER BGCOLOR="#FFFFFF" width=100%><BR>
		<FONT FACE="ARIAL" SIZE=5><%=title%></FONT> <p ALIGN=LEFT>

<%
'---------------------------------------------------------------------------
'Custom Heading to category
'---------------------------------------------------------------------------
If Status=1 then
%>
<font size 3 face="Arial"><p align="center">
<FONT SIZE=2 FACE="ARIAL">Please visit our adoption center at: <br>
<b>The Best Little Rabbit, Rodent & Ferret House<br> 
14317 Lake city Way NE, Seattle WA 98125<br> (phone) 206-365-9105 </b><br><p align=left>
We only do adoptions within the Seattle and Puget Sound area. Companions rabbits must live inside in order to be part of the family. If interested in a particular animal, you must come in to the shelter to meet one another. 

<p align=left>
<%
end if
%>

</FONT>
<P>

	<!--Populate selection drop-down box-->

<FORM METHOD="get" ACTION="AdoptedCurrent.asp"><CENTER>
<FONT SIZE=2 COLOR=#b8860b FACE=ARIAL><B>Our Rabbits</B></FONT>

<SELECT NAME="Status" LENGTH="15">
<OPTION VALUE=1 selected>Looking for a home
<OPTION VALUE=2>Recently Adopted
</SELECT>

<INPUT TYPE="submit" VALUE="Go"></CENTER>
</FORM>

<%
'--------------------------------------------------------------------
' Open the connection
'--------------------------------------------------------------------
Dim objConn, sConnect
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
strSQL="Select * from Adopt where adopted=0 and archive=0" 
'and status=1"
elseif Status=2 then
strSQL="Select * from Adopt where adopted=1" 
'and status=2"
elseif Status=3 then
strSQL="Select * from Adopt where archive=0" 
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
			<a href="AdoptedCurrent.asp?Status=<%=Status%>&pg=<%=i%>"> <%= i %> </a> |
<%
			
			Else Response.Write i

			End if
		Next
		end if
end if
%>
<a href="default.htm">Home</a>|<p>

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

dim varImg1
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
<IMG SRC="BunnyImages/<%=varImg1%>"  ALT="<%=objRS("Name")%>" border=1>
<%
if not isNull(objRS("Picture2")) and objRS("Picture2")<>""  then
%>
<IMG SRC="BunnyImages/<%=objRS("Picture2")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture3")) and objRS("Picture3")<>""  then
%>
<IMG SRC="BunnyImages/<%=objRS("Picture3")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture4")) and objRS("Picture4")<>""  then
%>
<IMG SRC="BunnyImages/<%=objRS("Picture4")%>"  ALT="<%=objRS("Name")%>" border=1>
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
<IMG SRC="BunnyImages/<%=varImg1%>" ALT="<%=objRS("Name")%>" border=1>
<%
if not isNull(objRS("Picture2")) and objRS("Picture2")<> "" then
%>
<IMG SRC="BunnyImages/<%=objRS("Picture2")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture3")) and objRS("Picture3")<>""  then
%>
<IMG SRC="BunnyImages/<%=objRS("Picture3")%>"  ALT="<%=objRS("Name")%>" border=1>
<%
end if
%>
<%
if not isNull(objRS("Picture4")) and objRS("Picture4")<>""  then
%>
<IMG SRC="BunnyImages/<%=objRS("Picture4")%>"  ALT="<%=objRS("Name")%>" border=1>
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
			<a href="AdoptedCurrent.asp?Status=<%=Status%>&pg=<%=i%>"> <%= i %> </a> |
<%
			
			Else Response.Write i

			End if
		Next
		end if
 end if
If norec = false then
%>
<a href="default.htm">Home</a>|
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
<tr><td colspan=2 align=center><img src="footer2c.gif" usemap="#linkmap" alt="" border=0>
</TABLE></center>

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>