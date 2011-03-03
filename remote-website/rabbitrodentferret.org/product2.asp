<%@ LANGUAGE=VBSCRIPT %>
<% OPTION EXPLICIT %>

<!--#include file="adovbs.inc"-->
<!--#include file="strconn.asp"-->


<%

'---------------------------------------------------
'Read in product catagory and animal type from link
'----------------------------------------------------

Dim intCat, intAnimal

intCat=Request.Querystring("Cat")
intAnimal=Request.Querystring("Animal")

If intCat="" Then
intCat = 1
End if

If intAnimal="" Then
intAnimal=1
End If

'----------------------------------------------------
'Convert animal type and catagory into text for display
'------------------------------------------------------

Function GetAnimal(intAnimal)
Select Case intAnimal
  Case 1:
	GetAnimal="RABBIT"
  Case 2:
	GetAnimal="RODENT"
  Case 3:
	GetAnimal="FERRET"
End Select
End Function

Function GetCatagory(intCat)
Select Case intCat
  Case 1:
	GetCatagory="TOYS"
  Case 2:
        GetCatagory="HOUSING"
  Case 3:
	GetCatagory="FURNISHINGS"
  Case 4:
	GetCatagory="FOOD"
  Case 5:
	GetCatagory="GROOMING"
  Case 6:
	GetCatagory="HEALTHCARE"
  Case 7:
        GetCatagory="<font size=1 face=""arial"">PEOPLE STUFF</font>"
  Case 8:
	GetCatagory="<font size=1 face=""arial"">BOOKS & VIDEOS</font>"
  Case 9:
	GetCatagory="MISC."
End Select
End Function

'---------------------------------------------------------------------
'Dimension SQL statement based on animal type for later database query
'---------------------------------------------------------------------

Dim strSQL, strSQLF

Select Case intAnimal
	
	Case 1:
		strSQL = "SELECT ID, Name, Image, Width, Height, DescriptionLng, Variation1Des, Variation1Price, Catagory, Status, Rabbits, Variations, InOrder FROM Products Where Catagory=" & intCat & "AND Rabbits=1 AND Status=1" &_
                         " Order by InOrder Desc, ID Desc"
		strSQLF = "SELECT Photo, Width, Height, FirstName," &_
				"Caption, Rotation, [Current] FROM " &_
				"Rabbits Where Rotation=" & intCat & "AND [Current]=1"
	
		Case 2:
		strSQL = "SELECT ID, Name, Image, Width, Height, DescriptionLng, Variation1Des, Variation1Price, Catagory, Status, Rodents, Variations, InOrder FROM Products Where Catagory=" & intCat & "AND Rodents=1 AND Status=1" &_
			" Order by InOrder Desc, ID Desc"
		strSQLF = "SELECT [Photo], [Width], [Height], [FirstName]," &_
				"[Caption], [Rotation], [Current] FROM " &_
				"Rodents Where [Rotation]=" & intCat & " AND [Current]=1"

	Case 3:
		strSQL = "SELECT ID, Name, Image, Width, Height, DescriptionLng, Variation1Des, Variation1Price, Catagory, Status, Ferrets, Variations, InOrder FROM Products Where Catagory=" & intCat & "AND Ferrets=1 AND Status=1" &_
			" Order by InOrder Desc, ID Desc"
		strSQLF = "SELECT [Photo], [Width], [Height], [FirstName]," &_
				"[Caption], [Rotation], [Current] FROM " &_
				"Ferrets Where [Rotation]=" & intCat & " AND [Current]=1"

End Select

%>

<HTML><HEAD>

<!--#include file="title.asp"-->

<STYLE TYPE="text/css"><!--A { text-decoration: none }A:hover { text-decoration: underline }--></STYLE>
<TITLE>Best Little Rabbit, Rodent and Ferret House - Quality Supplies for Companion Animals</TITLE>
</HEAD>

<BODY BGCOLOR="#FFCC66" LINK="#4486c4"  ALINK="#999944" vlink="#ff9533">

<FORM METHOD = GET ACTION="product.asp">

<!--Display title-->

<table BORDER=0 CELLPADDING=0 CELLSPACING=0 BGCOLOR="#FFFFFF" WIDTH=100% VALIGN=top>

	<TR>
		<TD BGCOLOR=#FFCC66></TD>		
		<TD COLSPAN="3" BGCOLOR="#FFCC66" HEIGHT="20" ALIGN=Center VALIGN=bottom><a name="topofpage"><IMG SRC="images\spacer.gif" WIDTH=530 HEIGHT=20 ALT="Best Little Rabbit, Rodent & Ferret House"></a></TD>
	</TR>

	<TR>
		<TD ROWSPAN=30 bgcolor="#FFCC66" WIDTH=5%><IMG SRC="images\spacer.gif" WIDTH=10 HEIGHT=1 alt=""></TD>
		<TD ALIGN=left VALIGN=top><IMG SRC="images\corner.gif" WIDTH="40" HEIGHT="40" ALT=""></TD>
		
<!--Display top links inside a nested table using include file-->

		<TD ALIGN=center ROWSPAN=2 valign=top>

	
			<TABLE  BGCOLOR="#FFCC66" WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
				<TR>
				  <TD HEIGHT=10 BGCOLOR="#ffffff"><IMG SRC="images\spacer.gif" WIDTH=1 HEIGHT=10 ALT=""></td>
	
				</TR>

				<TR BGCOLOR="#FFFFFF">
			          <TD ALIGN = center><!--#include file="links.asp"--></TD>
				</TR>
	
			</TABLE>
		</TD>

<!--End link display-->

	
		<TD ALIGN =right VALIGN=top><IMG SRC="images\corner1.gif" WIDTH="40" HEIGHT="40" BORDER=0 ALT=""></TD>

	</TR>

	<TR>
<!--Display side links in nested table using include file-->

		<TD ROWSPAN=6 VALIGN=top>
					
			<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=107>
				<TR>
				  <TD VALIGN=bottom><IMG SRC="images\head1.gif" WIDTH="10" HEIGHT="47" BORDER=0></TD>
				  <TD VALIGN=bottom><IMG SRC="images\head2.gif" WIDTH="87" HEIGHT="47" BORDER=0></TD>
				  <TD VALIGN=bottom ALIGN=right COLSPAN=2><IMG SRC="images\head3.gif" WIDTH="10" HEIGHT="47" BORDER=0></TD>
				</TR>
				<TR>
				  <TD>&nbsp;</TD>	
				  <TD><!--#include file="sidelink.asp"--></TD>
				  <TD BGCOLOR="#FFFFFF" WIDTH="5"><IMG SRC="images\spacer.gif" WIDTH=5 HEIGHT=1 ALT=""></TD>
				  <TD BGCOLOR="#FFCC66" WIDTH="5" ROWSPAN=2><IMG SRC="images\spacer.gif" WIDTH=5 HEIGHT=1 ALT=""></TD>
				</TR>
				<TR>
				  <TD COLSPAN=3 BGCOLOR="#FFFFFF" HEIGHT=10><IMG SRC="images\spacer.gif" WIDTH=1 HEIGHT=10 ALT=""><TD>
				</TR>
				<TR>
				  <TD COLSPAN=4 BGCOLOR="#FFCC66" HEIGHT=5><IMG SRC="images\spacer.gif" WIDTH=1 HIGHT=5></TD>
				</TR>
			</TABLE>
		</TD>

<!--End side link display-->		

		<TD HEIGHT=40><IMG SRC="images\spacer.gif" HEIGHT=40 WIDTH=1></TD>

	</TR>

	<TR>
		<TD BGCOLOR="#FFCC66" HEIGHT="5" COLSPAN=2><IMG SRC="images\spacer.gif" WIDTH=1 HEIGHT=5 ALT=""></TD>
	</TR>

	<TR>

		<td BGCOLOR="#FFFFFF" COLSPAN=2 VALIGN=top ALIGN=center>

<!--Begin product display table-->

			<CENTER><TABLE BORDER=0 CELLPADDING=4 CELLSPACING=2>
				<TR>
				  <TD COLSPAN=2 ALIGN=center VALIGN=bottom>

<table border=0 width=100%>
<tr><td rowspan=4 align=right valign=top><img src="images\spacer.gif" width=18 height=118></td>
<td align=right valign=top><img src="images\rescue.gif" width=311 height=20></td>
<td rowspan=5 valign=top align=left ><img src="images\rescue1.gif" width=20 height=160></td>
<td rowspan=4 width = 30%><td><tr>
<tr><td align=right valign=bottom> 

<!--Display Catagory Title images with variable title in center square-->

				<TABLE WIDTH=239 BORDER=0 CELLSPACING=0 CELLPADDING=0>
					<tr>
					  <TD ROWSPAN=3><IMG SRC="images\logoa.gif" WIDTH=73 HEIGHT=110 ALT=""></TD>
		    			  <TD><IMG SRC="images\logob.gif" WIDTH=111 HEIGHT=37 ALT=""></TD>
		    			  <TD ROWSPAN=3><IMG SRC="images\logod.gif" WIDTH=55 HEIGHT=110 ALT=""></TD>
					</TR>
					<TR>
					  <TD WIDTH=111 HEIGHT=40 BGCOLOR=#9CCEFF ALIGN=center><FONT FACE="arial">

<!--Call functions to print selected Animal and Catagory in title image-->

						<B><%=GetAnimal(intAnimal)%><BR><%=GetCatagory(intCat)%></B></FONT></TD>
					</TR>
					<TR><TD><IMG SRC="images\logoc.gif" WIDTH=111 HEIGHT=33 ALT=""></TD></tr>
				
				</TABLE>
			
				</td>
				</tr>

<!--End catagory title display-->

		
	<TR>
		<TD  ALIGN=right VALIGN=bottom>
<table border=0><tr><td>
			
<%
'--------------------------------------------------------------
'Open a connection to database 
'--------------------------------------------------------------

Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
'objConn.ConnectionString = Application("strConn")
objConn.ConnectionString=strconnect


objConn.Open

'---------------------------------------------------------------
'Set max number of products to display per page and create recordset objects for 
'Products table and Featured Rabbits table
'---------------------------------------------------------------

Dim pg

pg=Trim(Request("pg"))
If pg="" Then pg=1

Dim objRS
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorType=adOpenStatic
objRS.PageSize=4


'-------------------------------------------------------------------------------------------
'Open and query Products table, determine total number of pages required and print links to all pages
'------------------------------------------------------------------------------------------

objRS.Open strSQL, objConn

If Not objRS.EOF THEN
	objRS.AbsolutePage=pg
	Dim i, pgcount
	If objRS.PageCount>0 then
	pgcount=objRS.PageCount
        else
	pgcount=0
	end if
	Response.Write "<B><FONT FACE=""arial"" COLOR=#999933 SIZE=2>Page " & pg & " of " & pgcount & "<BR>"
		If pgcount>1 Then
		Response.Write "<b><FONT FACE=""arial"" COLOR=#4486c4>Go to page "		
		 For i=1 to objRS.PageCount
	 		If i <> cInt(pg) THEN 
%>

			<A HREF="product.asp?Cat=<%=intCat%>&animal=<%=intAnimal%>&pg=<%=i%>"><FONT COLOR=#ff9900><%=i%></FONT></A>

<% 
			Else 
			Response.Write i

			End if
		Next
		end if
%>

	</TD><td>
<A HREF="https://secure6.brinkster.com/blrrfh/cart.asp"><IMG SRC="images\check.gif" BORDER=0 WIDTH=45 HEIGHT=40></A></TD>
<td><img src="images\spacer.gif" width=45 height=1 ></td></TR></TABLE>
</TD></TR>
</td></tr>

<tr><td></td><td colspan=2><hr noshade width=40%></td></tr>
</table>



<!--End form with method set to get (querystring)-->

</FORM>

<% 

'---------------------------------------
'Assign a variable to order button display
'----------------------------------------

Dim strOrder
strOrder="<FONT COLOR=#999933 FACE=""arial"">Quantity:</FONT><input type=text name=""Quantity"" size=""2"" value=1><input type=submit value=""Add To Basket""></td></tr>"

'----------------------------------------------------
'Begin product display
'---------------------------------------------------

Do While Not objRS.EOF and rowCount < objRS.PageSize
	Dim rowCount
	rowCount=rowCount+1

'-----------------------------------------------------------------------
'If no image span image space with text, otherwise display image and text
'-------------------------------------------------------------------------


If IsNull(objRS("Image")) THEN

	
	Response.Write "<TR><TD ALIGN=left COLSPAN=2><FONT FACE=""arial"" SIZE=4><B>" & objRS("Name") & "</B></FONT><P>"

	Response.Write "<FONT FACE=""arial"">" & objRS("DescriptionLng") & "</FONT></TD></TR>"

	

Else
	
	Response.Write "<TR><TD WIDTH=" & objRS("Width") & " VALIGN=MIDDLE><IMG WIDTH=" & objRS("Width") & " HEIGHT=" & objRS("Height") & " SRC=""products/" & objRS("Image") & " ALT=Photo VALIGN=top></TD>"
	Response.Write	"<TD align=left><FONT FACE=""arial"" SIZE=4><B>" & objRS("Name") & "</B></FONT><P>"

	Response.Write "<FONT FACE=""arial"">" & objRS("DescriptionLng") & "</FONT></TD></TR>"

	
END IF


'------------------------------------------------------------------
'Dimension variables to hold product variations and price information
'--------------------------------------------------------------------

Dim strProduct, arrProduct, strPrice, arrPrice, strElement, strEncName
strProduct= objRS("Variation1Des")
strPrice=objRS("Variation1Price")
strEncName=Server.HTMLEncode(objRS("Name"))

'-------------------------------------------------------------------
'If there are no variations of the product print only the price and order box
'Otherwise split the array of variations and prices into seperate elements
'---------------------------------------------------------------------


If IsNull(strProduct) Then

	Response.Write "<FORM METHOD=""post"" ACTION=""https://secure6.brinkster.com/blrrfh/cart.asp"">"	
	Response.Write "<tr><TD align=right COLSPAN=2 VALIGN=bottom>" & "&#36; " & "<font color=#6699cc><b>" & strPrice & "</b></font><IMG SRC=""images\spacer.gif"" WIDTH=20>" 
	Response.Write  "<input type=hidden name=""Model"" VALUE=" & CHR(34) & objRS("ID") & "^" & strEncName & "^" & strProduct & "^" & strPrice & CHR(34) &">"
	Response.Write strOrder
        Response.Write "</form>"

Else

	arrProduct=Split(strProduct, ",", -1, 1)
	arrPrice=Split(strPrice, ",", -1, 1)

'-------------------------------------------------------------------------
'If there is more than 1 variation of the product print variations in a 
'drop-down box on a seperate row, assigning the value of the product ID 
'number to each choice
'----------------------------------------------------------------------------
	
If UBound(arrProduct)=>1 Then

%>
	<FORM METHOD="POST" ACTION="https://secure6.brinkster.com/blrrfh/cart.asp">	
	<TR><TD COLSPAN=2 align=right valign=bottom>
		
	
		<SELECT NAME=Model>
<% 
	Dim NameEnc, VarEnc
	For strElement=0 to UBound(arrProduct)
	NameEnc = objRS("Name")
	NameEnc = Server.HTMLEncode(NameEnc)
	VarEnc = arrProduct(strElement)
	VarEnc = Server.HTMLEncode(VarEnc)
	
	Response.Write "<OPTION VALUE=" & CHR(34) & objRS("ID") & "^" & NameEnc & "^" & VarEnc & "^" & arrPrice(strElement) & CHR(34) &">"

	Response.Write arrProduct(strElement) & "&#160;" & "&#160;" & "&#36; " & arrPrice(strElement) &  "<BR>"

	Next

%>	
	</SELECT>
	
	<IMG SRC="images\spacer.gif" WIDTH=20 HEIGHT=1>

<%
	Response.Write strOrder

%>
	</form>
		
<%
	Else

'------------------------------------------------------
'If there is one variation and one only, print it on the same row as the image'
'-------------------------------------------------------------------------------
%>
<form method="post" action="https://secure6.brinkster.com/blrrfh/cart.asp">
<%

VarEnc = Server.HTMLEncode(strProduct)

	Response.Write "<tr><TD  COLSPAN=2 align=right valign=bottom><font face=""arial"">" & strProduct & "</font><img src=""images\spacer.gif"" width=20>" & "&#36;" & "<font color=#6699cc><b>" & strPrice & "</b></font><img src=""images\spacer.gif"" width=20>"
	Response.Write "<input type=hidden name=Model VALUE=" & CHR(34) & objRS("ID") & "^" & strEncName & "^" & VarEnc & "^" & strPrice & CHR(34) &">"

	Response.Write strOrder

	Response.Write "</form>"

	End If
End If
%>
	
	<TR><TD COLSPAN=4><IMG SRC="images\spacer.gif" WIDTH=2 HEIGHT=10 ALT=""><HR NOSHADE SIZE=1 WIDTH=80%></TD></TR>
<%

	objRS.MoveNext
	Loop
%>

<%

Else 
%>

<!--If there are no records, print a readable error message-->

<tr><td><FONT FACE="ARIAL" SIZE=4>We're sorry.  There are no items in this catagory yet.  Please check back later!</TD>
</TR></table></td></tr></table>
<%
End if
objRS.close
set objRS=nothing

If pg<>1 then

'--------------------------------------------
'Close product recordset object and connection
'----------------------------------------------

objconn.close
set objConn=nothing
Else
%>

<%

'------------------------------------------------------------------------------
'If page 1 of catagory, print a Featured companion
'------------------------------------------------------------------------------

dim objRSF

Set objRSF = Server.CreateObject("ADODB.Recordset")
objRSF.CursorType=adOpenStatic

objRSF.Open strSQLF, objConn
If not objRSF.EOF then




%>
<TR><TD COLSPAN=2 ALIGN=center>
  

    	<TABLE WIDTH="446" border=0 BGCOLOR="#BDD5EB" CELLPADDING=0 CELLSPACING=0>
		
		  	
		 <tr>
		  <TD BGCOLOR="#474d64"><a name="featured"><IMG SRC="images\spacer.gif" WIDTH=1 HEIGHT=1 ALT=""></a></TD>
		  <TD BGCOLOR=#474D64><IMG SRC="images\spacer.gif" WIDTH=4 HEIGHT=1 ALT=""></TD>
		  <TD BGCOLOR=#474D6F><IMG SRC="images\spacer.gif" WIDTH=442 HEIGHT=1 ALT=""></TD>
		  <TD BGCOLOR=#474D6F><iMG SRC="images\spacer.gif" WIDTH=1 HEIGHT=1 ALT=""></TD>
		</TR>
		<TR>
		  <TD BGCOLOR="#474d64" WIDTH="1"><IMG SRC="images\spacer.gif" WIDTH=1 ALT=""></TD>
		  <TD BGCOLOR=#BDD5EB><IMG SRC="images\spacer.gif" WIDTH=4  ALT=""></TD>
		  <TD VALIGN="middle" ALIGN="center" WIDTH="442">
		<IMG VSPACE=10 HSPACE=10 ALIGN=right SRC="Features/<%=objRSF("Photo")%>" WIDTH=<%=objRSF("Width")%> HEIGHT=<%=objRSF("Height")%> ALT=<%=objRSF("FirstName")%>">
		    <P VALIGN="middle"><FONT COLOR=#000000 FACE="arial"><BR><H3><B><%=objRSF("FirstName")%></B></H3></P><P ALIGN=left VALIGN=middle><%=objRSF("Caption")%></FONT></TD>
		  <TD BGCOLOR=#474D64><IMG SRC="images\spacer.gif" WIDTH=1 ALT=""></td>
		</TR>
		<TR>
		  <TD COLSPAN=4 BGCOLOR=#474D64><IMG SRC="images\spacer.gif" HEIGHT=1></TD>
		<TR>		  
		
		<TR>
		  <TD COLSPAN=4 ALIGN=center BGCOLOR=#FFF0C1><FONT FACE="arial" COLOR=#000000 ><B>Our featured companions already have happy homes, but many
			more are still waiting.  Click <A HREF="http://www.houserabbit.org/AdoptedCurrent.asp">here</A> to visit adoptable rabbits, rodents and ferrets at BLRRFH.</B></FONT></TD>
		</TR>
	</TABLE>
   
	
</TD></TR>

<% end if %>

<!--End featured companion display-->

<TR><TD COLSPAN=2><IMG SRC="images\spacer.gif" HEIGHT=20></TD></TR>

<%
'-------------------------------
'Close featured companion recordset object and connection
'-------------------------------

objRSF.Close
Set objRSF=Nothing
objConn.Close
Set objConn=Nothing
End if


%>
	</TABLE></CENTER></TD></TR>

<tr><td  colspan=2 align=center>
<table border=0><tr><td align=center>
<%

Response.Write "<B><FONT FACE=""arial"" COLOR=#999933 SIZE=2>Page " & pg & " of " & pgcount & "<BR>"
		If pgcount>1 Then
		Response.Write "<b><FONT FACE=""arial"" COLOR=#4486c4>Go to page "		
		 For i=1 to pgcount
	 		If i <> cInt(pg) THEN 
%>

			<A HREF="product.asp?Cat=<%=intCat%>&animal=<%=intAnimal%>&pg=<%=i%>"><FONT COLOR=#ff9900><%=i%></FONT></A>

<% 
			Else 
			Response.Write i

			End if
		Next
		end if

               
%>
</td>
<td>
<a href="#topofpage"><img src="images\home.gif" width=45 height=40 border=0></a>

<a href="https://secure6.brinkster.com/blrrfh/cart.asp"><img src="images\check.gif" border=0 width=45 height=40></a></td>
</tr>
</table>
</td>

</TR>
<tr><td colspan=2><hr size=1 width=100%></td></tr>
<tr><td colspan=2 align=center>

<!--#include file="pgbottom.asp"-->
</td></tr></table>
</BODY>
</HTML>


