<% @language=vbscript %>
<% option explicit %>

<!--#include file="strconn.asp"-->


<%
response.buffer=true
If not Session("UserLevel")=2 then
	response.redirect("default.asp")

else

dim update
update=Request.form("Edprod")
if update <> "" then
strSQL="SELECT * from Products where ID = " & update

Dim objConn, strSQL
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString=strconnect

objConn.Open
Dim objRS
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, objConn

dim edID, edName, edDescriptionLng, edImage, edWidth, edHeight
dim edCatagory, edRabbits, edRodents, edFerrets, edInOrder, edStatus
dim edVariation1Des, edVariation1Price

Do while not objRS.EOF
edID= objRS("ID")
edName = objRS("Name")
If not isNull(edName) then
edName = Server.HtmlEncode(edName)
end if
edDescriptionLng = objRS("DescriptionLng")
edImage = Trim(objRS("Image"))
edWidth = objRS("Width")
edHeight = objRS("Height")
edCatagory = objRS("Catagory")
edRodents = objRS ("Rodents")
edRabbits = objRS("Rabbits")
edFerrets = objRS("Ferrets")
edInOrder= objRS("InOrder")
edStatus = objRS("Status")
edVariation1Des = objRS("Variation1Des")
edVariation1Price=objRS("Variation1Price")

If edVariation1Des <>"" then
edVariation1Des = Server.HtmlEncode(edVariation1Des)

dim arrDesc
arrDesc=Split(edVariation1Des, ",")
end if

If edVariation1PRice<>"" then

edVariation1Price = objRS("Variation1Price")

dim arrPrice
arrPrice=Split(edVariation1PRice, ",")
end if


objRS.Movenext
Loop

objRS.Close
set objRS = nothing
objConn.Close
set objConn=nothing

end if
%>

<HTML>
<HEAD>
<SCRIPT LANGUAGE="JavaScript">

function empty(val){
   var start=0
   if (val == "" || val==" " || val==undefined || val=="undefined" || val==null){
   start=1;
}
return start;
}

 
function isnumeric(teststring) {
   var len=teststring.length;
   count= 0;
   num = 2;
   
   for (count = 0; count< len; count++) {
   if (teststring.charAt(count) != 0 && teststring.charAt(count)!=1 && teststring.charAt(count)!=2
   && teststring.charAt(count)!=3 && teststring.charAt(count)!=4 && teststring.charAt(count)!=5
   && teststring.charAt(count)!=6 && teststring.charAt(count)!=7 && teststring.charAt(count)!=8
   && teststring.charAt(count)!="9" && teststring.charAt(count)!= "."){ 
    num = 1;   
    
}   
   
  }

return num;
}
function match(vard, varp) {
 var noprice = 0;
if (vard > 0 && varp <= 0) 
noprice = 1;
if (varp != "" && vard == "") 
noprice = 1;
return noprice; 
}


function validate() {

 if (document.newproduct.name.value.length < 1 || document.newproduct.name.value.length > 125) {
 alert("Please enter a product name between 1 and 40 characters.");
 return false;
}

if (document.newproduct.descriptionlng.value.length < 1) {
alert ("Please enter a product description.");
return false;
}

var imgconv=document.newproduct.image.value;
var imgconvl=imgconv.toLowerCase();

  if (empty(document.newproduct.image.value)==0  && imgconvl.indexOf(".jpg")<0 && imgconvl.indexOf(".gif")<0) {
   alert("If you enter an image, it must be a .jpg or .gif");
     return false;
}

  
 

 if (document.newproduct.rabbits.checked !=1 && document.newproduct.rodents.checked !=1 
     && document.newproduct.ferrets.checked !=1) {
  alert ("Product should be displayed in at least one animal catagory (rabbits, rodents or ferrets)");
  return false;
}


if (document.newproduct.varp1.value.length <=0) {
alert ("You must enter at least one price.");
return false;
}
 vari=new Array(10);
vari[0]=document.newproduct.varp1.value;
vari[1]=document.newproduct.varp2.value;
vari[2]=document.newproduct.varp3.value;
vari[3]=document.newproduct.varp4.value;
vari[4]=document.newproduct.varp5.value;
vari[5]=document.newproduct.varp6.value;
vari[6]=document.newproduct.varp7.value;
vari[7]=document.newproduct.varp8.value;
vari[8]=document.newproduct.varp9.value;
vari[9]=document.newproduct.varp10.value;

 vards=new Array(10);
vards[0]=document.newproduct.vard1.value;
vards[1]=document.newproduct.vard2.value;
vards[2]=document.newproduct.vard3.value;
vards[3]=document.newproduct.vard4.value;
vards[4]=document.newproduct.vard5.value;
vards[5]=document.newproduct.vard6.value;
vards[6]=document.newproduct.vard7.value;
vards[7]=document.newproduct.vard8.value;
vards[8]=document.newproduct.vard9.value;
vards[9]=document.newproduct.vard10.value;

var i=1
for (i=1; i<10; i++) {


if (empty(vards[i]) == 0 && empty(vari[i]) == 1 ) {
alert ("There must be a price for each description");
return false;
}
}

i=1;

for (i=1; i<10; i++) {

if (empty(vari[i]) == 0 && empty(vards[i]) == 1 ) {
alert ("There must be a description for each price, unless there is only one price");
return false;
}
}

i=1;
for (i=1; i<10; i++) {
if (empty(vards[0])==1 && empty(vards[i])==0) {
alert ("There must be a description for each price, unless there is only one price");
return false;
}
}
i=0;
for (i=0; i<10; i++) {


if (vards[i].indexOf(",") >=0) {

alert ("Don't use commas in your descriptions.");
return false;
}
}

i=0;
for (i=0; i<10; i++){

if (empty(vari[i])==0) {
 
  if (isnumeric(vari[i]) == 1){
  alert ("Don't use letters or symbols in the price of Variation " + (i+1) );  
return false;
}
}
}
i=0;
for (i=0; i<10; i++) {

if (empty(vari[i])==0) {
  
len=vari[i].length;
  if (vari[i].charAt(len-3) != "."){
  alert ("Enter price of Variation " + (i+1) + " with 2 decimal places:  x.xx");
  return false;
   }
}

}
  return true;
}

function preview() {
 
var display = "startvalue";
if (document.newproduct.image.value.length > 0 && document.newproduct.vard2.value.length > 0) {
 display = "isimgmultvar";
}
if (document.newproduct.image.value.length > 0 && document.newproduct.vard2.value.length <= 0) {
 display = "isimgsingvar";
}

if (document.newproduct.image.value.length <= 0 && document.newproduct.vard2.value.length > 0) {
display = "noimgmultvar";
}

if (document.newproduct.image.value.length <=0 && document.newproduct.vard2.value.length <= 0) {
display = "noimgsingvar";
}



DispWin = window.open('','Newwin','toolbar=no,status=no,width=640,height=500')

switch(display) {
  case "isimgmultvar" :

   
 	message="<html><body><table BORDER=0 width=95%><tr><TD valign=middle width= " + document.newproduct.width.value + ">";     
        message+= "<img width=" + document.newproduct.width.value;
        message+= " height = " + document.newproduct.height.value + " src=http://www.rabbitrodentferret.org/products/" + document.newproduct.image.value + "></td>";
   	message+= "<td align=left><FONT FACE=arial SIZE=4 ><b>" + document.newproduct.name.value + "</b></font><font face=arial size=2><p>" + document.newproduct.descriptionlng.value + "</p></td></tr>";
        message+= "<tr><form><td colspan=2 align=right valign=bottom><select>"; 
        message+= "<option>" + document.newproduct.vard1.value + "  " + document.newproduct.varp1.value + "<br>";
        message+= "<option>" + document.newproduct.vard2.value + "  " + document.newproduct.varp2.value + "<br>";
	message+= "<option>" + document.newproduct.vard3.value + "  " + document.newproduct.varp3.value + "<br>";
	message+= "<option>" + document.newproduct.vard4.value + "  " + document.newproduct.varp4.value + "<br>";
	message+= "<option>" + document.newproduct.vard5.value + "  " + document.newproduct.varp5.value + "<br>";
	message+= "<option>" + document.newproduct.vard6.value + "  " + document.newproduct.varp6.value + "<br>";
	message+= "<option>" + document.newproduct.vard7.value + "  " + document.newproduct.varp7.value + "<br>";
        message+= "<option>" + document.newproduct.vard8.value + "  " + document.newproduct.varp8.value + "<br>";
	message+= "<option>" + document.newproduct.vard9.value + "  " + document.newproduct.varp9.value + "<br>";
	message+= "<option>" + document.newproduct.vard10.value + "  " + document.newproduct.varp10.value + "<br>";
	message+= "</select><img src='http://www.rabbitrodentferret.org/images/spacer.gif' width=20 height=1><font color=#999933 face=arial>Quantity:</font>";
	message+= "<input type=text size=2 value=1><input type=submit value='Add to Basket'></form></td></tr>";
	message+= "<tr><td colspan= 2 align=center><hr noshade width=60% size=2></td></tr></table>";
  	message+= "<p><font face=arial size=3>Th above is a preview of how this product will display on the site once it is submitted ";
	message+= " to the database.  If you are satisfied with it, close this window and click the 'Submit' button on the Add new Product page.</p>";
	message+= "<p> Otherwise, go back to the Add New Product page and make appropriate changes.</font></p>";
break;
  case "isimgsingvar" :
     	message="<html><body><table BORDER=0 width=95%><tr><TD valign=middle width= " + document.newproduct.width.value + ">";     
        message+= "<img width=" + document.newproduct.width.value;
        message+= " height = " + document.newproduct.height.value + " src=http://www.rabbitrodentferret.org/products/" + document.newproduct.image.value + "></td>";
   	message+= "<td align=left><FONT FACE=arial SIZE=4 ><b>" + document.newproduct.name.value + "</b></font><font face=arial size=2><p>" + document.newproduct.descriptionlng.value + "</p></td></tr>";
        message+= "<tr><form><td colspan=2 align=right valign=bottom>";
        message+=  document.newproduct.vard1.value + "<img src='http://www.rabbitrodentferret.org/images/spacer.gif' width=20 height=1><font face=arial color=#6699cc><b>&#36; " + document.newproduct.varp1.value ;
        message+= "</b></font><img src='http://www.rabbitrodentferret.org/images/spacer.gif' width=20 height=1><font color=#999933 face=arial>Quantity: </font>";
        message+= "<input type=text size=2 value=1><input type=submit value='Add to Basket'></form></td></tr>";
	message+= "<tr><td colspan=2 align=center><hr noshade width=60% size=2></td></tr></table>";
        message+= "<p><font face=arial size=3>The above is a preview of how this product will display on the site once it is submitted ";
	message+= " to the database.  If you are satisfied with it, close this window and click the 'Submit' button on the Add new Product page.</p>";
	message+= "<p> Otherwise, go back to the Add New Product page and make appropriate changes.</font></p>";
  break;

  case "noimgmultvar" :

	message="<html><body><table BORDER=0 width=95%><tr>";
   	message+= "<td align=left colspan=2><FONT FACE=arial SIZE=4 ><b>" + document.newproduct.name.value + "</b></font><font face=arial size=2><p>" + document.newproduct.descriptionlng.value + "</p></td></tr>";
        message+= "<tr><form><td colspan=2 align=right valign=bottom><select>"; 
        message+= "<option>" + document.newproduct.vard1.value + "  " + document.newproduct.varp1.value + "<br>";
        message+= "<option>" + document.newproduct.vard2.value + "  " + document.newproduct.varp2.value + "<br>";
	message+= "<option>" + document.newproduct.vard3.value + "  " + document.newproduct.varp3.value + "<br>";
	message+= "<option>" + document.newproduct.vard4.value + "  " + document.newproduct.varp4.value + "<br>";
	message+= "<option>" + document.newproduct.vard5.value + "  " + document.newproduct.varp5.value + "<br>";
	message+= "<option>" + document.newproduct.vard6.value + "  " + document.newproduct.varp6.value + "<br>";
	message+= "<option>" + document.newproduct.vard7.value + "  " + document.newproduct.varp7.value + "<br>";
        message+= "<option>" + document.newproduct.vard8.value + "  " + document.newproduct.varp8.value + "<br>";
	message+= "<option>" + document.newproduct.vard9.value + "  " + document.newproduct.varp9.value + "<br>";
	message+= "<option>" + document.newproduct.vard10.value + "  " + document.newproduct.varp10.value + "<br>";
	message+= "</select><img src='http://www.rabbitrodentferret.org/images/spacer.gif' width=20 height=1><font color=#999933 face=arial>Quantity:</font>";
	message+= "<input type=text size=2 value=1><input type=submit value='Add to Basket'></form></td></tr>";
	message+= "<tr><td colspan= 2 align=center><hr noshade width=60% size=2></td></tr></table>";
  	message+= "<p><font face=arial size=3>The above is a preview of how this product will display on the site once it is submitted ";
	message+= " to the database.  If you are satisfied with it, close this window and click the 'Submit' button on the Add new Product page.</p>";
	message+= "<p> Otherwise, go back to the Add New Product page and make appropriate changes.</font></p>";

 break;

 case "noimgsingvar" :

	message="<html><body><table BORDER=0 width=95%><tr>";
   	message+= "<td align=left colspan=2><FONT FACE=arial SIZE=4 ><b>" + document.newproduct.name.value + "</b></font><font face=arial size=2><p>" + document.newproduct.descriptionlng.value + "</p></td></tr>";
        message+= "<tr><form><td colspan=2 align=right valign=bottom>";
        message+=  document.newproduct.vard1.value + "<img src='http://www.rabbitrodentferret.org/images/spacer.gif' width=20 height=1><font face=arial color=#6699cc><b>&#36; " + document.newproduct.varp1.value ;
        message+= "</b></font><img src='http://www.rabbitrodentferret.org/images/spacer.gif' width=20 height=1><font color=#999933 face=arial>Quantity: </font>";
        message+= "<input type=text size=2 value=1><input type=submit value='Add to Basket'></form></td></tr>";
	message+= "<tr><td colspan=2 align=center><hr noshade width=60% size=2></td></tr></table>";
        message+= "<p><font face=arial size=3>The above is a preview of how this product will display on the site once it is submitted ";
	message+= " to the database.  If you are satisfied with it, close this window and click the 'Submit' button on the Add new Product page.</p><p>";
	message+= "Otherwise, go back to the Add New Product page and make appropriate changes.</font></p>";
      	
 break;
 default:
     message="No preview available";

}
    DispWin.document.write(message);
}
</script>
</head>
<body >
<font face="arial">
<form name = "newproduct" method="post" action="addproduct2.asp" onSubmit="return validate();">

<center><font face=arial size=5>

<%
if update <> "" then
Response.Write "Edit a Product"
else
Response.Write "Add a Product"
end if
%>
</font><hr width=40% size=1 noshade></center>
<center><font color="red">* denotes a required field</font></center><p>
<center><font face=arial size=4 color=#999933><u>Name and Description of Product</u></font></center><p>


<font color="red"><b>*</b> </font><font face=arial><b>Name</b> of Product - must be between 1 and 40 characters<br> <input type="text" size="40" name="name" value="<%=edName%>"><br>
<br><font color="red">* </font>Product <b>Description</b> - You may use html tags in this section like &#60; b &#62; and &#60;/b &#62;<br>
<textarea rows = 5 cols = 40 name ="descriptionlng" size="50"><%=edDescriptionLng%></textarea><br><hr><p>

<center><font face=arial size = 4 color=#999933><u>Image Information <font size=3>(skip if none available)</font></u></font></center><p>

<b>Image</b> Filename: <input type="text" size="20" name="image" value="<%=edImage%>">( <i>filename</i>.jpg or <i>filename</i>.gif)<br>  
<b>(Don't forget to upload the actual image to the images folder under webroot in the site directory)</b>.<br><br>
Image <b>Width:</b> <input type="text" size="5" name="width" value="<%=edWidth%>"> pixels<br>
Image <b>Height:</b> <input type="text" size="5" name="height" value="<%=edHeight%>"> pixels<br>
Width and Height must be between 10 and 200 pixels<hr>

<p><center><font face=arial size = 4 color=#999933><u>Display Information </font></u></font></center><p>


<b><font color="red">* </font>Category:</b>
 <select name="catagory">
<%
dim selected, i, catname

for i=1 to 9
selected=""
if edCatagory=i then
selected="selected"
end if

 Select Case i

 Case "1"
  catname="Toys"
 Case "2"
  catname="Housing"
 Case "3"
  catname="Furnishings"
 Case "4"
  catname="Food"
 Case "5"
  catname="Grooming"
 Case "6"
  catname="Healthcare"
 Case "7"
  catname="People Stuff"
 Case "8"
  catname="Books & Videos"
 Case "9"
  catname = "Misc"
 End Select
Response.Write "<option value =" & i & " " & selected & ">" & catname
next
%>
</select><p>

<%
dim checked
checked=""
If edRabbits=1 then
checked = "checked"
end if 
Response.Write "<input type=checkbox name=rabbits value=1 " & checked & ">Rabbits?<br>"

checked=""
If edRodents=1 then
checked = "checked"
end if 
Response.Write "<input type=checkbox name=rodents value=1 " & checked & ">Rodents?<br>"

checked=""
If edFerrets=1 then
checked = "checked"
end if 
Response.Write "<input type=checkbox name=ferrets value=1 " & checked & ">Ferrets?<p>"


%>
<b>Order</b> of Display (enter a high number if you want the product displayed first (i.e. a product with a "5" will diplay before a product with a "1")<br>
If no number is entered products will be displayed with the most recent entries first.<br>
<input type="text" name="inorder" size="2" value="<%=edInOrder%>"><br><br>

<%

checked=""
if edStatus=1 or update="" then
checked = "checked"
end if
Response.Write "<input type=checkbox name=status value=1 " & checked & "><b>Display Immediately</b> (Uncheck if you don't want this product displayed yet)<p>"
%>
<hr>
<center><font face=arial size = 4 color=#999933><u>Variations and Pricing</font></u></center><p>


Product <b>Variations</b> (i.e. small, medium, large)<br><br>
<b>List each variation in order on a separate line and don't use commas.<br>List price in format x.xx or xx.xx without a dollar sign.</b><br>
<font color=red>You must list at least 1 price.</font>  <br>

<table border=0 ><tr><td>
<%
If isarray(arrDesc) then
For i = 0 to Ubound(arrDesc)
Response.Write "<font face=arial>Variation " & (i+1) & "  Description <input type=text name=vard" & (i+1) & " size=20 Value="""
%>
<%=arrDesc(i)%>
<% 
Response.Write """></font><br>"
Next

For i=(UBound(arrDesc)+1) to 9
Response.Write "<font face=arial>Variation " & (i+1) & "  Description <input type=text name=vard" & (i+1) & " size=20></font><br>"
Next 

Else
for i=0 to 9
Response.Write "<font face=arial>Variation " & (i+1) & "  Description <input type=text name=vard" & (i+1) & " size=20></font><br>"
Next 
end if
                                          
%>
</td><td>

<%
If isarray(arrPrice) then
For i = 0 to Ubound(arrPrice)
Response.Write "<font face=arial>Variation " & (i+1) & "  Price <input type=text name=varp" & (i+1) & " size=5 Value="""
%>
<%=arrPrice(i)%>
<% 
Response.Write """></font><br>"
Next

For i=(UBound(arrPrice)+1) to 9
Response.Write "<font face=arial>Variation " & (i+1) & "  Price <input type=text name=varp" & (i+1) & " size=5></font><br>"
Next 

Else
for i=0 to 9
Response.Write "<font face=arial>Variation " & (i+1) & "  Price <input type=text name=varp" & (i+1) & " size=5></font><br>"
Next 
end if
                                          
%>
</td></tr>
</table>
 <hr>
<p>

<input type="button" value="Preview"  onClick="preview();">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
<input type="submit" value="Submit" onClick="return confirm('You have submitted the product to the database. Are you sure?')">
<input type="hidden" name="update" value="<%=update%>">
</form>
<form action="updatep.asp">
<input type="submit" value="Cancel" onClick="return confirm('Cancel this addition or edit?')">
</form>
<%
If update <> "" then
%>
<form method=post action="addproduct2.asp">

<input type= "submit" value="Delete this product" onClick="return confirm('You are about to delete this product.  Are you sure?')">
<input type = "hidden" name = update value="<%=update%>">
<input type = "hidden" name = delete value="delete">
</form>

<%
end if
%>
</body>
</html>
<%
end if
%>
