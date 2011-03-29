
<html>
<head>
<title>Best Little Rabbit, Rodent and Ferret House - Register</title>
<!--#include file="google-analytics.js"-->
</head>
<body bgcolor="ffffff" LINK=#39638C VLINK=#FF9933 ALINK=#999933>
<center>
<table width=80% border=0 cellpadding=4 cellspacing=0>
<tr><td align=center><img src="images\checkout.gif" width=248 height=122 alt="Register or Log In"><br><img align=left src="images\titlebaska.gif" width=359 height=24></td></tr>
<tr>
  <td bgcolor=#bdd5eb>
  <font  face="Arial" color=#000000>
  <b>Login</b>
  </font>
  </td>
</tr>
<tr>
  <td>

  <form method="post" action="register.asp">

  
  <font face="Arial" size="3" color=#ffa405><b>
  If you've previously registered, just enter your username and password:</b>
  </font>
  <font face="Courier" size="2" color=#39638C> 
  <p><b>username:</b>  
  <input name="username" size="20"></b>
  <br><font face="Courier" size="2" color=#39638C> <b>password:</b>  
  <input type="password" name="password" size="20"></b>
  <input type="submit" value="Login">

  
  </font>
  </form>  

  </td>
</tr>
<tr>
  <td bgcolor=#bdd5eb>
  <font  face="Arial">
  <b>Register</b>
  </font>
  </td>
</tr>
<tr>
  <td>
  
  <form method="post" action="register.asp">
  
  <font face="Arial" size="3" color=#ffa405>
  <b>If you are a new user, please complete the following form:</b><br></font><font size=4 face=arial color=#ffa405>
   
  Fields marked with an <font face=courier color="red" size=2><b>*</b></font> are mandatory.</font>
  
  <p><font face=arial size=2 color=#000000><b>Login Information: </font> <font color="#999933" size=2><b> If you enter a username and password we
   will store your information so you don't have to enter it next time! <b></font><p>
  
<P align=center><font face=arial size=2 color=#39638c>>>We keep all your information completely confidential<<</font></P>
  <font face="Courier" size="2" color=#39638C><br><b>new username:</b>  
  <input name="newusername" size=20 maxlength=20 value="<%=strNewUserName%>"><br>

<font face="arial" size="2" color=#000000> 
  Use alpha and numeric characters only for your username.</font><br>
  
  <br><b>new password:</b>  
  <input  type="password" name="newpassword" value="<%=strNewPassword %>" size=20 maxlength=20>
  <br></font>
<font face="arial" size="2" color=#000000> 
  Use at least six alpha and numeric characters for your password.</font><br>
	
  <font face="Arial" Size=2 color=#000000>
  <p><b>Customer Information:</b>
  </font>

<font face="Courier" size="2" color=#39638C>
  <br><b><font color="red">*</font>name:</b>
  <input name="name" size=20 maxlength=50 value="<%=strCustName %>"></font>

  <font face="Courier" size="2" color=#39638C>
  <br><b><font color="red">*</font>street:</b>
  <input name="street" size=20 maxlength=50 value="<%=strStreet %>">
  <br><b><font color="red">*</font>city:</b>
  <input name="city" size=20 maxlength=50 value="<%=strCity %>">
  <br><b><font color="red">*</font>state:</b>
  
<select name="State">
<Option Selected>
<Option value="AL">AL
<Option value = "AK">AK
<Option>AZ
<Option>AR
<Option>CA
<Option>CO
<Option>CT
<Option>DE
<Option>DC
<Option>FL
<Option>GA
<Option>GU
<Option>HI
<Option>ID
<Option>IL
<Option>IN
<Option>IA
<Option>KS
<Option>KY
<Option>LA
<Option>ME
<Option>MH
<Option>MD
<Option>MA
<Option>MI
<Option>MN
<Option>MS
<Option>MO
<Option>MT
<Option>NE
<Option>NV
<Option>NH
<Option>NJ
<Option>NM
<Option>NY
<Option>NC
<Option>ND
<Option>MP
<Option>OH
<Option>OK
<Option>OR
<Option>PW
<Option>PA
<Option>PR
<Option>RI
<Option>SC
<Option>SD
<Option>TN
<Option>TX
<Option>UT
<Option>VT
<Option>VI
<Option>VA
<Option >WA
<Option>WV
<Option>WI
<Option>WY
</select>

  <br><b><font color="red">*</font>zip:</b>
  <input name="zip" size=20 maxlength=20 value="<%=strZip %>">
  <br>

 <b>telephone:</b>
<input name="phone" size=12 maxlength=20 value="<%=strPhone %>"><br></font>
<font face="Courier" size="2" color=#39638C>
<b><font color="red">*</font>email address:</b>  
  <input name="email" size=30 maxlength=75 value="<%=strEmail%>">
  </font><br>
  <font face="Arial" size="2" color=#000000>
  <p><b>Payment Information:</b>
  </font>
  <font face="Courier" size="2" color=#39638C>
  <br><b><font color="red">*</font>type of credit card:</b>  
  <select name="cctype">
  <option value="VISA"> VISA
  <option value="Mastercard">MasterCard
  </select>

  <br><b><font color="red">*</font>credit card number:</b> 
  <input name="ccnumber" size=20 maxlength=20 value="<%=strccNumber%>">

<font face="Courier" size="2" color=#39638C>
  <br><b><font color="red">*</font>credit card expires:</b> (mm/yy)
  <input name="ccexpires" size=4 maxlength=20 value="<%=strccExp%>">
  </font>
  <input type="submit" value="send">

  
  </form>  

  </td>
</tr>
</table>
<hr noshade width=80%>
<!--#include file="pgbottom.asp"-->

</body>




<% Response.Cookies("mt")("pagetitle") = "" : Server.Execute("/stats/track.asp") %>
</html>