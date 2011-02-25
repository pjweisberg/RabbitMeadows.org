<html>
<head><title>HouseRabbit.org Sendmail</title>
<meta NAME="Keywords" CONTENT="live online House Rabbit, Rabbitmedia,
	online Rabbits, streaming live videos,Adoptions, Rabbit Facts, Rabbit Sex, Rabbit news,
	online media ">
<meta NAME="Description" CONTENT="HouseRabbit.org is a the definitive site for Recued Rabbits in the 
Northwest and other parts of the country. HRS is a non-provit organization.">
<meta NAME="Robots" CONTENT="index all">
<script>
	function lite(whichObj,name){
		whichObj.src = "./images/" + name + "_hi.gif";
	}

	function deLite(){
		var tmpObj;
		for(var i = 0; i < 5; i++){
			tmpObj = eval('document.btn' + i);	
			if (tmpObj.src.indexOf("_hi") > -1) {
				tmpObj.src = tmpObj.src.substring(0,tmpObj.src.indexOf("_hi")) + ".gif";
			}
		}
	}
   </script>

<style>
	a{color:#396bb5; text-decoration:none; font-weight:bold}
	a:visited{color:#555555}
	a:hover{text-decoration:underline; }
   </style>
</head>
<body>
<center>
<h1>Sendmail For HouseRabbit</h1>
<hr>
<%
If request("sendmail")="" then
	Call Display()
else
	If instr(request("from"),"@")=0 then
		response.write "Error - incorrect email"
		Call display()
	else
		If instr(request("from"),".")=0 then
			response.write "Error - incorrect email"
			Call display()
		else
			'go ahead and send it
			sql="Insert into email values('" & replace(request("from"),"'","''") & "')"
			'response.write sql
			'--------------------------------------------------------------------
			' Open the connection
			'--------------------------------------------------------------------
			Set Conn = Server.CreateObject("ADODB.Connection")
			sConnect = "dsn=houseRabbit"
			Conn.Open sConnect
			conn.execute sql
			'--------------------------------------------------------------------
			' Close everything
			'--------------------------------------------------------------------
			Conn.close
			set Conn=nothing
			'--------------------------------------------------------------------
			' Send the email
			'--------------------------------------------------------------------
			set sm = Server.CreateObject ("MPS.Sendmail")
			feedback=sm.SendMail (request("from"), "info@houserabbit.org",request("subject"),request("message"))
			set sm=nothing
			response.write "Message Sent, Thank You !"
		end if
	end if
end if


sub Display()
%>
<form action="sendmail.asp" method="POST">
<table>
<tr>
<td>From</td><td><input type="textbox" name="from" width="10" value="<%=request("from")%>"></td>
</tr>
<td>Type of email</td>
<td><select name="type">

<%
If request("type")="Adopt" then
	sSelect=" SELECTED"
else
	sSelect=""
end if
%>
<option value="Adopt" <%=sSelected%>>Adopt</option>
<%
If request("type")="supplies" then
	sSelect=" SELECTED"
else
	sSelect=""
end if
%>
<option value="supplies" <%=sSelected%>>Supplies</option>
<option value="Help" <%=sSelected%>>Help</option>
<%
If request("type")="Help" then
	sSelect=" SELECTED"
else
	sSelect=""
end if
%>
<option value="FeedBack" <%=sSelected%>>FeedBack</option>
<%
If request("type")="FeedBack" then
	sSelect=" SELECTED"
else
	sSelect=""
end if
%>
<option value="Membership" <%=sSelected%>>Membership</option>
<%
If request("type")="Membership" then
	sSelect=" SELECTED"
else
	sSelect=""
end if
%>
</select></td>
</tr>
<tr>
<td>Subject</td><td><input type="textbox" name="subject" width="10" value="<%=request("subject")%>"></td>
</tr>
<tr>
<td>Message</td><td><textarea name="message" cols="50" rows="10"><%=request("message")%></textarea></td>
</td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" name="sendmail" value="sendmail"></td>
</tr>
</table>

</form>
<%
end sub
%>
</body>
</html>

