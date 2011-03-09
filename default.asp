<% @LANGUAGE=VBScript %>
<% Option Explicit %>



<html><head></head>
<body>
<%
response.write "test "
dim address
address=request.servervariables("HTTP_HOST")

response.write address
%>
</body></html>


<%

address=lcase(address)
If address="www.barbaradeeb.org" or address="barbaradeeb.org" then
    response.redirect "http://BarbaraDeeb.org/BarbaraDeeb.org/index.html"
Elseif address=("www.rabbitmeadowssanctuary.org") then
    Response.Redirect "http://www.rabbitmeadows.org/shelter/RabbitMeadowsSanctuary.asp"
Else
    Response.Redirect "http://www.rabbitmeadows.org/shelter/"
end if
%>