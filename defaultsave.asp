<% @LANGUAGE=VBScript %>
<% Option Explicit %>



<HTML><head></head>
<body>
<%
response.write "test"
dim address
address=request.servervariables("HTTP_HOST")

response.write address
%>
</body></html>


<%

address=lcase(address)
If address="www.rabbitrodentferret.org" then

response.redirect "http://www.rabbitrodentferret.org/rabbitrodentferret.org/index.asp"
Elseif address="www.washingtonhouserabbitsociety.org" then
response.redirect "http://www.rabbitrodentferret.org/rabbitrodentferret.org/index.asp"
Elseif address="www.barbaradeeb.org" then
response.redirect "http://BarbaraDeeb.org/BarbaraDeeb.org/index.html"
Elseif address=("www.deerbrookhaven.org") then
response.redirect "http://www.rabbitrodentferret.org/deerbrookhaven.org/default.htm"
Elseif address=("BarbaraDeebScholarshipFund.org") then
response.redirect "\BarbaraDeeb.org\default.htm"
Elseif address=("BarbaraDeebVeterinaryScholarship.org") then
response.redirect "\BarbaraDeeb.org\default.htm"'elseif address=("deerbrookhaven.org") then
response.redirect "\deerbrookhaven.org\default.htm"
Elseif address=("hrabbit.brinkster.net") then
Response.Redirect "\rabbitrodentferret.org\product.asp"
Elseif address=("www.rabbitmeadowssanctuary.org") then
Response.Redirect "http://www.rabbitrodentferret.org/RabbitMeadowsSanctuary.asp"
Elseif address=("www.rabbitmeadows.org") then
Response.Redirect "http://www.rabbitrodentferret.org/rabbitrodentferret.org/rabbitmeadowssanctuary.asp"
end if
%>