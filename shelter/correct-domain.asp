<%
    dim hostname
    hostname = Request("HTTP_HOST")
    hostname = lcase(hostname)
    if not (hostname="www.rabbitmeadows.org" or hostname="localhost" or hostname="127.0.0.1" or hostname="192.168.1.101") then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader "Location", "http://www.rabbitmeadows.org"&Request("URL")
        Response.End
    end if
%>
