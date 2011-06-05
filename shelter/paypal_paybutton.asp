<%
Function paypal_common()
    Response.write "<form action='https://www.paypal.com/cgi-bin/webscr' method='post'>"
    Response.write "<input type='hidden' name='business' value='Sandi@RabbitRodentFerret.org'/>"
    Response.write "<input type='hidden' name='no_shipping' value='1'/>"
    Response.write "<input type='hidden' name='shipping' value='0.00'/>"
    Response.write "<input type='hidden' name='tax' value='0'/>"
    Response.write "<input type='hidden' name='return' value='http://www.rabbitmeadows.org/shelter/'/>"
    Response.write "<input type='hidden' name='cancel_return' value='http://www.rabbitmeadows.org/shelter/'/>"
    
End Function

Function paypal_paybutton(item, price)
    paypal_common

    Response.write "<input type='hidden' name='cmd' value='_xclick'/>"
    Response.write "<input type='image' src='http://images.paypal.com/images/x-click-but02.gif' name='submit' alt='PayPal - it&apos;s fast, free and secure!'/>"
    Response.write "<input type='hidden' name='item_name' value='" + item + "'>"
    Response.write "<input type='hidden' name='amount' value='" + price + "'>"
    Response.write "<input type='hidden' name='undefined_quantity' value='1'>"
    Response.write "</form>"
End Function

Function paypal_donatebutton()
    paypal_common

    Response.write "<input type='hidden' name='cmd' value='_donations'/>"
    Response.write "<input type='image' src='http://images.paypal.com/images/x-click-but04.gif' name='submit' alt='Donate with PayPal - it&apos;s fast, free and secure!'/>"
    Response.write "</form>"
End Function

Function paypal_multi(item, choices)
    paypal_common

    Response.write "<input type='hidden' name='cmd' value='_xclick'>"
    Response.write "<input type='hidden' name='undefined_quantity' value='1'>"
    Response.write "<input type='hidden' name='item_name' value='" + item + "'>"
    Response.write "<input type='hidden' name='on0' value='" + item + "'>" 
    Response.write "<table>"
    Response.write "<tr><td>"
    Response.write "<input type='image' src='http://images.paypal.com/images/x-click-but02.gif' name='submit' alt='PayPal - it&apos;s fast, free and secure!'/>"
    Response.write "</td><td>"
    Response.write item + ": "
    Response.write "<select name='os0'>"

    Dim i
    i = 0
    While i < UBound(choices)
        Response.write "<option value ='" + choices(i) + "'>"
        Response.write choices(i) + " "
        i = i + 1
        Response.write "$" + CStr(choices(i))
        Response.write "</option>"
        i = i + 1
    Wend
    Response.write "</select>"
    Response.write "</td></tr>"
    Response.write "</table>"

    i = 0
    While i < UBound(choices)
        Response.write "<input type='hidden' name='option_select" + CStr(fix(i/2)) + "' value='" + choices(i) + "'>"
        i = i + 1
        Response.write "<input type='hidden' name='option_amount" + CStr(fix(i/2)) + "' value='" + CStr(choices(i)) + "'>"
        i = i + 1
    Wend

    Response.write "<input type='hidden' name='option_index' value='0'>"
    
    Response.write "</form>"
End Function
%>