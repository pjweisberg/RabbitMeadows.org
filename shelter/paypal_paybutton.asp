<%
Function paypal_paybutton(item, price)
    Response.write "<form action='https://www.paypal.com/cgi-bin/webscr' method='post'>"
    Response.write "<input type='hidden' name='cmd' value='_xclick'/>"
    Response.write "<input type='hidden' name='business' value='Sandi@RabbitRodentFerret.org'/>"
    Response.write "<input type='hidden' name='no_shipping' value='1'/>"
    Response.write "<input type='hidden' name='shipping' value='0.00'/>"
    Response.write "<input type='hidden' name='tax' value='0'/>"
    Response.write "<input type='hidden' name='return' value='http://www.rabbitmeadows.org/shelter/'/>"
    Response.write "<input type='hidden' name='cancel_return' value='http://www.rabbitmeadows.org/shelter/'/>"
    Response.write "<input type='image' src='http://images.paypal.com/images/x-click-but02.gif' name='submit' alt='PayPal - it&apos;s fast, free and secure!'/>"

    Response.write "<input type='hidden' name='item_name' value='" + item + "'>"
    Response.write "<input type='hidden' name='amount' value='" + price + "'>"
    Response.write "<input type='hidden' name='undefined_quantity' value='1'>"
    Response.write "</form>"
End Function
%>
