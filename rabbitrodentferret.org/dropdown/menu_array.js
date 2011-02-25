/*
 Milonic DHTML Menu
 Written by Andy Woolley
 Copyright 2002 (c) Milonic Solutions. All Rights Reserved.
 Plase vist http://www.milonic.co.uk/menu or e-mail menu3@milonic.com
 You may use this menu on your web site free of charge as long as you place prominent links to http://www.milonic.co.uk/menu and
 your inform us of your intentions with your URL AND ALL copyright notices remain in place in all files including your home page
 Comercial support contracts are available on request if you cannot comply with the above rules.

 Please note that major changes to this file have been made and is not compatible with versions 3.0 or earlier.

 You no longer need to number your menus as in previous versions. 
 The new menu structure allows you to name the menu instead, this means that you can remove menus and not break the system.
 The structure should also be much easier to modify, add & remove menus and menu items.
 
 If you are having difficulty with the menu please read the FAQ at http://www.milonic.co.uk/menu/faq.php before contacting us.

 Please note that the above text CAN be erased if you wish as long as copyright notices remain in place.
*/

//The following line is critical for menu operation, and MUST APPEAR ONLY ONCE. If you have more than one menu_array.js file rem out this line in subsequent files
menunum=0;menus=new Array();_d=document;function addmenu(){menunum++;menus[menunum]=menu;}function dumpmenus(){mt="<script language=javascript>";for(a=1;a<menus.length;a++){mt+=" menu"+a+"=menus["+a+"];"}mt+="<\/script>";_d.write(mt)}
//Please leave the above line intact. The above also needs to be enabled if it not already enabled unless this file is part of a multi pack.



////////////////////////////////////
// Editable properties START here //
////////////////////////////////////

// Special effect string for IE5.5 or above please visit http://www.milonic.co.uk/menu/filters_sample.php for more filters
if(navigator.appVersion.indexOf("MSIE 6.0")>0)
{
	effect = "Fade(duration=0.2);Alpha(style=0,opacity=88);Shadow(color='#777777', Direction=135, Strength=5)"
}
else
{
	effect = "Shadow(color='#777777', Direction=135, Strength=5)" // Stop IE5.5 bug when using more than one filter
}


timegap=500				// The time delay for menus to remain visible
followspeed=5			// Follow Scrolling speed
followrate=40			// Follow Scrolling Rate
suboffset_top=10;		// Sub menu offset Top position 
suboffset_left=10;		// Sub menu offset Left position

style1=[				// style1 is an array of properties. You can have as many property arrays as you need. This means that menus can have their own style.
"navy",					// Mouse Off Font Color
"cce6ff",				// Mouse Off Background Color
"ffebdc",				// Mouse On Font Color
"4b0082",				// Mouse On Background Color
"000000",				// Menu Border Color 
11,						// Font Size in pixels
"normal",				// Font Style (italic or normal)
"normal",					// Font Weight (bold or normal)
"Verdana, Arial",		// Font Name
4,						// Menu Item Padding
"dscript98/arrow.gif",			// Sub Menu Image (Leave this blank if not needed)
,						// 3D Border & Separator bar
"666666",				// 3D High Color
"000099",				// 3D Low Color
"Purple",				// Current Page Item Font Color (leave this blank to disable)
"pink",					// Current Page Item Background Color (leave this blank to disable)
"dscript98/arrowdn.gif",			// Top Bar image (Leave this blank to disable)
"ffffff",				// Menu Header Font Color (Leave blank if headers are not needed)
"000099",				// Menu Header Background Color (Leave blank if headers are not needed)
]



addmenu(menu=[		// This is the array that contains your menu properties and details
"mainmenu",			// Menu Name - This is needed in order for the menu to be called
1,					// Menu Top - The Top position of the menu in pixels
1,				// Menu Left - The Left position of the menu in pixels
,					// Menu Width - Menus width in pixels
1,					// Menu Border Width 
,					// Screen Position - here you can use "center;left;right;middle;top;bottom" or a combination of "center:middle"
style1,				// Properties Array - this is set higher up, as above
1,					// Always Visible - allows the menu item to be visible at all time (1=on/0=off)
"left",				// Alignment - sets the menu elements text alignment, values valid here are: left, right or center
effect,				// Filter - Text variable for setting transitional effects on menu activation - see above for more info
,					// Follow Scrolling - Tells the menu item to follow the user down the screen (visible at all times) (1=on/0=off)
1, 					// Horizontal Menu - Tells the menu to become horizontal instead of top to bottom style (1=on/0=off)
0,					// Keep Alive - Keeps the menu visible until the user moves over another menu or clicks elsewhere on the page (1=on/0=off)
,					// Position of TOP sub image left:center:right
,					// Set the Overall Width of Horizontal Menu to 100% and height to the specified amount (Leave blank to disable)
,					// Right To Left - Used in Hebrew for example. (1=on/0=off)
,					// Open the Menus OnClick - leave blank for OnMouseover (1=on/0=off)
,					// ID of the div you want to hide on MouseOver (useful for hiding form elements)
,					// Reserved for future use
,					// Reserved for future use
,					// Reserved for future use

,"SimplytheBest.net","http://simplythebest.net",,"Back to the home page",1 // "Description Text", "URL", "Alternate URL", "Status", "Separator Bar"
,"Directories&nbsp;&nbsp;","show-menu=directories",,"",1
,"Network&nbsp;&nbsp;","show-menu=network",,"",1
,"PlanMagic software&nbsp;&nbsp;","show-menu=downloads",,"",1
,"Search&nbsp;&nbsp;","show-menu=search",,"",1
,"Other&nbsp;&nbsp;","show-menu=other",,"",1
])

	addmenu(menu=["directories",
	,,185,1,"",style1,,"left",effect,,,,,,,,,,,,
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Shareware & Freeware","show-menu=shareware",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;DHTML scripts","show-menu=dhtml",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;CGI scripts","show-menu=cgi",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Information","show-menu=info",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Device drivers","http://simplythebest.net/drivers/",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Affiliate programs","http://simplythebest.net/affiliates/",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Music & Musicians","http://simplythebest.net/music/",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Free Fonts","http://simplythebest.net/fonts/",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Hosting providers","http://simplythebest.net/hosting/",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Free Sounds","http://simplythebest.net/sounds/",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Metasearch","http://web.simplythebest.net/",,,1
	])
	
		addmenu(menu=["shareware",
		,,110,1,"",style1,,"left",effect,,,,,,,,,,,,
		,"Business","http://simplythebest.net/shareware/business/",,,0
		,"Graphics","http://simplythebest.net/shareware/graphics/",,,0
		,"Programming","http://simplythebest.net/shareware/programming/",,,0
		,"Utilities","http://simplythebest.net/shareware/utilities/",,,0
		,"Internet tools","http://simplythebest.net/shareware/web_utilities/",,,0
		,"Hobby","http://simplythebest.net/shareware/hobby/",,,0
		])

		addmenu(menu=["dhtml",
		,,130,1,"",style1,,"left",,,,,,,,,,,,,
		,"Animation","http://simplythebest.net/scripts/DHTML_scripts/dhtml_animation.html",,,0
		,"Background","http://simplythebest.net/scripts/DHTML_scripts/dhtml_background.html",,,0
		,"Buttons","http://simplythebest.net/scripts/DHTML_scripts/dhtml_buttons.html",,,0
		,"Calculators","http://simplythebest.net/scripts/DHTML_scripts/dhtml_calculators.html",,,0
		,"Cookies","http://simplythebest.net/scripts/DHTML_scripts/dhtml_cookies.html",,,0
		,"Enhancements","http://simplythebest.net/scripts/DHTML_scripts/dhtml_enhancements.html",,,0
		,"Forms","http://simplythebest.net/scripts/DHTML_scripts/dhtml_forms.html",,,0
		,"Image rotations","http://simplythebest.net/scripts/DHTML_scripts/dhtml_images.html",,,0
		,"Menus","http://simplythebest.net/scripts/DHTML_scripts/dhtml_menu_scripts.html",,,0
		,"Messages","http://simplythebest.net/scripts/DHTML_scripts/dhtml_messages.html",,,0
		,"Password protection","http://simplythebest.net/scripts/DHTML_scripts/dhtml_passwords.html",,,0
		,"Scrollers","http://simplythebest.net/scripts/DHTML_scripts/dhtml_scroller_scripts.html",,,0
		])

		addmenu(menu=["cgi",
		,,140,1,"",style1,,"left",effect,,,,,,,,,,,,
		,"Access counters","http://simplythebest.net/scripts/perl_scripts/counter_scripts.html",,,0
		,"Ad rotation scripts","http://simplythebest.net/scripts/perl_scripts/ad_banner_scripts.html",,,0
		,"Auction scripts","http://simplythebest.net/scripts/perl_scripts/auction_scripts.html",,,0
		,"Bulletin Boards","http://simplythebest.net/scripts/perl_scripts/bulletin_board_scripts.html",,,0
		,"Chat scripts","http://simplythebest.net/scripts/perl_scripts/chat_scripts.html",,,0
		,"Databases","http://simplythebest.net/scripts/perl_scripts/database_scripts.html",,,0
		,"File downloading","http://simplythebest.net/scripts/perl_scripts/download_scripts.html",,,0
		,"Form processing","http://simplythebest.net/scripts/perl_scripts/form_scripts.html",,,0
		,"Greeting card service","http://simplythebest.net/scripts/perl_scripts/greetingcard_scripts.html",,,0
		,"Guestbooks","http://simplythebest.net/scripts/perl_scripts/guestbook_scripts.html",,,0
		,"Mailing list scripts","http://simplythebest.net/scripts/perl_scripts/mailing_scripts.html",,,0
		,"Password protection","http://simplythebest.net/scripts/perl_scripts/password_scripts.html",,,0
		,"Searching scripts","http://simplythebest.net/scripts/perl_scripts/search_scripts.html",,,0
		,"Security scripts","http://simplythebest.net/scripts/perl_scripts/security_scripts.html",,,0
		,"Shopping carts","http://simplythebest.net/scripts/perl_scripts/shopping_cart_scripts.html",,,0
		,"Surveys, rate, vote","http://simplythebest.net/scripts/perl_scripts/survey_scripts.html",,,0
		])

		addmenu(menu=["info",
		,,130,1,"",style1,,"left",effect,,,,,,,,,,,,
		,"3D graphics","http://simplythebest.net/info/design/",,,0
		,"Anti-virus","http://simplythebest.net/info/virus_info/",,,0
		,"ASP","http://simplythebest.net/info/programming/",,,0
		,"ColdFusion","http://simplythebest.net/info/programming/",,,0
		,"Databases","http://simplythebest.net/info/programming/",,,0
		,"Delphi","http://simplythebest.net/info/programming/",,,0
		,"E-commerce","http://simplythebest.net/info/internet/",,,0		
		,"Encryption","http://simplythebest.net/info/security/",,,0
		,"Java","http://simplythebest.net/info/programming/",,,0
		,"Linux/Unix","http://simplythebest.net/info/unix/",,,0
		,"Metatags generator","http://simplythebest.net/info/design/",,,0
		,"MP3","http://simplythebest.net/info/sound/",,,0
		,"MySQL","http://simplythebest.net/info/programming/",,,0
		,"Networks","http://simplythebest.net/info/pc/",,,0
		,"PERL","http://simplythebest.net/info/programming/",,,0
		,"PHP","http://simplythebest.net/info/programming/",,,0
		,"Servers","http://simplythebest.net/info/servers/",,,0
		,"Spyware","http://simplythebest.net/info/spyware/",,,0
		,"Visual Basic","http://simplythebest.net/info/programming/",,,0
		,"XML","http://simplythebest.net/info/programming/",,,0
		])

	addmenu(menu=["network",
	,,130,1,"",style1,,"left",effect,,,,,,,,,,,,
	,"Advertising","http://simplythebest.net/sponsors.htm",,,1
	,"Email us","mailto:email@simplythebest.net",,,1
	,"Link to us","http://simplythebest.net/link2us.html",,,1
	,"Sitemap","http://simplythebest.net/sitemap.html",,,1
	,"Visitor links","http://simplythebest.net/ouraward.html",,,1
	])
	
	addmenu(menu=["downloads",,,150,1,,style1,0,"left",effect,0,,,,,,,,,,,
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Business plan","http://planmagic.com/business_planning.html",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Marketing plan","http://planmagic.com/marketing_planning.html",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Financial plan","http://planmagic.com/financial_planning.html",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Hotel plan","http://planmagic.com/business_plan/hotel_business_plan.html",,,1
	,"<img src=dscript98/newsimage.gif border=0>&nbsp;Restaurant plan","http://plan-a-restaurant.com",,,1
	])
	
	addmenu(menu=["search",
	,,120,1,"",style1,,"",effect,,,,,,,,,,,,
	,"<img src=dscript98/google_icon.gif border=0>&nbsp;Google.com", "http://www.google.com",,,1
	,"<img src=dscript98/yahoo_icon.gif border=0>&nbsp;Yahoo", "http://www.yahoo.com",,,1
	,"<img src=dscript98/av_icon.gif border=0>&nbsp;Altavista", "http://altavista.com",,,1
	,"<img src=dscript98/excite.gif border=0>&nbsp;Excite", "http://www.excite.com",,,1
	])
	

	addmenu(menu=["Other",
	,,125,1,,style1,0,"left","randomdissolve(duration=0.5);Shadow(color='#999999', Direction=135, Strength=5)",0,,,,,,,,,,,
	,"Menu Authors Site","http://www.milonic.co.uk/menu",,"Milonic DHTML Menu Site",1
	])
	

dumpmenus()