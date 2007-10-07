<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

sFormAction = "designfaq.asp"
sName = "Design FAQ"
sTitle = "Design FAQ"
thisRedirect = "designfaq.asp"
sMenu="Design"
createHead thisRedirect

%>
<B>Where are my normal design menus?</b><BR>
A new improved design interface has now been added in place of your old one. The usual design menu links such as Header and Footer, Color and Text, Action Buttons are now part of your template and can be found on the design layout page of your template. To go to the layout page of a template, go to Design->Custom Templates->View/Edit Templates, click the Edit link for a template to go to its layout page.

<BR><BR><B>Will this affect my current store design?</b><BR>
No, it will not affect your current store design which will continue to function as it always has. 


<BR><BR><B>How do I create a new template?</b><BR>
To create a new template, go to Design->Custom Templates->Add Template. You can now add more than one template to your store and customize each one of them to your liking. You can either use the simple object creator from the template layout page to create standard design objects for your template or you can enter your own custom html containing the common design elements of your design in the Header and Footer page. 


<BR><BR><B>How do I add design objects to the layout page?</b><BR>
click the Add link next to each design area on the template layout page to start adding standard design objects to your template. 


<BR><BR><B>Can I still use the Theme Manager to apply a standard template?</b><BR>
Yes, you can continue to use the theme manager like before from Design->Theme Manager. Moreover, you can now completely customize a standard template. When you apply a standard theme to your store, a new template is created in your template list. You can modify this newly created theme from its layout page. Go to Design->Custom Templates->View/Edit Templates to see a list of your templates.


<BR><BR><B>Can I preview a template before applying it?</b><BR>
Yes, you can now preview every template that you create before you apply it to your store to make sure you've got everything exactly as you wanted.
To preview a template, go to its design layout page. You can view a list of your templates from Design->Custom Templates->View/Edit Templates


<BR><BR><B>How do I save my template so that I dont mess it up while making changes?</b><BR>
Create a copy of your template from its layout page and work on the copy while making changes so that you can switch back to the original to start over again in case you mess up the template copy. You can now also Preview a template from its layout page before applying it. To go the layout page of a template, go to Design->Custom Templates->View/Edit Templates


<BR><BR><B>What are the top, down arrows on the layout page for?</b><BR>
They help you to move the position of an object in a design area. Click on a design object or click the select link next to a design object to select it, you can then move its position in the design area by using the top and bottom arrow icons. Click on the save icon next to the arrows to save the new position of the objects in that area.


<BR><BR><B>Can I move a design object from one area to another?</b><BR>
Yes, click on the Edit link next to an object to go to its Edit screen. Then change the design area field to the new area where you want to move the object to.


<BR><BR><B>How do I customize the design areas?</b><BR>
Click on the Edit link next to each design area to modify its properties.


<BR><BR><B>I have made changes to my template but they arent taking effect on my storefront?</b><BR>
All changes made to your template dont take effect until you apply the template to your store. To apply a template to your store, go to its layout page.
You can preview your changes though by Previewing a template from its layout page


<BR><BR><B>How do I go to the layout page of a template?</b><BR>
Go to Design->Custom Templates->View/Edit Templates, click the Edit link for a template to go to its layout page.

<%
createFoot thisRedirect, 0%>
