<!--#include file="Global_Settings.asp"-->
<!--#include file="pagedesign.asp"-->
<%

sFormAction = "Inventory_action.asp"
sName = "Store_Activation"
sFormName = "features"
sTitle = "Features, Changes and Additions"
thisRedirect = "features.asp"
sMenu = "none"
createHead thisRedirect
%>
<tr bgcolor='#FFFFFF'><TD><B>10/6/2007</b><BR><UL>
<LI>UPS and Fedex realtime rates can now be enabled at the same time
<LI>Fixed issed with USPS military address rates at times with UPS enabled</LI>
<LI>Added state dropdowns for United States and Canadian addresses in store setup to prevent incorrect entries
<LI>Added ability to send a test newsletter to a single email address</LI>
<li>Added item layout variable OBJ_FRIEND_URL_OBJ which will return the URL of the send to a friend page</li>
</UL></TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>8/19/2007</b><BR><UL>
<LI>Enabled <a href=real_time_settings.asp class=link>Google Checkout</a> as a payment method
<LI>Added heading of invoice # or order # at the top of printable invoices
<LI>Added passing of Fedex account number for fedex rates to allow for individually negotated rates
<LI>Modified store Log Out page in stores to display a message to user that they are logged out.
<LI>
</UL></TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>8/19/2007</b><BR><UL>
<LI>Color picker modified to work with Firefox
<LI><a href=google_testers.asp class=link>Google Checkout Beta Testing</a> enabled
</UL></TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>8/1/2007</b><BR><UL>
<LI>Added ability to use variables in <a href=emails.asp class=link>email notifications</a> subjects
<LI>Modified PayPal refunded status to mark order as credited instead of not verified
<LI>Added ability to modify <a href=emails.asp class=link>newsletter signup email</a> subject
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>7/5/2007</b><BR><UL>
<LI><a href=shipping_class_realtime.asp class=link>Realtime Shipping</a> API Updated, Canada Post support for new URL, Conway Support Removed, USPS fetching latest rates
<LI>Added ability to <a href=company.asp class=link>subscribe/unsubscribe</a> to internal StoreSecured newsletter.  First edition to be sent next week.
<LI>SmarterMail upgraded to version 4.3
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>6/16/2007</b><BR><UL>
<LI>Updated database backend tables
<LI>Modified checkout process to take a snapshot of shopping cart so that if cart is modified during checkout process user must restart checkout to get new shipping, tax values etc
<LI>Updated quantity discounts to take effect even when different attributes of an item are purchased
<LI>Added ability to add tracking numbers and capture funds on multiple orders at a time
<LI>Added ability to set a minimum order value.  If set above 0 shoppers will be unable to checkout if their purchase does not meet the minimum.
<LI>SmarterMail upgraded to version 4.2
<LI>Removed internally generated references to the old image folder path, "images/images_storeid", these folders are now referenced as "images"
<LI>Added color indicators for different order statuses on order details page
<LI>Added referring url and client ip for orders if available
<LI>Added ability to see abandoned orders from an earlier stage in the checkout process.  These orders however cannot be manually completed as no shipping, tax, etc has been chosen yet.
<LI>Removed known robots from list of shoppers
<LI>Opened up newsletter emails again to be sent to yahoo addresses
<LI>Added new page site_map_dept.asp which is a list of departments only
<LI>Added ability to not apply template to pages.  Pages without the template applied will still use the default style.
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>6/1/2007</b><BR><UL>
<LI>Added 2 new variables to the item layout OBJ_IMAGES_URL_OBJ and OBJ_IMAGEL_URL_OBJ.  These variables can be used to just get the path of the image and nothing else, for instance to use for making popups etc.
<LI>Modified item mass edit attribute page to pull and show all settings for the current attributes
<LI>Google Base item feeds will now be sent once a week instead of daily.  Recently Google has been removing existing items when a new feed is submitted and then taking up to 24 hours to repost them.  This change should ensure your items stay published.  Cap on number of items sent has been removed however, stores with over 3000 items will have their feed sent monthly instead of weekly.
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>5/14/2007</b><BR><UL>
<LI>SmarterMail version 4.1 installed, webmail interface updated
<LI>Greylisting will be enabled on the mailserver in 48 hours.  Greylisting is very effective for stopping spam however if your emails are extremely time sensitive you may need to turn off this feature as it does delay emails.  To turn off greylisting login to smartermail as the domain admin and choose Domain Settings-->Default User Settings and check the box to Bypass Greylisting.  With greylisting turned on emails will be delayed slightly.<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>5/14/2007</b><BR><UL>
<LI>Added ability to mark a calculated shipping method only for use when realtime rates are not returned.</LI></UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>5/6/2007</b><BR><UL>
<LI>Modified template header/footer page to separate template head from template body section for easier editing purposes.  Actual display of templates will be unchanged.</LI></UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>3/22/2007</b><BR><UL>
<LI>Modified realtime shipping restricted options.  Will now match if the text is in the shipping method name.  Ie Flat Rate Envelope in the restricted options list would remove USPS Flat Rate Envelope Small and USPS Flat Rate Envelope Large.  Please ensure with the new looser matching that you are not restricting options you did not intend to.</LI></UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>3/3/2007</b><BR><UL>
<LI>Added ability to create up to 5 navigation link and button menus with different content</LI>
<LI>Separated department nav menus so that they can be turned on or off for each top level department</LI>
<li>Modified coupons/gift certificate/rewards to allow all to be used at the same time</li>
<li>Modified coupon/gift certificate entry, add apply button to apply to order</li>
<li>Gift certificates may now be used to pay for shipping and taxes</li>
<li>Synched all invoice and cart views to be consistent across the store pages</li>
<li>Added downloadable file link for electronic products in email invoice attachment</li>
<li>Added ability to rename pages once they are created</li>
<li>Modified store design creation and pages to show updates immediately after making changes</li>
<li>Moved nav layout menu option inside of template edit</li>
<li>Modified all prices to internally store to 2 decimal places</li>
<li>Modified automatic nav buttons to center text inside the button</li>
<LI>Added new variables to expose the html head information for advanced users in the template header/footer
<LI>Email notification will now include a link to download the file.  This link requires that the user have created a login and password to verify identity.
<LI>Modified Bluepay gateway data to send billing data for avs instead of shipping data, per Bluepay
<LI>Internal code and database enhancements to optimize page actions</LI>

</UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>12/1/2006</b><BR><UL>
<LI>Items that are not in a department will not be shown in store or site map.  All items must have a department.
<LI>Search results, homepage, and item accessories will use the featured items layout
<LI>Department views will use the small item layout
<LI>The item accessories keyword will now show the whole item featured layout and not just the item name
<LI>Modified department urls to use the syntax http://www.domainname.com/items/departmentname/list.htm
<LI>Modified item urls to use the syntax http://www.domainname.com/items/department/itemfilename-detail.htm
<LI>Added auto-generated Google sitemap at http://www.domainname.com/auto_sitemap.xml
<LI>Added auto-generated Product Rss feed at http://www.domainname.com/products.xml
<LI>Modified search results to display on new page, search_items.asp
<LI>Added search engine friendly links to additional pages on paged departments
<LI>Modified item edit pages to include a single department selector.  You may choose as many departments as you would like but there is no longer the concept of main and additional departments.
<LI>Old department and item urls will redirect to the new urls to ensure smooth transition
<LI>Default auto-generated links for the homepage will now use only the domain name and will not include the store.asp as part of the url
<LI>Added item filename field to item edit pages, field accepts only the letters a-z, numbers 0-9, the dash (-) and the underscore (_)
<LI>Added Preview Dept and Preview Item buttons to top of respective edit pages
<LI>Updated statistics program to capture individual names for different items
<LI>Import Items Format Updated, please see the new format before uploading items
<LI>Modified item and department pages to automatically use item name and item description for meta title tag and meta description tag if no separate meta title and description is entered
</UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>10/23/2006</b><BR><UL>
<LI>Modified Item/Department Add/Edit to have basic and advanced mode (for existing users, use Advanced Edit/Add options to continue with the old screens)
<LI>Added popup html editor which can be used when editor is hidden for those who prefer to only use the editor sometimes, setting the editor to Hide will in many cases load the page quicker however using the Editor will take an extra step
<LI>Set extended item fields 1-5 to always use popup editor regardless of settings to increase load speed
<LI>Added showing of active menu in red color for clarity
<LI>Added navigation breadcrumbs in top right corner to easily backtrack
<LI>Added name of thing being edited, saved, added to button description for clarity
<LI>Added Site # request at login for increased security
<LI>When using Item Basic Add, merchant cannot continue if no departments exist, a department must be created first
<LI>Increased font size of menu text and page title text for easier viewing
<LI>Switched login page to secure mode by default
<LI>Modified generated 3rd party add to cart code to include attributes, configurations and user defined fields
<LI>Increased size of fields on admin customer edit and add screens
<LI>Removed text qualifier option on import
</UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>8/29/2006</b><BR><UL>
<LI>Updated coupons and promotions, added ability to exclude or include items
<LI>Speeded up processing for coupons and promotions
<LI>Limited promotions to 4000 characters for performance
<LI>Added ability to limit access to pages and departments by customer group
<LI>Upgraded email spam control system
<LI>Made best seller report sortable on quantity
<LI>Added automatic unsubscribe link to emails sent to customer groups
<LI>Added 301 Redirect if store is using multiple domain names so that 1 is always the primary
</UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>5/21/2006</b><BR><UL>
<LI>Modified search functionality to better find matches, search departments, etc, order of search terms does not matter
<LI>Added additional keywords for department layout, ie count of items and departments
<LI>Added additional site servers, updated existing site servers
<LI>Updated Smarter Mail web interface
<LI>Updated HTML Editor to newest release
<LI>Modified affiliate program to use subtotals not grand totals
</UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>3/28/2006</b><BR><UL>
<LI>Added ability to customize 404 page not found page
<LI>Added ability to export newsletter subscribers
<LI>Added detailed agreement and terms on cancellation page
<LI>Added handling to check for invalid store homepage url to prevent endless loops
<LI>Updated printed help manuals and webhelp for new admin interface
<LI>Updated Smarter Mail web interface
<LI>Added more advanced email virus / spam filtering
<LI>Modified realtime gateway page to show only selected gateway instead of all
<LI>Added ability to view record of logins to store management area

</UL>
<BR>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>3/2/2006</b><BR><UL>
<LI>Modified design of admin interface, new look and feel
<LI>Modified navigation menus, moved around items to more easily find options
<LI>Fixed problem on resend email notification where first/last name was not showing
<LI>Added new html editor onscreen instead of as popup window
<LI>Added popups to indicate when export shipping/export departments are completed
<LI>Modified page manager view to be sortable
<LI>Modified payment methods view to be sortable and allow additional payment methods to have their own custom messages
<LI>Added a way for store owners to remove old domain names themselves
<LI>Added callback service for live help
<LI>Modified department view to be sortable/non tree structure for easier finding/editing
<LI>Added ability to create notes for orders that are internal or external
<LI>Updated quickbooks export functionality
<LI>Added tag functionality to built in editors to allow for easy insertion of custom keywords
<LI>Modified thumbnail creation to create thumbnails for only those images in the current page list to prevent timeouts
<LI>Added ability to turn on/off html editor
<LI>Updated knowledgebase for new interface
</UL>
<BR>
<B>Known Issue</b>: Saving does not work for some users initially as their browser is not picking up the changes
<BR><B>Solution</b>: Hit F5 on the page you are having trouble with, this will force your browser to pickup the most recent changes, this needs to be done only once

</TD></TR>

<tr bgcolor='#FFFFFF'><TD><B>1/21/2006</b><BR><UL>
<LI>Modified <a href=froogle_settings.asp class=link>Froogle</a> page to show creation date for Froogle file link
<LI>Modified <a href=shipping_all_list.asp class=link>Shipping List</a> to show all shipping methods of different types in single list
<LI>Modified shipping edit screen to allow customer to switch a shipping detail from one method to another
<LI>Moved remaining shipping settings from misc view to <a href=shipping_class_realtime.asp class=link>shipping settings page</a>
<LI>Added detailed logging of ip address and username for users logging in to manage sites for StoreSecured internal staff
<LI>Fixed problem on order edit screen.  If shipping method name had comma, shipping details previously could not be edited.
<LI>Modified view cart page to allow customer to make changes to the cart and hit any button to update the changes, versus only hitting the Update cart button</UL>
</UL></TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>1/9/2006</b><BR><UL>
<LI>Modified <a href=orders.asp class=link>orders page</a> to include link to view orders on Order Id column and separate Customer Id column
<LI>Modified <a href=my_customer_base.asp class=link>customer page</a> to include email address and link to email the customer
<LI>Modified date format on both of the above to take up less space
<LI>Modified <a href=edit_items.asp class=link>item search</a> to show item id instead of visible status
</UL>
</TD></TR>
<tr bgcolor='#FFFFFF'><TD><B>1/1/2006</b><BR><UL>
<LI>New secure certificate logo on secure payment pages
<LI>Added more user friendly shipping errors
<LI>Added more user friendly Paypal errors
<LI>Added more user friendly Authorize.net errors
<LI>Added javascript checking for shipping method chosen before customer proceeds
<LI>Added auto-select of shipping method if only 1 exists
<LI>Created separate state dropdowns for Canada and US
<LI>Modified label on state and zip fields for addresses outside the US
<LI>Added admin ability to edit the state for an order already placed
<LI>Corrected problem with searching for returned orders showing no results
<LI>Modified invoice display to be cleaner, matched past orders page to receipt page
<LI>Added ability to resend order email notification from admin screen
<LI>Added minimum order weight of .1 for realtime shipping to stop errors on 0 weight
<LI>Modified 404 error page to display 404 status for google sitemaps users
<LI>Added escape sequence for email on contact us page to prevent harvesting by spiders
<LI>Use lookups in include files for departments, countries, locations to speed up loading
<LI>Added storage of verification code for paypal transaction
<LI>Modified rewards to relookup amount just before cash in
<LI>Added PST designation to order notification emails
<LI>Fixed problem which occurred on purchase of the same gift certificate multiple times in the same order
<LI>Added store url to receipts
<LI>Removed fax and phone information from receipts when empty
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>11/10/2005</b><BR><UL>
<LI>Added ability to configure your own <a href=payment_methods.asp class=link>custom payment methods names</a>
<LI>Added ability to <a href=misc_settings.asp class=link>turn off shipping</a> entirely for a store so no shipping choices are ever presented to user
<LI>Added ability to create a <a href=customize_invoice.asp class=link>custom header and footer</a> for the store invoices
<LI>Modified customer registration pages to change state and zip code field labels for other countries to more appropriate designations
<LI>Modified bluepay gateway to use newer API version vs post and post back version for quicker and more secure execution
<LI>Removed total shoppers line from <a href=reports_totals1.asp class=link>traffic and totals report</a> in favor of viewing number of visitors from reports-->statistics
<LI>Added information on <a href=page_manager.asp class=link>page edit screen</a> to show what kind of default content is already placed on a page that cannot be changed
<LI>Added filename column to <a href=page_manager.asp class=link>page list</a> screen
<LI>Added ability to designate a <a href=shipping_class3_list.asp class=link>base shipping fee</a> to be added onto per item shipping methods
<LI>Added the ability to edit all pages from the <a href=page_manager.asp class=link>page manager</a>, previously some were excluded
<LI>Modified wording on <a href=message_customers.asp class=link>send newsletter</a> screen to indicate emails are only sent to customers requesting promo emails
<LI>Modified <a href=item_edit.asp class=link>item add</a> to return to added item in edit mode when saved
<LI>Added the ability to <a href=template_list.asp class=link>apply the template</a> from any of the template editing screens
<LI>Added <a href=my_customer_base.asp class=link>group picker</a> on the customer edit screen to allow assigning a customer directly into a group
<LI>Added new <a href=admin_home.asp class=link>announcements area</a> on homepage
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>9/19/2005</b><BR><UL>
<LI>Optimized department browse pages and item detail pages to load quicker even as number of products in a department increases
<LI>Modified edit items, orders and customer pages with new search mechanism to search on almost any field
<LI>Added jump to page links on all list views where number of records exceeds current page
<LI>Decreased maximum rows per page in list displays from 1000 to 500 
<LI>Removed export checkboxes in favor of export buttons (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>8/22/2005</b><BR><UL>
<LI>New <a href=pearl.asp class=link>Pearl Site Builder Plan</a>
<LI>Added color previews to screens with color selections
<LI>Added <a href=payment_methods.asp class=link>Purchase order payment</a> option (BRONZE+)
<LI>Added ability to define a weight range to <a href=shipping_class2_list.asp class=link>base fee plus weight fee shipping</a> (BRONZE+)
<LI>Send <a href=emails.asp class=link>automatic email notifications</a> as html or ascii text (BRONZE+)
<LI>Added <a href=fields.asp class=link>checkbox</a> to additional field types which can be added at checkout (SILVER+)
<LI>Added <a href=fields_page.asp class=link>form builder</a> to create forms (SILVER+)
<LI>Added ability to turn on/off ip tracking for store (SILVER+)
<LI>Modified out of stock message to not show how many out of stock (SILVER+)
<LI>Added backordered message on invoices with out of stock items (SILVER+)
<LI>Removed quantity update for items with quantity control disabled (SILVER+)
<LI>Added ability to include <a href=misc_settings.asp class=link>thumbnail in shopping cart</a> and invoice view (SILVER+)
<LI>Click on imported files to view them directly onscreen  (GOLD+)
<LI>Added check for sheet1 when using Excel files (GOLD+)
<LI>Moved import definition to new link instead of onscreen (GOLD+)
<LI>Removed Fedex realtime tracking, dhl rates due to compatibility issues with fedex/dhl upgrade
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>7/15/2005</b><BR><UL>
<LI>Added ability to <a href=misc_settings.asp class=link>remove retail price</a> from shopping cart and invoices
<LI>Added <a href=customers_groups.asp class=link>customer picker</a> to customer group
<LI>Added ability to <a href=my_customer_base.asp class=link>mass or quick edit customers</a>
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>6/25/2005</b><BR><UL>
<LI>Added <a href=orders.asp class=link>order pick list</a> to pick inventory for multiple orders at once
<LI>Added ability to sort <a href=reports_best_seller.asp class=link>best seller report</a> by column headings
<LI>Added ability for customer to mark whether shipping address is residential or business for lower rates
<LI>Added card code whats this link for customers on checkout screen
<LI>Added indicators on <a href=edit_items.asp class=link>edit item screen</a> for whether attributes, accessories etc exist
<LI>Added ability for customers to enable or disable promotional email settings themselves
<LI>Added ability to store <a href=my_customer_base.asp class=link>tax exempt number</a> for customers
<LI>Added ability to view <a href=total_space.asp class=link>total space used</a> by store
<LI>Added <a href=left_image_picker.asp class=link>paging on image list</a> to display only so many images per page and prevent timeouts.
<LI>Added ability to <a href=orders.asp class=link>edit order coupon amount</a>
<LI>Added ability to <a href=reports_totals1.asp class=link>excluded ip addresses</a> to be from reports-->traffic and totals
<LI>Added ability to <a href=page_links_manager.asp class=link>define a link to a url</a> in the nav links and buttons (SILVER+)
<LI>Added ability to <a href=page_manager.asp class=link>protect affiliate signup</a> page (SILVER+)
<LI>Added ability to <a href=page_manager.asp class=link>copy pages</a> (SILVER+)
<LI>Added handling to reverify quantity in stock before final purchase completion (SILVER+)
<LI>Added ability to toggle on/off item detail <a href=misc_settings.asp class=link>prev/next/up buttons</a> (SILVER+)
<LI>Added ability to put <a href=misc_settings.asp class=link>featured items above or below</a> sub departments (SILVER+)
<LI>Added brand new <a href=template_list.asp class=link>template design interface</a> (SILVER+)
<LI>Added the ability to <a href=template_list.asp class=link>save different templates and preview</a> before applying (SILVER+)
<LI>Added ability to <a href=import_newsletter.asp class=link>import newsletter subscribers</a> from text file (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>5/12/2005</b><BR><UL>
<LI>Added retail price field to cart and invoice screens
<LI>Added ability to add attachments to <a href=support_list.asp class=link>support requests</a>
<LI>Added tabs to certain screens for easier display
<LI>Added ability to <a href=my_customer_base.asp class=link>search for customers</a> with/without purchases
<LI>Modified <a href=domain.asp class=link>domain name</a> screen to allow for secondary name (BRONZE+)
<LI>Added ability to specify different prices by <a href=edit_items.asp class=link>customer group</a> for each item (GOLD+)
<LI>Fixed display problems for certain browsers on <a href=reward_settings.asp class=link>rewards</a>/<a href=affiliate_settings.asp class=link>affiliates</a> screens (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>3/25/2005</b><BR><UL>
<LI>Added help link directly from each setup page to the corresponding manual page
<LI>Added ability to <a href=misc_settings.asp class=link>auto hide out of stock inventory items</a> automatically (SILVER+)
<LI>Added ability to give new <a href=new_page.asp class=link>custom pages</a> a custom name (SILVER+)
<LI>Added ability to edit the <a href=edit_items.asp class=link>item price matrix</a> (SILVER+)
<LI>Added <a href=import_department.asp class=link>department import/export</a> (GOLD+)
</UL></td></tr>
<tr bgcolor='#FFFFFF'><TD><B>3/4/2005</b><BR><UL>
<LI>Right aligned all currency text in stores
<LI>Added a way to <a href=left_image_picker.asp class=link>move images</a> to different folders from web based form
<LI>Modified <a href=left_image_picker.asp class=link>image preview</a> to fit better into frame for large image names
<LI>Added new admin menu bar
<LI>Modified store admin design
<LI>Added rewards keyword for <a href=message_customers.asp class=link>mass email</a> (SILVER+)
<LI>Modified newsletter signup to stop duplicate emails (SILVER+)
</UL></td></tr>
<tr bgcolor='#FFFFFF'><TD><B>2/23/2005</b><BR><UL>
<LI>Added <a href=real_time_settings.asp class=link>checksbynet gateway</a> support
<LI>Upgraded <a href=real_time_settings.asp class=link>bluepay gateway</a> to version 2.0
<LI>Changed default <a href=orders.asp class=link>view on orders screen</a> to show verified unshipped orders
<LI>Added first name and last name variables for use in <a href=designer_template.asp class=link>header/footer</a> (SILVER+)
<LI>Modified <a href=my_customer_base.asp class=link>customers list view</a> to show budget, rewards and promo emails (SILVER+)
<LI>Added promotion text to line items where promotional pricing is applied (SILVER+)
<LI>Added quantity in stock to <a href=edit_items.asp class=link>item edit list</a> screen (SILVER+)
<LI>Modified <a href=my_customer_base.asp class=link>customer export</a> format to match the import format(GOLD+)
<LI>Added promotional email column to <a href=import_customers.asp class=link>customer import</a> (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>1/25/2005</b><BR><UL>
<LI>Added <a href=precharge_settings.asp class=link>precharge guaranteed credit card</a> payments integration
<LI>Added ability to <a href=orders.asp class=link>print multiple invoices</a>
<LI><a href=orders.asp class=link>Auto capture, void and credit</a> for authorize.net, linkpoint and plugnpay
<LI>Added ability to <a href=left_image_picker.asp class=link>delete images</a> with special characters
<LI>When <a href=edit_items.asp class=link>searching for items in admin</a>, search through additional depts
<LI>Added free <a href=siteanalysis.asp class=link>search engine site analysis</a>
<LI>Added select state as the default user registration state to require customers to choose an option
<LI>Added select shipping as the default shipping choice to require customers to choose an option
<LI>Modified width of receipt page to better fit in most templates
<LI><a href=import_shipping.asp class=link>Import and export shipping</a> from text file (GOLD+)
<LI>Modified <a href=edit_items.asp class=link>item mass edit and quick edit</a> to allow changing all fields (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>1/6/2005</b><BR><UL>
<LI>Added <a href=http://reseller.storesecured.com target=_blank class=link>rebranded reseller program</a>
<LI>Fixed problem with page content updates behind a firewall
<LI>Added enhanced email filtering for spam (BRONZE+)
<LI>Upgraded webmail interface (BRONZE+)
<LI>Added whitelists for email (BRONZE+)
<LI>Added ability to <a href=countries_select.asp class=link>limit countries</a> in pull down lists (BRONZE+)
<LI>Added <a href=edit_items.asp class=link>5 extended data fields</a> for displaying misc item data (SILVER+)
<LI>Added fields to <a href=import_items.asp class=link>item import</a> and export (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>11/20/2004</b><BR><UL>
<LI>Added <a href=real_time_settings.asp class=link>Propay Gateway Integration</a>
<LI>Changed store admin page titles to reflect page displayed to make it easier to go backward in browser history
<LI>Added shipment tracking utility to all stores for UPS, USPS, Fedex, DHL, Canada Post
<LI>Added ability to <a class=link href=newemail.asp>add/edit/delete email accounts</a> (BRONZE+)
<LI>Added ability to include ascii <a href=emails.asp class=link>invoice details</a> in order notification body (BRONZE+)
<LI>Added ability to limit <a href=inventory_display.asp class=link>display of rows</a> on featured items pages (SILVER+)
<LI>Add ability to upload Item Ship from location with <a href=import_items.asp class=link>item import</a> (GOLD+)
<LI>Added ability to customer budget modifications, ability to use %OBJ_BUDGET_OBJ% keyword in template (GOLD+)
<LI>Added ability to print ups shipping labels from the order shipping screen (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>10/14/2004</b><BR><UL>
<LI>Added new popup windows for Admin menus
<LI>Added <a href=real_time_settings.asp class=link>Xor payment gateway</a> integration
<LI>Added popup sku chooser on both <a href=coupon_manager.asp class=link>coupon</a> and <a href=promotion_manager.asp class=link>promotion</a> pages (GOLD+)
<LI>Added ability to create <a href=fields.asp class=link>custom fields</a> at checkout (GOLD+)
<LI>Added ability to upload <a href=banners_list.asp class=link>banners for affiliates</a> to use (GOLD+)
<LI>Modified <a href=coupon_manager.asp class=link>coupons</a> to allow a coupon to only be valid for a certain purchase total</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>9/30/2004</b><BR><UL>
<LI>Select pdf files/images etc to link to automatically from internal editor link in html editor
<LI>Add ability to create link back to items from view cart to external web pages
<LI>Added form to get information for new sample stores
<LI>Export items now includes column headers (GOLD+)
<LI>Export items now can export more then 1 department at a time and advanced searching available (GOLD+)
<LI>Import items directly from excel file no need to convert to csv or tab delimited file (GOLD+)
<LI>Updated froogle feed to allow for separate username and text file (GOLD+)
<LI>Realtime shipping added support for Airborne, DHL, Canada Post, Conway Freight, new setup options for Fedex, UPS (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>9/6/2004</b><BR><UL>
<LI>Auto <a href=left_image_picker.asp class=link>thumbnail images</a>
<LI>Live online chat with Easystorecreator operators
<LI><a href=real_time_settings.asp class=link>Cybersource payment gateway</a> integration
<LI>Added automatic alt tags to item display pages
<LI>Added ability to define <a href=insurance_class.asp class=link>insurance rates</a>
<LI>Ability to add <a href=misc_settings.asp class=link>flat handling weight</a> (SILVER+)
<LI>Added ability to <a href=store_dept.asp class=link>remove department name</a> (SILVER+)
<LI>Ability to <a href=store_dept.asp class=link>protect departments</a> (SILVER+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>8/12/2004</b><BR><UL>
<LI>Added http compression algorithm to speed up pages
<LI>Server hardware upgrade
<LI>Added text free form state box for countries other then US and Canada
<LI>Increased maximum theoretical cart order quantity
<LI>Added ability to waive shipping costs for individual inventory items
<LI>Added date to <a href=orders.asp class=link>order</a> and <a href=my_customer_base.asp class=link>customers</a> reports
<LI>Modified <a href=admin_home.asp class=link>admin homepage</a> to have more links directly to other areas
<LI>Gave each store its own root directory (BRONZE+)
<LI>Added <a href=item_layout.asp class=link>item layout</a> codes for SKU and ID (SILVER+)
<LI>Added auto-email to store admin when a rma request is generated (SILVER+)
<LI>Added ability to <a href=import_items.asp class=link>import item</a> additional item fields including accessories, order, meta tags, etc (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>7/21/2004</b><BR><UL>
<LI>Added ability to copy existing inventory item into <a href=edit_items.asp class=link>new item</a>
<LI>Added progess bar to <a href=upload_images.asp class=link>image upload</a> page
<LI>Added eftnet <a href=real_time_settings.asp class=link>payment gateway integration</a>
<LI>Added ip tracking to Easystorecreator <a href=billing.asp class=link>payment pages</a>
<LI>Added new flash and windows media file tutorials to various pages. 
<LI>Added ability to <a href=import_items.asp class=link>import/export inventory items</a> view order, meta keywords and meta description (GOLD+)
<LI>Added ability to <a href=my_customer_base.asp class=link>export customer</a> promotional email setting, budget, reward, taxable, etc (GOLD+)
</UL></td></tr>
<tr bgcolor='#FFFFFF'><TD><B>6/26/2004</b><BR><UL>
<LI>Added date popup windows for date selection on reports
<LI>Added client side form checking for stores
<LI>Modified numeric fields to only accept numeric data in store engine
<LI>Added <a href=real_time_settings.asp class=link>2Checkout version 2</a> integration
<LI>Added 17 <a href=theme_manager.asp class=link>new templates</a>
<LI>Added <a href=shipping_class.asp class=link>shipping methods</a> defined by zip code
<LI>Added ability to define a order range for <a href=shipping_class5_list.asp class=link>percentage of total order shipping</a>
<LI>Added ability to <a href=inventory_display.asp class=link>arrange items/departments</a> 4 or 5 across
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>6/14/2004</b><BR><UL>
<LI>Updated WYSIWYG Editor Tool</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>6/10/2004</b><BR><UL>
<LI>Removed budget left field from modify customer account page unless budget left is greater then 0
<LI>Ordered department drop down boxes alphabetically
<LI>Moved individual page keywords, description, title to appear before generic keywords, description, title
<LI>Added message for search to indicate that no matching items where found
<LI>Added client side form checking to check for required data before submitting forms
<LI>Modified submit buttons to disable automatically upon submit to prevent double click
<LI>Added ability to remove certain shipping rates from <a href=realtime_options.asp class=link>realtime shipping options</a> (GOLD+)
<LI>Added ability to specify allowable countries for <a href=realtime_options.asp class=link>realtime shipping</a> (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>5/12/2004</b><BR><UL>
<LI>Added record sorting to list views, click on any column heading to sort on that column, click again to sort descending
<LI>Added recordset paging to list views, only 25 records will be shown per page, click next or previous to view additional records
<LI>Modified default customer registration country to reflect store country
<LI>Added confirm message for most delete links
<LI>Added checkboxes to delete multiple records at once for most list views
<LI>Added ability to view abandoned carts from <a href=orders.asp class=link>orders report</a>
<LI>Added different colored shading for active top menu selection
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>4/17/2004</b><BR><UL>
<LI>Added item and department name to product url and page title tag to increase search engine rankings
<LI>Added ability to choose <a href=shipping_class.asp class=link>multiple shipping method</a> calculation types simultaneously
<LI>Added ability to change timezone, exclude ip addresses and turn on/off dns lookup for <a href=store_statistics.asp class=link>statistics</a> (BRONZE+)
<LI>Added ability to include a <a href=item_layout.asp class=link>send to a friend</a> link in the item layouts (SILVER+)
<LI>Added ability to create multiple <a href=location_manager.asp class=link>shipping locations</a> and define where each item is shipped from for more accurate real-time shipping (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>4/7/2004</b><BR><UL>
<LI>Added ability to restrict <a href=security.asp class=link>access by login</a> in the admin interface
<LI>Added keyword identifiers to shipping emails
<LI>Added coupon code to order invoices
<LI>Modified view cart page to show items in the order added
<LI>Added Worldpay Recurring Billing Support (GOLD+)
<LI>Modified <a href=message_customers.asp class=link>mail merge</a> feature to send html emails (GOLD+)
<LI>Modified rewards dollars redemption (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>3/29/2004</b><BR><UL>
<LI>Modified login page to display new user registration link more prominently
<LI>Added ability to designate shipping should be charged taxes when using <a href=tax_list.asp class=link>taxes by zip code</a>
<LI>Added automatic image upload from inventory edit, department edit, etc
<LI>Added online <a href=support_list.asp class=link>support request</a> and tracking system
<LI>Added ability to protect certain pages and allow access to only certain customers (SILVER+)
<LI>Added the ability to <a href=coupon_manager.asp class=link>limit coupons</a> to designated skus (GOLD+)
<LI>Added capture of affiliate address information automatically on signup (GOLD+)
<LI>Added labels to the shopping cart, and receipt pages for user defined fields (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>2/28/2004</b><BR><UL>
<LI>Modified <a class=link href=reports_totals1.asp>traffic and total reports</a> shopper statistics, to not count search engine spiders and bots 
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>2/26/2004</b><BR><UL>
<LI>Added ability to upload item attributes (GOLD+)
<LI>Added ability <a href=export_items.asp class=link>export all item data</a> to a delimited file (GOLD+)
<LI>Added ability to order multiple attributes at the same time (GOLD+)
<LI>Removed returned orders from order report totals

</UL></td></tr>
<tr bgcolor='#FFFFFF'><TD><B>2/20/2004</b><BR><UL>
<LI>Added <a href=ftp.asp class=link>FTP Server</a> access (GOLD+)
<LI>Added ability to <a href=left_image_picker.asp class=link>upload images</a> and other files to different directories and view images by directory via web based interface (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>2/18/2004</b><BR><UL>
<LI>Added ability to allow customers to enter a <a href=misc_settings.asp class=link>gift message</a> on order (SILVER+)
<LI>Added card code image and popup help to secure payment page to help customer understand what should be placed in the card code field
<LI>Added popup error if customer attempts to put the same item into their cart more then once
<LI>Added ability to <a href=edit_items.asp class=link>mass edit or quick edit</a> sale price and sale dates for multiple items at the same time (GOLD+)
<LI>Added link at the top of every page to indicate a store has been scheduled for <a href=cancel_store.asp class=link>cancellation by store owner</a> and ability for store owner to <a href=uncancel_store.asp class=link>cancel the cancellation</a>.
<LI>Added ability to allow customer to <a href=edit_items.asp class=link>set their own price</a> for selected items (SILVER+)
<LI>Added ability to include a <a href=edit_items.asp class=link>5th user defined field</a> (GOLD+)
<LI>Now accepting Discover credit cards for <a href=update_cc.asp class=link>store monthly and yearly payments</a>.
<LI>Modified statistics package to use local server time for all reporting purposes (BRONZE+)
<LI>Modified <a href=my_customer_base.asp class=link>customer report</a> to hide customer passwords from view
<LI>Modified <a href=import_items.asp class=link>import items</a> to remove extraneous quotation marks added by some file generators (GOLD+)
<LI>Fixed <a href=emails.asp class=link>notification emails</a> and contact us to not modify emails with quotation marks (BRONZE+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>1/26/2004</b><BR><UL>
<LI>Modified dynamic links to look static to allow search engines to more easily index store pages.  Note the previous store owners can continue
to use the old links but may change to the new links to enhance a search engines ability to read their site.
<LI>Modified statistics tool, all statistics from 1/26/04 on will be captured using the new statistics tool
<LI>Added ability to keep user defined field values when changing attribute values
<LI>Added PRI realtime gateway integration
<LI>Added Yahoo Product Submit Data Feed
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>12/31/2003</b><BR><UL>
<LI>Added ability to order items in a manner other than alphabetical (SILVER+)
<LI>Added ability to <a href=upload_images.asp class=link>upload css,htm,zip,pdf</a> and other safe file extensions
<LI>Added ability to purchase <a href=custom_services.asp?Template=1 class=link>custom template creation</a> or <a href=custom_services.asp?Quick=1 class=link>quick start program</a> online from admin interface
<LI>Modified <a href=edit_items.asp class=link>item edit</a> screen to not allow use of multiple edit buttons if only 1 item is selected
<LI>Removed several seldom used pages from basic store package
<LI>Basic stores must now use default store layout.  Bronze + can choose from available templates.
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>12/21/2003</b><BR><UL>
<LI>Added 9 new <a href=theme_manager.asp class=link>design templates</a> (BRONZE+)
<LI>Added recurring billing support (SILVER+)
</UL></td></tr>


<tr bgcolor='#FFFFFF'><TD><B>12/14/2003</b><BR><UL>
<LI>Added <a href=yahoo_settings.asp class=link>Yahoo Product Submit</a> automatic daily data feed (GOLD+)
<LI>Added ability to download a file outside of Easystorecreator for ESD file download (GOLD+)
<LI>Added ability to <a href=left_image_picker.asp class=link>delete multiple images</a> at once
<LI>Added store id to management header so that customer is aware of their account number for support requests
<LI>Added new recommended international <a href=merchants.asp class=link>payment processors</a> and US small business payment provider
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>12/1/2003</b><BR><UL>
<LI>Added popup message if shopper tries to process order a second time to stop the second transaction.	Most payment gateways have internal checking for this, some of our newer gateways do not.
<LI>Removed description from merchant invoice order details
<LI>Added billing information fields to order export
<LI>Modified eCheck field Bank ABA code to read Bank Routing Number and added description
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>11/25/2003</b><BR><UL>
<LI>Added <a href=real_time_settings.asp class=link>Protx gateway realtime</a> credit card integration
<LI>Added website monitoring link so our merchants can see the <a href="http://florida.websitepulse.net/uptime/M173I6840.html" target="_blank" class=link>current uptime stats of their website</a>, bottom right
<LI>Added <a href=reports_coupons.asp class=link>Coupon Report</a> (GOLD+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>11/15/2003</b><BR><UL>
<LI>Added auto email to stores with overdue payments and auto store closure after 7 days of failed payments
<LI>Removed summary statistics from homepage, they can now be found <a href=reports_totals1.asp class=link>here</a>
<LI>Added online store <a href=features2.asp class=link>feature request</a>
<LI>Added <a href=my_account.asp class=link>My Account</a> menu group
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>11/12/2003</b><BR><UL>
<LI>Added ability to set meta tags for each individual page, item or department	(SILVER+)
<LI>Moved checkout without registering and registration links from the bottom to the top of the login page
</UL></td></tr>

<tr bgcolor='#FFFFFF'><TD><B>11/11/2003</b><BR><UL>
<LI>Moved servers to Rackspace Managed Hosting
<LI>Upgraded server hardware
<LI>Added web mail beta version, try it out and tell us what you think (BRONZE+)
<LI>Added ability to view the contents of the carts contents on a missing order and complete that order if necessary from order details

</UL></td></tr>


<tr bgcolor='#FFFFFF'><TD><B>11/4/2003</b><BR><UL>
<LI>Added online distributed processing of orders which will result in quicker orders when servers are busy
<LI>Modified orders breakdown, traffic totals and homepage statistics to reflect calendar day statistics instead of 24 hour statistics
<LI>Added progress bar to all upload pages except image upload
<LI>Enhanced upload pages to work more quickly and experience less timeouts for large file uploads > 2 MB
<LI>Added ability to modify credit card billing information for your Easystorecreator account online

</UL></td></tr>


<tr bgcolor='#FFFFFF'><TD><B>10/20/2003</b><BR><UL>
<LI>Added online invoicing system to view Easystorecreator invoices
<LI>Added online cancellation of service
<LI>Added coupon import utility (GOLD+)
<LI>Modified store payment method to not have a default to prevent customers from continuing without choosing a valid payment method
<LI>Added ability to print purely textual customer invoices from order details screen
<LI>Added direct link to order details in admin order notification email
<LI>Added ability to add on a flat handling fee to all orders in store (SILVER+)
<LI>Added Bank of America and WorldPay Integration
<LI>Removed username and password requirement for using USPS realtime shipping (SILVER+)
<LI>Added instructions for aquiring a UPS realtime shipping access license, username and password (GOLD+)
<LI>Added online automatic Easystorecreator Service upgrade
<LI>Modified Easystorecreator menus to show grey for pages which are not allowed at the current service level
</UL></td></tr>

<tr bgcolor='#FFFFFF'><td><B>9/25/2003</B><BR><UL><LI>Added ability to remove top department and item browsing area (SILVER+)
<LI>Added department next, previous, and jump to letters at the bottom as well as at the top of department browsing where required (GOLD+)
<LI>Added free payment option which is the default choice if total = 0
<LI>Modified item and department browsing to result in faster load times for customers
<LI>Added customer import from text file feature (GOLD+)
<LI>Added ability to include custom text at either the top or the bottom of each page instead of only the top (SILVER+)
<LI>Modified administrative interface for faster page loads</UL></td></tr>

<tr bgcolor='#FFFFFF'><td><B>8/26/2003</B><BR><UL><LI>Added option to remove A-Z jump letters (SILVER+)
<LI>Added option to hide PO Number Box (SILVER+)
<LI>Added option to hide departments with no items (SILVER+)
<LI>Added option to turn off statistics capture for faster page load times (SILVER+)</UL></td></tr>

<tr bgcolor='#FFFFFF'><td><B>8/25/2003</B><BR><UL><LI>Allow customer group filtering by previous departmental purchase (SILVER+)
<LI>Added printable invoice link on receipt page
<LI>Added ability to order departments (BRONZE+)
</UL></td></tr>

<tr bgcolor='#FFFFFF'><td><B>8/17/2003</B><BR><UL><LI>Allow merchant to turn on/off department listing (BRONZE+)
<LI>Now accepting American Express and Paypal payments for Easystorecreator service upgrade
<LI>Added ability to print shipping labels
<LI>Added auto return authorization (SILVER+)</UL></td></tr>

  <tr bgcolor='#FFFFFF'><td><B>8/10/2003</B><BR><UL><LI>Added feature to allow customers to use a previously charged credit card on return to merchants store.
  <LI>Modified ship to multiple locations to allow shopper to ship the same item purchased more than once to separate locations. (SILVER+)</UL></td></tr>

	  <tr bgcolor='#FFFFFF'><td><B>8/3/2003</B><BR><UL><LI>Added rewards program (GOLD+)
	  <LI>Allow auto capture of funds on shipment for Authorize.net</UL></td></tr>

	  <tr bgcolor='#FFFFFF'><td><B>7/19/2003</B><BR><UL><LI>Allow admin to change text on action buttons to say anything in addition to choosing images.
	  <LI>Added newsletter signup, cancel and management functions (SILVER+)</UL></td></tr>

	  <tr bgcolor='#FFFFFF'><td><B>7/17/2003</B><BR><UL><LI>Implemented message queueing mechanism to speed up pages which send notification emails.</UL></td></tr>
	  <tr bgcolor='#FFFFFF'><td><B>7/9/2003</B><BR><UL><LI>Added real-time shipping (GOLD+)
	  <LI>Increased load speed for stores</UL></td></tr>
<tr bgcolor='#FFFFFF'><td><B>6/24/2003</B><BR><UL><LI>Added feature to allow shopper to easily return to the page they were currently
	  on after adding to cart.
	  <LI>Fixed problem with show and hide cart on checkout</UL></td></tr>

	<tr bgcolor='#FFFFFF'><td><B>6/21/2003</B><BR><UL><LI>Added feature to disallow multiple shipments on checkout
	  <LI>Added feature to allow specifying order of attributes display besides alphabetical</UL></td></tr>
	<tr bgcolor='#FFFFFF'><td><B>6/19/2003</B><BR><UL><LI>Added support for taxes by country
	<LI>Added support for hiding the login screen on shopper checkout</UL></td></tr>
	<tr bgcolor='#FFFFFF'><td><B>6/17/2003</B><BR><UL><LI>Added support for hiding an items price which will make it so a shopper cannot
	add this item to their cart and must call for pricing.
	<LI>Added support to Electronic Transfer Gateway</UL></td></tr>
	<tr bgcolor='#FFFFFF'><td><B>6/16/2003</B><BR><UL><LI>Added support for non-numeric zip codes
	<LI>Added integration of new payment gateways PSI Gate and BluePay
	<LI>Added state field and included canadian states
	<LI>Added support for taxes by state in addition to taxes by zip with option to tax shipping</UL></td></tr>
	<tr bgcolor='#FFFFFF'><td><B>6/06/2003</B><BR><UL><LI>Added more features to wizard.
	</UL></td></tr>
	<tr bgcolor='#FFFFFF'><td><B>5/23/2003</B><BR><UL><LI>New look for the administrative interface.</UL></td></tr>

	<tr bgcolor='#FFFFFF'><td><B>5/18/2003</B><BR><UL><LI>Added ability to accept Paypal in addition to using another merchant account provider.
	<LI>5 Department and 5 Item Templates available to choose from to customize the look of your store or design your own layout.</UL></td></tr> 
		<tr bgcolor='#FFFFFF'><td><B>5/04/2003</B><BR><UL><LI>Enhanced Security for Real-time settings, now stored in encrypted format and changes
		are not allowed unless logged in securely.</UL></td></tr>
		<tr bgcolor='#FFFFFF'><td><B>4/29/2003</B><BR><UL><LI>Logging of AVS and Card Code results from credit card transactions
		<LI>Verisign Payflow Pro Integration</UL></td></tr>
		<tr bgcolor='#FFFFFF'><td><B>4/23/2003</B><BR><UL><LI>Authorize.Net eCheck Support
		</UL></td></tr>

		<tr bgcolor='#FFFFFF'><td><B>4/21/2003</B><BR><UL><LI>Enhanced Site Statistics</UL></td></tr>

		<tr bgcolor='#FFFFFF'><td><B>4/20/2003</B><BR><UL><LI>Froogle.com Shopping Portal integration
		<LI>24 New Themes added
		<LI>Links to signup for both Paypal and Authorize.Net added
		<LI>Fixed error with quotes in store name
		<LI>Added check to see if dns are pointing correctly for secondary domain
		<LI>Added admin email utility to send email directly to easystorecreator admins
		</UL>
		</td></tr>
		<tr bgcolor='#FFFFFF'><td><B>4/14/2003</B><BR><UL>
		<LI>Fixed error when choosing text font
		</UL></td></tr>
		<tr bgcolor='#FFFFFF'><td><B>4/06/2003</B><UL><LI>Added web email</UL></td></tr>

<% createFoot thisRedirect, 0%>
