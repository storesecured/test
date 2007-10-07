<%
sStyleGeneral=""
if sMenu="general" then
        sStyleGeneral="1"
end if 
sStyleShipping=""
if sMenu="shipping" then
        sStyleShipping="1"
end if
sStyleDesign=""
if sMenu="design" then
        sStyleDesign="1"
end if
sStyleMarketing=""
if sMenu="marketing" then
        sStyleMarketing="1"
end if
sStyleInventory=""
if sMenu="inventory" then
        sStyleInventory="1"
end if
sStyleStatistics=""
if sMenu="statistics" then
        sStyleStatistics="1"
end if
sStyleCustomers=""
if sMenu="customers" then
        sStyleCustomers="1"
end if
sStyleOrders=""
if sMenu="orders" then
        sStyleOrders="1"
end if
sStyleMyAccount=""
if sMenu="account" then
        sStyleMyAccount="1"
end if
%>

<script type="text/javascript" language="JavaScript1.2">
var isHorizontal=1;

var pressedItem = -1;
var key="173b2119ei6tg";
var key1="165b1680ei6tg";

var blankImage="images/spacer.gif";
var fontStyle="normal 9pt Arial";
var fontColor=["#FFFFFF","#FFFFFF"];
var fontDecoration=["none","underline"];

var itemBackColor=["#007FBA","#01324A"];
var itemBorderWidth=0;
var itemAlign="left";
var itemBorderColor=["#6655ff","#665500"];
var itemBorderStyle=["solid","solid"];
var itemBackImage=["",""];

var menuBackImage="";
var menuBackColor="#007FBA";
var menuBorderColor="#FFFFFF";
var menuBorderStyle="solid";
var menuBorderWidth=1;
var transparency=100;
var transition=0;
var transDuration=0;
var shadowColor="#999999";
var shadowLen=0;
var menuWidth="100%";

var itemCursor="hand";
var itemTarget="_self";
var statusString="text";
var subMenuAlign = "left";
var iconTopWidth  = 0;
var iconTopHeight = 0;
var iconWidth=0;
var iconHeight=0;
var arrowImageMain=["images/arrow_d.gif","images/arrow_d-over.gif"];
var arrowImageSub=["images/arrow_r.gif","images/arrow_r-over.gif"];
var arrowWidth=0;
var arrowHeight=0;
var itemSpacing=0;
var itemPadding=4;

var separatorImage="";
var separatorWidth="0";
var separatorHeight="0";
var separatorAlignment="center";

var separatorVImage="";
var separatorVWidth="0";
var separatorVHeight="0";

var moveCursor = "move";
var movable = 0;
var absolutePos = 0;
var posX = 20;
var posY = 100;

var floatable=1;
var floatIterations=8;

var itemStyles =
[
  ["itemBackColor=#dddddd,#bbccee", "itemBorderWidth=1","itemBorderColor=#dddddd,#316AC5"],
  ["itemBackColor=#B11F0F,#000000"]
];

var menuStyles =
[
  ["menuBorderWidth=1","menuBackColor=#ffffff"]
];

var menuItems =
[
	["-"],
	["General",,,,,,<%=sStyleGeneral %>],
		["|Settings","company.asp",,,,,<%=sStyleGeneral %>],
		["|Payments","payment_manager.asp",,,,,<%=sStyleGeneral %>],
		["||Methods","payment_manager.asp",,,,,<%=sStyleGeneral %>],
		["|||View/Edit","payment_manager.asp",,,,,<%=sStyleGeneral %>],
		["|||Add","payment_edit.asp",,,,,<%=sStyleGeneral %>],
		["||Gateway","real_time_settings.asp",,,,,<%=sStyleGeneral %>],
		["||Fraud Control","maxmind_settings.asp",,,,,<%=sStyleGeneral %>],
		["||Accept Credit Cards","merchantacct.asp",,,,,<%=sStyleGeneral %>],
		["|Return Policy","returns.asp",,,,,<%=sStyleGeneral %>],
		["|Privacy Policy","privacy.asp",,,,,<%=sStyleGeneral %>],
		["|Homepage","homepage.asp",,,,,<%=sStyleGeneral %>],
		["|Taxes","tax_list.asp",,,,,<%=sStyleGeneral %>],
			["||View/Edit","tax_list.asp",,,,,<%=sStyleGeneral %>],
			["||Add State Tax","tax_add_state.asp",,,,,<%=sStyleGeneral %>],
			["||Add Zipcode Tax","tax_add.asp",,,,,<%=sStyleGeneral %>],
			["||Add Country Tax","tax_add_country.asp",,,,,<%=sStyleGeneral %>],
			["||Import Taxes by Zip","import_zips.asp",,,,,<%=sStyleGeneral %>],
		["|Email","message_list.asp",,,,,<%=sStyleGeneral %>],
			["||Setup Instructions","message_list.asp",,,,,<%=sStyleGeneral %>],
			["||Notifications","emails.asp",,,,,<%=sStyleGeneral %>],
			["||WebMail","http://mail182.easystorecreator.com:8001","","","_blank",,<%=sStyleGeneral %>],
		["|Domain Name","domain.asp",,,,,<%=sStyleGeneral %>],
		["|FTP","ftp.asp",,,,,<%=sStyleGeneral %>],
		["|Open/Close Store","activation.asp",,,,,<%=sStyleGeneral %>],
		["|Custom Fields","fields.asp",,,,,<%=sStyleGeneral %>],
			["||View/Edit","fields.asp",,,,,<%=sStyleGeneral %>],
			["||Add","custom_fields.asp",,,,,<%=sStyleGeneral %>],
		["|Id Numbers","store_manag.asp",,,,,<%=sStyleGeneral %>],
	["Shipping",,,,,,<%=sStyleShipping %>],
			["|Settings","shipping_class_realtime.asp",,,,,<%=sStyleShipping %>],
			["|Methods","shipping_all_list.asp",,,,,<%=sStyleShipping %>],
			["||View/Edit","shipping_all_list.asp",,,,,<%=sStyleShipping %>],
			["||Add","shipping_all_class.asp",,,,,<%=sStyleShipping %>],
			["||Import","import_shipping.asp",,,,,<%=sStyleShipping %>],
			["||Export","export_shipping.asp",,,,,<%=sStyleShipping %>],
			["|Locations","location_manager.asp",,,,,<%=sStyleShipping %>],
				["||View/Edit","location_manager.asp",,,,,<%=sStyleShipping %>],
				["||Add","new_location.asp",,,,,<%=sStyleShipping %>],
			["|Labels","shipping_labels.asp",,,,,<%=sStyleShipping %>],
			["|Insurance","insurance_class.asp",,,,,<%=sStyleShipping %>],
				["||View/Edit Flat Fee","insurance_class1_list.asp",,,,,<%=sStyleShipping %>],
				["||View/Edit Total Order Matrix","insurance_class4_list.asp",,,,,<%=sStyleShipping %>],
				["||View/Edit % Total Order","insurance_class5_list.asp",,,,,<%=sStyleShipping %>],
				["||Add Flat Fee","insurance_class1_add.asp",,,,,<%=sStyleShipping %>],
				["||Add Total Order Matrix","insurance_class4_add.asp",,,,,<%=sStyleShipping %>],
				["||Add % Total Order","insurance_class5_add.asp",,,,,<%=sStyleShipping %>],
			["|Allowable Countries","countries_select.asp",,,,,<%=sStyleShipping %>],


	["Design",,,,,,<%=sStyleDesign %>],
		["|Pages","page_manager.asp",,,,,<%=sStyleDesign %>],
			["||View/Edit","page_manager.asp",,,,,<%=sStyleDesign %>],
			["||Add","new_page.asp",,,,,<%=sStyleDesign %>],
		["|Links","page_links_manager.asp",,,,,<%=sStyleDesign %>],
			["||View/Edit","page_links_manager.asp",,,,,<%=sStyleDesign %>],
			["||Add","new_page_link.asp",,,,,<%=sStyleDesign %>],
			["||List All","links.asp",,,,,<%=sStyleDesign %>],
		["|Forms","fields_page.asp",,,,,<%=sStyleDesign %>],
			["||View/Edit","fields_page.asp",,,,,<%=sStyleDesign %>],
			["||Add","custom_fields_page.asp",,,,,<%=sStyleDesign %>],
		["|Template","template_list.asp",,,,,<%=sStyleDesign %>],
			["||Choose Existing","new_theme_manager.asp",,,,,<%=sStyleDesign %>],
			["||View/Edit","template_list.asp",,,,,<%=sStyleDesign %>],
			["||Add","custom_template.asp",,,,,<%=sStyleDesign %>],
		["|Customize Invoice","customize_invoice.asp",,,,,<%=sStyleDesign %>],
		["|Files/Images","",,,,,<%=sStyleDesign %>],
			["||List/Preview","left_image_picker.asp",,,,,<%=sStyleDesign %>],
			["||Upload","upload_images.asp",,,,,<%=sStyleDesign %>],
		
		
	
	["Marketing",,,,,,<%=sStyleMarketing %>],

                ["|Search Engines","add_url.asp",,,,,<%=sStyleMarketing %>],
		["|Newsletter","newsletter_manager.asp",,,,,<%=sStyleMarketing %>],
			["||Subscribers","newsletter_manager.asp",,,,,<%=sStyleMarketing %>],
			["|||View/Edit","newsletter_manager.asp",,,,,<%=sStyleMarketing %>],
			["|||Add","new_newsletter.asp",,,,,<%=sStyleMarketing %>],
			["|||Import","import_newsletter.asp",,,,,<%=sStyleMarketing %>],
			["|||Export","export_newsletter.asp",,,,,<%=sStyleMarketing %>],
			["||Send","message_customers.asp",,,,,<%=sStyleMarketing %>],
		["|Banners","banners.asp",,,,,<%=sStyleMarketing %>],
			["||View/Edit","banners.asp",,,,,<%=sStyleMarketing %>],
			["||Add","banners_add.asp",,,,,<%=sStyleMarketing %>],
		["|Coupons","coupon_manager.asp",,,,,<%=sStyleMarketing %>],
			["||View/Edit","coupon_manager.asp",,,,,<%=sStyleMarketing %>],
			["||Add","new_coupon.asp",,,,,<%=sStyleMarketing %>],
			["||Settings","coupon_settings.asp",,,,,<%=sStyleMarketing %>],
			["||Import","import_coupons.asp",,,,,<%=sStyleMarketing %>],
			["||Report","reports_coupons.asp",,,,,<%=sStyleMarketing %>],
		["|Promotions","promotion_manager.asp",,,,,<%=sStyleMarketing %>],
			["||View/Edit","promotion_manager.asp",,,,,<%=sStyleMarketing %>],
			["||Add","new_promotion.asp",,,,,<%=sStyleMarketing %>],
		["|Gift Certificates","gift_manager.asp",,,,,<%=sStyleMarketing %>],
			["||View/Edit","gift_manager.asp",,,,,<%=sStyleMarketing %>],
			["||Add","new_gift.asp",,,,,<%=sStyleMarketing %>],
		["|Affiliates","affiliates_manager.asp",,,,,<%=sStyleMarketing %>],
			["||View/Edit","affiliates_manager.asp",,,,,<%=sStyleMarketing %>],
			["||Add","new_affiliate.asp",,,,,<%=sStyleMarketing %>],
			["||Banners","",,,,,<%=sStyleMarketing %>],
				["|||View/Edit","banners_list.asp",,,,,<%=sStyleMarketing %>],
				["|||Add","affiliate_banner.asp",,,,,<%=sStyleMarketing %>],
			["||Settings","affiliate_settings.asp",,,,,<%=sStyleMarketing %>],
			["||Report","affiliates_reports.asp",,,,,<%=sStyleMarketing %>],
		["|Google Base","froogle_settings.asp",,,,,<%=sStyleMarketing %>],
		["|Rewards","reward_settings.asp",,,,,<%=sStyleMarketing %>],


	["Inventory","",,,,,<%=sStyleInventory %>],
		["|Departments","department_manager.asp",,,,,<%=sStyleInventory %>],
			["||View/Edit","department_manager.asp",,,,,<%=sStyleInventory %>],
			["||Add","store_dept_basic.asp",,,,,<%=sStyleInventory %>],
			["||Advanced Add","store_dept_add.asp",,,,,<%=sStyleInventory %>],
			["||Layout","layout.asp",,,,,<%=sStyleInventory %>],
			["||Import","import_department.asp",,,,,<%=sStyleInventory %>],
			["||Export","export_departments.asp",,,,,<%=sStyleInventory %>],
		["|Items","edit_items.asp",,,,,<%=sStyleInventory %>],
			["||Search","edit_items.asp",,,,,<%=sStyleInventory %>],
			["||Add","item_basic_edit.asp",,,,,<%=sStyleInventory %>],
			["||Advanced Add","item_edit.asp",,,,,<%=sStyleInventory %>],
			["||ESD Files","upload_files.asp",,,,,<%=sStyleInventory %>],
				["|||View","file_list.asp",,,,,<%=sStyleInventory %>],
				["|||Upload","upload_files.asp",,,,,<%=sStyleInventory %>],
			["||Pins","edit_pin1.asp",,,,,<%=sStyleInventory %>],
				["|||View/Edit","edit_pin1.asp",,,,,<%=sStyleInventory %>],
				["|||Add","pin_add.asp",,,,,<%=sStyleInventory %>],
				["|||Import","import_pin.asp",,,,,<%=sStyleInventory %>],
			["||Layout","item_layout.asp",,,,,<%=sStyleInventory %>],
			["||Import","import_items.asp",,,,,<%=sStyleInventory %>],
		["|Settings","item_settings.asp",,,,,<%=sStyleInventory %>],
		["|Report","reports_totals3.asp",,,,,<%=sStyleInventory %>],
		["|Best Sellers","reports_best_sellers.asp",,,,,<%=sStyleInventory %>],
		
	["Statistics","",,,,,<%=sStyleStatistics %>],
			["|Report","statistics.asp","","","_blank","0",<%=sStyleStatistics %>],
			["|Settings","store_statistics.asp",,,,,<%=sStyleStatistics %>],
			
	["Customers","",,,,,<%=sStyleCustomers %>],
			["|Search","my_customer_base.asp",,,,,<%=sStyleCustomers %>],
			["|Add","add_new_customer.asp",,,,,<%=sStyleCustomers %>],
			["|Groups","customers_groups.asp",,,,,<%=sStyleCustomers %>],
				["||View/Edit","customers_groups.asp",,,,,<%=sStyleCustomers %>],
				["||Add","create_new_customer_group.asp",,,,,<%=sStyleCustomers %>],
			["|Import","import_customers.asp",,,,,<%=sStyleCustomers %>],
			["|Budget","",,,,,<%=sStyleCustomers %>],
				["||Import","import_budget.asp",,,,,<%=sStyleCustomers %>],
		
	["Orders","",,,,,<%=sStyleOrders %>],
		["|Search","orders.asp",,,,,<%=sStyleOrders %>],
		["|Alternate View","reports_order_break_down.asp",,,,,<%=sStyleOrders %>],
		["|Report","reports_totals1.asp",,,,,<%=sStyleOrders %>],
		
	["My Account","",,,,,<%=sStyleMyAccount %>],
		["|Payments","my_payments.asp",,,,,<%=sStyleMyAccount %>],
			["||Prior Invoices","my_payments.asp",,,,,<%=sStyleMyAccount %>],
			["||Upgrade Service","billing.asp",,,,,<%=sStyleMyAccount %>],
			["||Update Payment Method","update_cc.asp",,,,,<%=sStyleMyAccount %>],
		["|Support Requests","support_list.asp",,,,,<%=sStyleMyAccount %>],
			["||View/Edit","support_list.asp",,,,,<%=sStyleMyAccount %>],
			["||Add","support_request.asp",,,,,<%=sStyleMyAccount %>],
		["|Feature Requests","features2.asp",,,,,<%=sStyleMyAccount %>],
			["||View","features2.asp",,,,,<%=sStyleMyAccount %>],
			["||Add","feature_request.asp",,,,,<%=sStyleMyAccount %>],
		["|Cancel Site","cancel_store.asp",,,,,<%=sStyleMyAccount %>],
		["|Space Usage","space_usage.asp",,,,,<%=sStyleMyAccount %>],
		["|Admin Logins","security.asp",,,,,<%=sStyleMyAccount %>],
			["||View/Edit","security.asp",,,,,<%=sStyleMyAccount %>],
			["||Add","security_add.asp",,,,,<%=sStyleMyAccount %>],
			["||Logs","access_manager.asp",,,,,<%=sStyleMyAccount %>],
		["|Download Exported Files","download_exported_files.asp",,,,,<%=sStyleMyAccount %>],
		["|Refer Friends","referral.asp",,,,,<%=sStyleMyAccount %>],
		["|Other Resources","resources.asp",,,,,<%=sStyleMyAccount %>],
];
apy_init();
</script>