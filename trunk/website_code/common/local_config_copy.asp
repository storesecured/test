<%

response.clear

db_Server = "127.0.0.1\OFFICE" 'DB Server IP or Alias
db_UIN		=	"administrator"	 'DB user name
db_pwd		=	"easystore"  'DB user password

sDebugIP="127.0.0.1"
sDebugStore=4991
iDebugOn=0

sLocalAddName = "-localdev"
fServerVersion = false
sComputerName = "EASYSTOREOFFICE"

root = "storesecured.com"
root_w_dot = "."&root

hackerfree_url="manage" & root_w_dot
site_host = root
Name = "StoreSecured"
ftp_server = "ftp" & root_w_dot
ftp_port = "21"
mail_server = "mail" & root_w_dot
ssl_oid="00106851"
manage_server = "manage" & root_w_dot

sSales_email = "sales@"&root
sSupport_email = "support@"&root
sNoReply_email = "noreply@"&root
sDeveloper_email = "melanie@"&root
sReport_email = "melanie@"&root
sPaypal_email = "paypal@"&root
sFollow_email = "adam@"&root

sInstalledDrive="c"
Application_Path = sInstalledDrive&":\storesecured_checkout\"
Mail_Path = sInstalledDrive&":\smartermail\domains\"
FTP_User_File = sInstalledDrive&":\program files\bpftp server\users.ini"

sEmergencyMail="8588374707@mobile.att.net"

Server_Ip="69.20.120.19"
Server_Ip2="192.168.1.19"

manage_server = fn_url_rewrite(manage_server)

%>