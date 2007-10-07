<%
' You shouldn't ever have to modify this file '

' ---- These constants are used by the UserFlags property

Const UF_SCRIPT                      = &H0001
Const UF_ACCOUNTDISABLE              = &H0002
Const UF_HOMEDIR_REQUIRED            = &H0008
Const UF_LOCKOUT                     = &H0010
Const UF_PASSWD_NOTREQD              = &H0020
Const UF_PASSWD_CANT_CHANGE          = &H0040
Const UF_TEMP_DUPLICATE_ACCOUNT      = &H0100
Const UF_NORMAL_ACCOUNT              = &H0200
Const UF_INTERDOMAIN_TRUST_ACCOUNT   = &H0800
Const UF_WORKSTATION_TRUST_ACCOUNT   = &H1000
Const UF_SERVER_TRUST_ACCOUNT        = &H2000
Const UF_DONT_EXPIRE_PASSWD          = &H10000
Const UF_MNS_LOGON_ACCOUNT           = &H20000

' ---- These constants are used by Services

Const ADS_SERVICE_ERROR_IGNORE       = &H0000
Const ADS_SERVICE_ERROR_NORMAL       = &H0001
Const ADS_SERVICE_ERROR_SEVERE       = &H0002
Const ADS_SERVICE_ERROR_CRITICAL     = &H0003

Const ADS_SERVICE_KERNAL_DRIVER      = &H0001
Const ADS_SERVICE_FILE_SYSTEM_DRIVER = &H0002
Const ADS_SERVICE_OWN_PROCESS        = &H0010
Const ADS_SERVICE_SHARE_PROCESS      = &H0020

Const ADS_SERVICE_BOOT_START         = &H0000
Const ADS_SERVICE_SYSTEM_START       = &H0001
Const ADS_SERVICE_AUTO_START         = &H0002
Const ADS_SERVICE_DEMAND_START       = &H0003
Const ADS_SERVICE_SERVICE_DISABLED   = &H0004

Const ADS_SERVICE_STARTED            = &H0001
Const ADS_SERVICE_START_PENDING      = &H0002
Const ADS_SERVICE_STOP_PENDING       = &H0003
Const ADS_SERVICE_RUNNING            = &H0004
Const ADS_SERVICE_CONTINUE_PENDING   = &H0005
Const ADS_SERVICE_PAUSE_PENDING      = &H0006
Const ADS_SERVICE_PAUSED             = &H0007
Const ADS_SERVICE_ERROR              = &H0008

' ---- These constants are used by PrintJob

Const ADS_JOB_OK                     = &H0000
Const ADS_JOB_PAUSED                 = &H0001
Const ADS_JOB_ERROR                  = &H0002
Const ADS_JOB_DELETING               = &H0004
Const ADS_JOB_PRINTING               = &H0010
Const ADS_JOB_OFFLINE                = &H0020
Const ADS_JOB_PAPEROUT               = &H0040
Const ADS_JOB_PRINTED                = &H0080
Const ADS_JOB_DELETED                = &H0100

' ---- These constants define site status

Const IIS_SITE_RUNNING               = &H00000002
Const IIS_SITE_STOPPED               = &H00000004
Const IIS_SITE_PAUSED                = &H00000006

' ---- These constants define isolation levels

Const IIS_SITE_IS_INPROC             = 0
Const IIS_SITE_IS_OUTPROC            = 1
Const IIS_SITE_IS_POOLED             = 2
Const INPROC                         = 0 
Const OUTPROC                        = 1 
Const POOLED                         = 2 

' ---- Miscellaneous constants

Const ERR_IIS_SITE_DOESNT_EXIST      = &H80070003

%>