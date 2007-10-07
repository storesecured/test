<%@ language=vbscript.encode%>
<!--#include file="local_config.asp"-->
<!--#include file="get_paths.asp"-->
<!--#include file="common_functions.asp"-->
<%

Set conn_store = Server.CreateObject("ADODB.connection")
Store_Database="Ms_Sql"
conn_store.Provider = "sqloledb.1"

db_Driver = "SQL Server" 'DB Server Driver Type (Now, only SQL Server)

db_ConnectionString = "Provider=sqloledb;Data Source=" & db_Server & ";UID=" & db_UIN & ";PWD=" & db_pwd & ";Initial Catalog=" & db_Database & ";Network Library=DBMSSOCN"
conn_store.Open db_ConnectionString

Set rs_store = Server.CreateObject("ADODB.Recordset")

%>

